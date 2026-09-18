import csv
import os
import re
import unicodedata
from datetime import datetime, timezone
from urllib.parse import urlparse
import feedparser
import pandas as pd
import requests
from bs4 import BeautifulSoup

OUTPUT_FILE = "output/rss_history.csv"
LOG_FILE = "output/log.txt"

# ============================================================
# 1. Configuration des Sources (11 Flux RSS + HTML Frandroid)
# ============================================================

RSS_FEEDS = {
    "Frandroid - RSS": "https://www.frandroid.com/feed",
    "Phonandroid": "https://www.phonandroid.com/feed",
    "01net": "https://www.01net.com/feed/",
    "Numerama": "https://www.numerama.com/feed/",
    "CCM": "https://www.commentcamarche.net/rss/?output=xml",
    "KultureGeek": "http://feeds.feedburner.com/Kulturegeek",
    "JournalDuGeek": "https://www.journaldugeek.com/feed/",
    "LesNumeriques": "https://www.lesnumeriques.com/rss.xml",
    "Korben": "https://korben.info/feed",
    "NextInpact": "https://www.nextinpact.com/rss",
    "TomsHardware": "https://www.tomshardware.fr/feed/",
    "Clubic": "https://www.clubic.com/feed/",
}

FRANDROID_SECTIONS = {
    "Frandroid - Actualités": "https://www.frandroid.com/actualites",
    "Frandroid - Tests": "https://www.frandroid.com/test",
    "Frandroid - Bons plans": "https://www.frandroid.com/bon-plan",
}

# ============================================================
# 2. Définition des Mots-Clés et Catégories (19 Catégories)
# ============================================================

CATEGORIES = {
    "SMARTPHONE": [
        "smartphone", "téléphone", "phone", "l'iphone", "iphone", "pixel", "flip", "siri", "forfait", 
        "android", "GrapheneOS", "ios", "5G", "6G", "réseau", "Esim", "Wifi", "wi-fi", "routeur", 
        "netgear", "ethernet", "traceur", "Tag", "lunette", "appli", "application", "anker", "redmi", 
        "whatsapp", "spacex", "strarlink", "sosh", "Honor", "One Ui", "poco", "realme", "Galaxy S"
    ],
    "PC": [
        "PC", "Geekom", "Framework", "SSD", "serveur", "Nas", "portable", "macbook", "asus", "lenovo", 
        "HP", "Dell", "acer", "MSI ryzen", "qualcomm", "Media Tek", "snapdragon", "amd", "nvdia", 
        "puce graphique", "Galaxy book"
    ],
    "PERIPHERIQUE": [
        "Clavier", "écran", "moniteur", "souris", "l'imprimante", "imprimante", "logitech", "corsair", 
        "steam deck", "chargeur", "charge", "batterie externe", "ugreen", "branche", "lexar", 
        "périphérique", "hub", "dock", "usb", "usb-c", "Ram", "chaise"
    ],
    "TABLETTE": [
        "Tablette", "Fold", "pad", "tab", "kindle", "remarkable", "liseuse", "kobo", "ebooks", "pliable", 
        "pliure", "pli", "Galaxy tab"
    ],
    "LOGICIELS_OS": [
        "Logiciels", "sofware", "windows", "macos", "power Toys", "office", "windows", "linux", "Bios", 
        "OS", "mail", "rss", "slack", "teams", "capture", "outlook", "onenote", "gmail", "drive", 
        "cleaner", "VPN", "traduction", "chat", "automatisation", "edge", "chrome", "safari", "firefox", 
        "mozilla", "opera", "vivaldi", "moteur de recherche", "github", "zip", "terminal", "torrent", 
        "chatbot", "interface", "wordpress", "web", "navigateur"
    ],
    "IA": [
        "l'IA", "IA", "AI", "intelligence artificielle", "data center", "LLM", "token", "anthropic", 
        "d'openAI", "openAI", "chatgpt", "gemini", "meta", "Claude", "openclaw", "perplexity", "prompt", 
        "copilot", "gemini", "deepseek", "mistral", "groq", "midjourney"
    ],
    "VIDEO": [
        "Iptv", "vids", "télévision", "téléviseur", "TV", "omni", "pics", "vlc", "video", "éditeur", 
        "3D", "plex", "téléviseur", "LG", "hisense", "TCL", "Hdmi", "RGB", "videoprojecteur", "Awol", 
        "xgimi", "Jmgo", "stick", "DJI", "osmo", "Insta360", "cast", "graphique", "édition", "image", 
        "gopro", "affichage"
    ],
    "AUDIO": [
        "casque", "audio", "casque audio", "casque gaming", "pods", "buds", "headphone", "jbl", 
        "shockz", "clip", "écouteurs", "réduction de bruit", "enceinte", "barre de son", "harman", 
        "sennheiser", "bose", "sonos", "égaliseur", "musique"
    ],
    "MOBILITE_DOUCE": [
        "mobilité", "vélo", "trottinette", "cargo", "sacoche", "antivol", "VAE", "navigo", "blablacar", 
        "taxi", "rer", "transport en commun", "ebike", "vtt", "segway", "veste"
    ],
    "VOITURE": [
        "voiture", "automobile", "conduite", "conduire", "rouler", "essence", "carburant", "monospace", 
        "berline", "citadine", "SUV", "permis", "volkswagen", "mercedes", "bmw", "byd", "mg", "renault", 
        "peugeot", "toyota", "nissan", "stellantis", "citroen", "dacia", "kia", "fiat", "tesla", "ford", 
        "jeep", "twingo", "autonome", "FSD", "waze", "google maps", "auto", "carplay", "taxe", "moto", 
        "scooter", "hybride", "voiture électrique", "borne"
    ],
    "MENAGE": [
        "aspirateur", "robot aspirateur", "dyson", "roborock", "narwal", "mova", "ecovacs", "dreame", 
        "tineco", "shark", "laveur", "lavage", "machine à laver", "vitres", "pressing", "défroisseur", 
        "centrale vapeur", "fer à repasser", "tambour", "maison", "linge", "vaisselle", "lave-vaisselle", 
        "vitre", "brosse à dent", "facture", "argent", "vapeur", "ventilateur", "nettoyeur", "nettoyant", 
        "purificateur"
    ],
    "CUISINE": [
        "Airfryer", "ninja", "moulinex", "thermomix", "cookeo", "coffee", "delonghi", "tefal", 
        "cuisine", "plaque", "four", "cafetière", "café", "expresso", "barbecue", "kenwood", "soda", 
        "agrume", "frigo", "réfrigérateur", "congélateur", "plante", "jardin", "chauffage", "chaudière", 
        "climatiseur", "prise", "home"
    ],
    "MAISON": [
        "solaire", "électricité", "compteur", "gaz", "énergie", "EDF", "solar", "starlink", "box", 
        "hue", "ikea", "robot", "freebox", "livebox", "home", "matter", "fibre", "caméra", "surveillance", 
        "serrure", "nuki", "sonnette", "keypad", "clé", "thermostat", "linky", "tuya", "aqara"
    ],
    "SANTE": [
        "Santé", "fitbit", "bracelet", "circa", "montre", "watch", "sommeil", "dormir", "ondes", 
        "withings", "peau", "sport", "course", "natation", "coros", "coach", "outdoor", "garmin", 
        "polar", "amazfit", "suunto", "muscle", "musculation", "randonnée", "band", "bracelet", 
        "whoop", "ecg", "artérielle", "bague"
    ],
    "LUDIQUE": [
        "jeu", "jeux", "ludique", "jeux de société", "jeux vidéo", "jeu vidéo", "enfant", "cinéma", 
        "film", "switch", "playstation", "xbox", "manette", "VR", "casque VR", "netflix", "canal", 
        "TF1", "cinema", "série", "ps5", "ps4", "nintendo", "steam", "gaming", "spotify", "valve", 
        "zelda", "mario", "super mario", "mariokart", "Pokemon", "ubisoft", "console", "drone", 
        "neo geo", "sega", "GTA"
    ],
    "ACTU_JUR_POL": [
        "UE", "Cnil", "souveraineté", "europe", "chine", "russie", "Iran", "condamnée", "condamné", 
        "condamne", "gafam", "conseil d'état", "arcom", "RGPD", "IA Act", "Bruxelles", "union européenne", 
        "Otan", "procès", "plaintes", "règles", "réglementaton", "loi", "juge", "justice", "tribunal", 
        "sanction", "DSA", "politique", "onu", "la France", "Etat", "Trump", "commission européenne", 
        "géopolitique", "diplomatie", "cyberattaque", "piratage", "taxe", "présidentielle", "amende", 
        "décret", "législation", "démarchage", "fraude", "vie privée", "interdiction", "réseaux sociaux", 
        "arnaque", "fuite de données", "régulateur", "régulation", "légal", "légalité", "illégal", 
        "illégalle", "ANSSI"
    ],
    "ACTU_ECO": [
        "dépense", "bourse", "milliards", "rachat", "capitalisation", "acquisition", "argent", "OPA", 
        "fusion", "bulle", "usine", "production", "productivité", "salaires", "salariés", "employé", 
        "grève", "énergie", "énergétique", "rente", "rentable", "perte", "bénéfice", "chiffre d'affaires", 
        "comptable", "crise", "pénurie", "inflation", "résultats financiers", "chômage", "licenciement", 
        "patron"
    ],
    "SCIENCES": [
        "Nasa", "planète", "spatiale", "fusée", "lune", "soleil", "mars", "CNRS", "téléscope"
    ],
    "SPONSO": [
        "bon plan", "bons plans", "brade", "promotion", "sponsorisé"
    ]
}

# ============================================================
# 3. Fonctions Utilitaires & Catégorisation
# ============================================================

def log_error(msg):
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now(timezone.utc)}] {msg}\n")

def normalize(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    return text.lower()

# Pré-normalisation des mots-clés
NORM_CATEGORIES = {
    cat: [normalize(kw) for kw in keywords]
    for cat, keywords in CATEGORIES.items()
}

def categorize(title):
    """
    Catégorise selon les mots du titre avec isolation par frontière de mot (\b).
    - Mot simple : détecte le singulier et le pluriel (avec un 's' final).
    - Expression composée : détecte le terme exact.
    """
    norm_title = normalize(title)
    for category, keywords in NORM_CATEGORIES.items():
        for kw in keywords:
            if " " in kw or "-" in kw or "'" in kw:
                pattern = r"\b" + re.escape(kw) + r"\b"
            else:
                pattern = r"\b" + re.escape(kw) + r"s?\b"
                
            if re.search(pattern, norm_title):
                return category
    return "NON_CLASSE"

def parse_date(date_val):
    if not date_val:
        return None
    if isinstance(date_val, datetime):
        return date_val if date_val.tzinfo else date_val.replace(tzinfo=timezone.utc)
    try:
        dt = datetime.fromisoformat(str(date_val).replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None

# ============================================================
# 4. Scraping RSS et HTML
# ============================================================

def scrape_rss(name, url):
    print(f"[RSS] Collecte de {name}...")
    articles = []
    try:
        feed = feedparser.parse(url)
        for entry in feed.entries:
            dt = None
            if hasattr(entry, "published_parsed") and entry.published_parsed:
                try:
                    dt = datetime(*entry.published_parsed[:6], tzinfo=timezone.utc)
                except Exception:
                    dt = None
            elif hasattr(entry, "updated_parsed") and entry.updated_parsed:
                try:
                    dt = datetime(*entry.updated_parsed[:6], tzinfo=timezone.utc)
                except Exception:
                    dt = None

            title = entry.get("title", "").strip()
            link = entry.get("link", "").strip()

            if title and link:
                articles.append({
                    "source": name,
                    "date": dt,
                    "titre": title,
                    "lien": link,
                    "categorie": categorize(title)
                })
    except Exception as e:
        log_error(f"Erreur RSS {name}: {e}")
    return articles

def get_frandroid_article_date(url):
    try:
        resp = requests.get(url, timeout=8)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            time_tag = soup.find("time")
            if time_tag and time_tag.get("datetime"):
                return parse_date(time_tag["datetime"])
    except Exception as e:
        log_error(f"Erreur date Frandroid {url}: {e}")
    return None

def scrape_frandroid_html():
    articles = []
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    for section_name, base_url in FRANDROID_SECTIONS.items():
        for page in range(1, 4):
            url = base_url if page == 1 else f"{base_url}/page/{page}"
            print(f"[HTML] Scraping {section_name} - Page {page}...")
            try:
                resp = requests.get(url, headers=headers, timeout=10)
                if resp.status_code != 200:
                    continue
                soup = BeautifulSoup(resp.text, "html.parser")
                cards = soup.find_all("article")

                for card in cards:
                    a = card.find("a")
                    if not a:
                        continue
                    link = a.get("href", "").strip()
                    title = a.get_text(strip=True)

                    if link and title:
                        dt = get_frandroid_article_date(link)
                        articles.append({
                            "source": section_name,
                            "date": dt,
                            "titre": title,
                            "lien": link,
                            "categorie": categorize(title)
                        })
            except Exception as e:
                log_error(f"Erreur Frandroid HTML {url}: {e}")
    return articles

# ============================================================
# 5. Gestion de l'Historique (Lecture, Fusion, Déduplication)
# ============================================================

def load_existing_history():
    articles = []
    if not os.path.exists(OUTPUT_FILE):
        return articles

    try:
        df_old = pd.read_csv(OUTPUT_FILE, sep=";", encoding="utf-8-sig")
        for _, row in df_old.iterrows():
            title = str(row.get("titre", "")).strip()
            link = str(row.get("lien", "")).strip()
            source = str(row.get("source", "")).strip()
            date_val = parse_date(row.get("date", ""))
            
            # Recatégorise systématiquement avec les nouvelles catégories
            cat = str(row.get("categorie", ""))
            if not cat or cat == "nan" or cat == "NON_CLASSE":
                cat = categorize(title)
            else:
                new_cat = categorize(title)
                if new_cat != "NON_CLASSE":
                    cat = new_cat

            if link and title:
                articles.append({
                    "source": source,
                    "date": date_val,
                    "titre": title,
                    "lien": link,
                    "categorie": cat
                })
        print(f"[Historique] {len(articles)} anciens articles chargés depuis le CSV.")
    except Exception as e:
        log_error(f"Erreur chargement CSV existant : {e}")

    return articles

# ============================================================
# MAIN
# ============================================================

def main():
    all_articles = []

    # 1. Charger l'historique CSV existant
    all_articles.extend(load_existing_history())

    # 2. Collecter les flux RSS
    for name, url in RSS_FEEDS.items():
        all_articles.extend(scrape_rss(name, url))

    # 3. Collecter les rubriques HTML de Frandroid (3 pages chacune)
    all_articles.extend(scrape_frandroid_html())

    # 4. Déduplication stricte par URL (lien identique)
    unique = {}
    for a in all_articles:
        link = a["lien"]
        if link not in unique or (unique[link]["date"] is None and a["date"] is not None):
            unique[link] = a

    cleaned_list = list(unique.values())

    # 5. Tri chronologique décroissant (plus récents en haut)
    cleaned_list.sort(
        key=lambda x: (x["date"] is None, x["date"] or datetime.min.replace(tzinfo=timezone.utc)),
        reverse=True
    )

    # 6. Écriture dans le CSV final
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["source", "date", "titre", "lien", "categorie"])

        for a in cleaned_list:
            date_str = a["date"].strftime("%d/%m/%Y %H:%M:%S") if a["date"] else ""
            writer.writerow([
                a["source"],
                date_str,
                a["titre"],
                a["lien"],
                a["categorie"]
            ])

    print(f"✔ Terminé ! {len(cleaned_list)} articles enregistrés dans {OUTPUT_FILE}")

if __name__ == "__main__":
    main()