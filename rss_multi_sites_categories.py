import os
import re
import csv
import urllib.parse
import unicodedata
from datetime import datetime, timezone
import feedparser
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Chemins des fichiers
CSV_FILE = os.path.join("output", "rss_history.csv")
EXCEL_FILE = os.path.join("output", "rss_history.xlsx")

# Flux RSS
RSS_FEEDS = {
    "01net": "https://www.01net.com/feed/",
    "JournalDuGeek": "https://www.journaldugeek.com/feed/",
    "LesNumeriques": "https://www.lesnumeriques.com/rss.xml",
    "Frandroid - RSS": "https://www.frandroid.com/feed",
    "Phonandroid": "https://www.phonandroid.com/feed",
    "Numerama": "https://www.numerama.com/feed/",
    "KultureGeek": "https://kulturegeek.fr/feed",
    "TomsHardware": "https://www.tomshardware.fr/feed/",
    "NextInpact": "https://www.nextinpact.com/rss/news.xml",
    "Korben": "https://korben.info/feed",
    "CCM": "https://www.commentcamarche.net/rss/rss-actualites",
}

# Dictionnaire des 19 catégories mises à jour
CATEGORIES = {
    "1. SMARTPHONE": [
        "smartphone", "téléphone", "phone", "l'iphone", "iphone", "pixel", "flip",
        "siri", "forfait", "android", "grapheneos", "ios", "5g", "6g", "réseau",
        "esim", "traceur", "tag", "lunette", "appli", "application", "anker",
        "redmi", "whatsapp", "sosh", "honor", "one ui", "poco", "realme",
        "galaxy s", "oppo", "pliant", "pliable"
    ],
    "2. PC": [
        "pc", "geekom", "framework", "ssd", "serveur", "nas", "portable", "macbook",
        "asus", "lenovo", "hp", "dell", "acer", "msi ryzen", "qualcomm",
        "media tek", "mediatek", "snapdragon", "amd", "nvidia", "puce graphique",
        "galaxy book", "carte mère", "cartes mères", "gpu", "arm"
    ],
    "3. Periphérique": [
        "clavier", "écran", "moniteur", "souris", "l'imprimante", "imprimante",
        "logitech", "corsair", "steam deck", "chargeur", "charge", "batterie externe",
        "ugreen", "branche", "lexar", "périphérique", "hub", "dock", "usb", "usb-c",
        "ram", "chaise", "rj 45", "wifi", "wi-fi", "routeur", "netgear", "ethernet"
    ],
    "4. Tablette": [
        "tablette", "fold", "pad", "tab", "kindle", "remarkable", "liseuse", "kobo",
        "ebooks", "pliable", "pliure", "pli", "galaxy tab", "livre"
    ],
    "5. Logiciels _ OS": [
        "logiciels", "sofware", "windows", "macos", "power toys", "office", "linux",
        "bios", "os", "mail", "rss", "slack", "teams", "capture", "outlook",
        "onenote", "excel", "word", "gmail", "drive", "cleaner", "vpn",
        "traduction", "chat", "automatisation", "edge", "chrome", "safari",
        "firefox", "mozilla", "opera", "vivaldi", "moteur de recherche", "github",
        "zip", "terminal", "torrent", "chatbot", "interface", "wordpress", "web",
        "navigateur", "google keep", "patch", "virus", "bloqueur", "dns",
        "expressvpn", "cyberghost", "surfshark", "pcloud", "cpanel", "gnome",
        "fedora", "ubuntu"
    ],
    "6. IA": [
        "l'ia", "ia", "ai", "intelligence artificielle", "data center", "llm",
        "token", "anthropic", "d'openai", "openai", "chatgpt", "gemini", "meta",
        "claude", "openclaw", "perplexity", "prompt", "copilot", "deepseek",
        "mistral", "groq", "midjourney", "grok"
    ],
    "7. Video": [
        "iptv", "vids", "télévision", "téléviseur", "tv", "omni", "pics", "vlc",
        "video", "éditeur", "3d", "plex", "lg", "hisense", "tcl", "hdmi", "rgb",
        "vidéoprojecteur", "awol", "xgimi", "jmgo", "stick", "dji", "osmo",
        "insta360", "cast", "graphique", "édition", "image", "gopro", "affichage"
    ],
    "8. Audio": [
        "casque", "audio", "casque audio", "casque gaming", "pods", "buds",
        "headphone", "jbl", "shockz", "clip", "écouteurs", "réduction de bruit",
        "enceinte", "barre de son", "harman", "sennheiser", "bose", "sonos",
        "égaliseur", "musique", "voix", "son"
    ],
    "9. Moiilité douce": [
        "mobilité", "vélo", "trottinette", "cargo", "sacoche", "antivol", "vae",
        "navigo", "blablacar", "taxi", "rer", "transport en commun", "ebike",
        "vtt", "segway", "veste", "avion", "aéroport", "gps", "pompe"
    ],
    "10. Voiture": [
        "voiture", "automobile", "conduite", "conduire", "rouler", "essence",
        "carburant", "monospace", "berline", "citadine", "suv", "permis",
        "volkswagen", "mercedes", "bmw", "byd", "mg", "renault", "peugeot",
        "toyota", "nissan", "stellantis", "citroen", "dacia", "kia", "fiat",
        "tesla", "ford", "jeep", "twingo", "autonome", "fsd", "waze",
        "google maps", "auto", "carplay", "taxe", "moto", "scooter", "hybride",
        "voiture électrique", "borne", "véhicule", "recharge", "robotaxi",
        "porshe", "catl"
    ],
    "11. Menage": [
        "aspirateur", "robot aspirateur", "aspirateurs robots", "aspirateur robot",
        "dyson", "roborock", "narwal", "mova", "ecovacs", "dreame", "tineco",
        "shark", "laveur", "lavage", "machine à laver", "vitres", "pressing",
        "défroisseur", "centrale vapeur", "fer à repasser", "tambour", "maison",
        "linge", "vaisselle", "lave-vaisselle", "vitre", "brosse à dent",
        "facture", "argent", "vapeur", "ventilateur", "nettoyeur", "nettoyant",
        "purificateur"
    ],
    "12. Cuisine": [
        "airfryer", "ninja", "moulinex", "thermomix", "cookeo", "coffee",
        "delonghi", "tefal", "cuisine", "plaque", "four", "cafetière", "café",
        "expresso", "barbecue", "kenwood", "soda", "agrume", "frigo",
        "réfrigérateur", "congélateur", "plante", "jardin", "chauffage",
        "chaudière", "climatiseur", "prise", "home", "interrupteurs", "broyeur"
    ],
    "13. Maison": [
        "solaire", "électricité", "compteur", "gaz", "énergie", "edf", "solar",
        "starlink", "box", "hue", "ikea", "robot", "freebox", "livebox", "home",
        "matter", "fibre", "caméra", "surveillance", "serrure", "nuki",
        "sonnette", "keypad", "clé", "thermostat", "linky", "tuya", "aqara",
        "vin", "strarlink", "thermomix", "robot cuiseur"
    ],
    "14. Santé": [
        "santé", "fitbit", "bracelet", "circa", "montre", "watch", "sommeil",
        "dormir", "ondes", "withings", "peau", "sport", "course", "natation",
        "coros", "coach", "outdoor", "garmin", "polar", "amazfit", "suunto",
        "muscle", "musculation", "randonnée", "band", "whoop", "ecg", "artérielle",
        "bague"
    ],
    "15. ludique": [
        "jeu", "jeux", "ludique", "jeux de société", "jeux vidéo", "jeu vidéo",
        "enfant", "cinéma", "film", "switch", "playstation", "xbox", "manette",
        "vr", "casque vr", "netflix", "canal", "tf1", "série", "ps6", "ps5",
        "ps4", "nintendo", "steam", "gaming", "spotify", "valve", "zelda",
        "mario", "super mario", "mariokart", "pokemon", "ubisoft", "console",
        "drone", "neo geo", "sega", "gta", "streaming", "media", "youtube",
        "disney+", "disney", "game boy", "lego"
    ],
    "16. actu jur et pol": [
        "ue", "cnil", "souveraineté", "europe", "etats-unis", "chine", "russie",
        "iran", "condamnée", "condamné", "condamne", "gafam", "conseil d'état",
        "arcom", "rgpd", "ia act", "bruxelles", "union européenne", "otan",
        "procès", "plaintes", "règles", "réglementation", "loi", "juge",
        "justice", "tribunal", "sanction", "dsa", "politique", "onu",
        "la france", "etat", "trump", "commission européenne", "géopolitique",
        "diplomatie", "cyberattaque", "piratage", "taxe", "présidentielle",
        "amende", "décret", "législation", "démarchage", "fraude", "vie privée",
        "interdiction", "réseaux sociaux", "arnaque", "fuite de données",
        "régulateur", "régulation", "légal", "légalité", "illégal", "illégale",
        "anssi", "crypto", "police", "données", "coookies", "ministre",
        "ministère", "usurpation", "guerre", "impôts", "fisc"
    ],
    "17. Actualité éco": [
        "dépense", "bourse", "milliards", "rachat", "capitalisation", "acquisition",
        "argent", "opa", "fusion", "bulle", "usine", "production", "productivité",
        "salaires", "salariés", "employé", "grève", "énergie", "énergétique",
        "rente", "rentable", "perte", "bénéfice", "chiffre d'affaires", "comptable",
        "crise", "pénurie", "inflation", "résultats financiers", "chômage",
        "licenciement", "patron"
    ],
    "18. Sciences": [
        "nasa", "planète", "spatiale", "fusée", "lune", "soleil", "mars", "cnrs",
        "télescope", "l'espace", "satellite", "système solaire", "chercheur",
        "nucléaire", "astéroïde", "cratère", "iss", "l'iss", "spacex"
    ],
    "19. Sponso": [
        "bon plan", "bons plans", "bonplan", "bon-plan", "bons-plans", "brade", 
        "promotion", "sponsorisé", "aliexpress", "promo"
    ]
}


def remove_accents(text):
    if not text:
        return ""
    normalized = unicodedata.normalize('NFD', text)
    return "".join(c for c in normalized if unicodedata.category(c) != 'Mn').lower()


def clean_url(url):
    parsed = urllib.parse.urlparse(url)
    qd = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    filtered = {k: v for k, v in qd.items() if not k.startswith("utm_")}
    new_query = urllib.parse.urlencode(filtered, doseq=True)
    return urllib.parse.urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path, parsed.params, new_query, parsed.fragment)
    )


def categorize_article(title, url=""):
    text_title = remove_accents(title)
    text_url = remove_accents(url)

    first_match_index = float('inf')
    best_category = "NON_CLASSE"

    for category, keywords in CATEGORIES.items():
        # Pour la catégorie 19 (Sponso), on cherche dans Titre + URL.
        # Pour toutes les autres catégories, on cherche UNIQUEMENT dans le Titre.
        if "19" in category or "sponso" in category.lower():
            text_to_check = f"{text_title} {text_url}"
        else:
            text_to_check = text_title

        for keyword in keywords:
            clean_keyword = remove_accents(keyword)
            # \b garantit que le mot commence exactement au début (ex: 'RAM' ne matchera pas 'Frame')
            # \w* autorise le pluriel ou les déclinaisons (ex: 'Tesla' matchera 'Teslas')
            pattern = r'\b' + re.escape(clean_keyword) + r'\w*'
            
            match = re.search(pattern, text_to_check)
            if match:
                match_start = match.start()  # Position de la première occurrence
                
                # Priorité au mot-clé qui apparaît le plus tôt
                if match_start < first_match_index:
                    first_match_index = match_start
                    best_category = category

    return best_category


def parse_date(date_str):
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str)
    except Exception:
        return None


def scrape_rss(source_name, feed_url):
    items = []
    try:
        feed = feedparser.parse(feed_url)
        for entry in feed.entries:
            title = entry.get("title", "").strip()
            link = clean_url(entry.get("link", "").strip())

            dt = None
            if hasattr(entry, "published_parsed") and entry.published_parsed:
                dt = datetime(*entry.published_parsed[:6], tzinfo=timezone.utc)
            elif hasattr(entry, "updated_parsed") and entry.updated_parsed:
                dt = datetime(*entry.updated_parsed[:6], tzinfo=timezone.utc)

            category = categorize_article(title, link)

            items.append({
                "source": source_name,
                "date": dt,
                "titre": title,
                "lien": link,
                "categorie": category
            })
    except Exception as e:
        print(f"Erreur lors du scraping de {source_name}: {e}")

    return items


def scrape_frandroid_web(url, source_name):
    headers = {"User-Agent": "Mozilla/5.0"}
    items = []

    try:
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code != 200:
            return items

        soup = BeautifulSoup(res.text, "html.parser")
        articles = soup.find_all("article")

        for article in articles:
            a_tag = article.find("a", href=True)
            if not a_tag:
                continue

            link = clean_url(a_tag["href"])
            title = a_tag.get_text(strip=True)

            time_tag = article.find("time")
            dt = None
            if time_tag and time_tag.has_attr("datetime"):
                try:
                    dt = datetime.fromisoformat(time_tag["datetime"].replace("Z", "+00:00"))
                except Exception:
                    dt = None

            if title:
                category = categorize_article(title, link)
                items.append({
                    "source": source_name,
                    "date": dt,
                    "titre": title,
                    "lien": link,
                    "categorie": category
                })

    except Exception as e:
        print(f"Erreur lors du scraping Web {source_name}: {e}")

    return items


def load_existing_history(filepath):
    history = {}
    if not os.path.exists(filepath):
        return history

    with open(filepath, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            link = row.get("lien")
            if link:
                history[link] = {
                    "source": row.get("source", ""),
                    "date": parse_date(row.get("date")),
                    "titre": row.get("titre", ""),
                    "lien": link,
                    "categorie": row.get("categorie", "NON_CLASSE")
                }
    return history


def generate_excel_from_csv():
    """Génère le fichier Excel avec 1 onglet Global + 19 onglets par catégorie."""
    if not os.path.exists(CSV_FILE):
        print("Fichier CSV introuvable, annulation de la génération Excel.")
        return

    df_all = pd.read_csv(CSV_FILE, delimiter=';')

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        # Onglet 1: Vue globale
        df_all.to_excel(writer, sheet_name="Tous_les_articles", index=False)

        # Onglets 2 à 20: 19 onglets de catégories
        for cat_name in CATEGORIES.keys():
            sheet_title = cat_name[:31]  # Limite de caractères par nom d'onglet Excel
            df_filtered = df_all[df_all["categorie"] == cat_name]
            df_filtered.to_excel(writer, sheet_name=sheet_title, index=False)

        # Onglet optionnel pour les articles non classés
        df_unclassed = df_all[df_all["categorie"] == "NON_CLASSE"]
        if not df_unclassed.empty:
            df_unclassed.to_excel(writer, sheet_name="NON_CLASSE", index=False)

    print(f"Fichier Excel mis à jour avec succès : '{EXCEL_FILE}' !")


def main():
    os.makedirs("output", exist_ok=True)
    history = load_existing_history(CSV_FILE)
    print(f"Articles dans l'historique : {len(history)}")

    new_items_count = 0

    # 1. Scraping des flux RSS
    for source, url in RSS_FEEDS.items():
        print(f"Scraping RSS : {source}...")
        items = scrape_rss(source, url)
        for item in items:
            link = item["lien"]
            if link not in history:
                history[link] = item
                new_items_count += 1
            else:
                # Réévaluation systématique avec la nouvelle logique de catégorisation
                history[link]["categorie"] = categorize_article(history[link]["titre"], link)

    # 2. Scraping spécifique Web Frandroid
    frandroid_sections = [
        ("Frandroid - Actualités", "https://www.frandroid.com/actualites"),
        ("Frandroid - Bons plans", "https://www.frandroid.com/bons-plans")
    ]

    for source_name, url in frandroid_sections:
        print(f"Scraping Web : {source_name}...")
        items = scrape_frandroid_web(url, source_name)
        for item in items:
            link = item["lien"]
            if link not in history:
                history[link] = item
                new_items_count += 1
            else:
                # Réévaluation systématique avec la nouvelle logique de catégorisation
                history[link]["categorie"] = categorize_article(history[link]["titre"], link)

    print(f"Nouveaux articles ajoutés : {new_items_count}")

    # Tri antéchronologique
    cleaned_list = list(history.values())
    cleaned_list.sort(
        key=lambda x: (x["date"] is not None, x["date"] or datetime.min.replace(tzinfo=timezone.utc)),
        reverse=True
    )

    # 3. Sauvegarde dans le CSV
    with open(CSV_FILE, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(["source", "date", "titre", "lien", "categorie"])

        for item in cleaned_list:
            date_str = item["date"].isoformat() if item["date"] else ""
            writer.writerow([
                item["source"],
                date_str,
                item["titre"],
                item["lien"],
                item["categorie"]
            ])

    print("Fichier CSV mis à jour !")

    # 4. Actualisation de l'Excel multi-onglets depuis le CSV
    generate_excel_from_csv()


if __name__ == "__main__":
    main()
