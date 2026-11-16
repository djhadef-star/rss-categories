from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
from datetime import datetime, timezone
import os
import re
import unicodedata
import urllib.parse
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import feedparser
import pandas as pd
import requests

# --- CONFIGURATION DES CHEMINS ---
CSV_FILE = os.path.join("output", "rss_history_2.csv")
EXCEL_FILE = os.path.join("output", "rss_history_2.xlsx")
RSS_XML_FILE = os.path.join("output", "feed.xml")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

# --- FLUX RSS ---
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
    "CCM": "https://www.commentcamarche.net/rss/",
}

# --- DICTIONNAIRE MIS À JOUR (20 CATÉGORIES) ---
CATEGORIES = {
    "1. SMARTPHONE": [
        "smartphone", "téléphone", "phone", "l'iphone", "iphone", "pixel", "flip", "siri", "forfait", "android", "GrapheneOS", "ios", "5G", "6G", "réseau", "Esim", "traceur", "Tag", "lunette", "appli", "application", "redmi", "whatsapp", "sosh", "Honor", '"One Ui"', "poco", "realme", '"S25"', '"S26"', '"S27"', '"S28"', "oneplus", "oppo", "pliant", "pliable", "BlackBerry", "motorola",
    ],
    "2. PC_OS": [
        "PC", "Geekom", "Framework", "SSD", "serveur", "Nas", "portable", "macbook", "asus", "lenovo", "HP", "Dell", "acer", "MSI", "ryzen", "qualcomm", '"Media Tek"', "mediatek", "snapdragon", "amd", "nvidia", '"puce graphique"', '"Galaxy book"', '"carte mère"', '"cartes mères"', "GPU", "arm", "laptop", "linux", "Bios", "windows", "Macos", "gnome", "fedora", "ubuntu", "patch", "DNS",
    ],
    "3. Periphérique": [
        "Clavier", "écran", "moniteur", "souris", "l'imprimante", "imprimante", "logitech", "corsair", "steam deck", "chargeur", "charge", '"batterie externe"', "ugreen", "branche", "lexar", "périphérique", "hub", "dock", "usb", "usb-c", "chaise", '"RJ 45"', "Wifi", "wi-fi", "routeur", "netgear", "ethernet", "répéteur", '"TP-link"',
    ],
    "4. Tablette": [
        "Tablette", "Fold", "pad", "tab", "kindle", "remarkable", "liseuse", "kobo", "ebooks", "pliable", "pliure", "pli", '"Galaxy tab"', "livre",
    ],
    "5. Logiciels _ OS": [
        "Logiciels", "software", '"power Toys"', "office", "mail", "rss", "slack", "teams", "capture", "outlook", "onenote", "excel", "word", "gmail", "drive", "cleaner", "VPN", "traduction", "chat", "automatisation", "edge", "chrome", "safari", "firefox", "mozilla", "opera", "vivaldi", "onglet", '"moteur de recherche"', "github", "zip", "terminal", "torrent", "chatbot", "interface", "wordpress", "web", "navigateur", '"google keep"', "virus", "bloqueur", "expressvpn", "cyberghost", "surfshark", "pcloud", "cpanel",
    ],
    "6. IA": [
        '"l\'IA"', "IA", '"intelligence artificielle"', '"data center"', "LLM", "token", "anthropic", '"d\'openAI"', '"openAI"', '"chatgpt"', "gemini", "meta", "Claude", "openclaw", "perplexity", "prompt", "copilot", "deepseek", "mistral", "groq", "midjourney", "grok", '"llm"', '"agent ai"', '"outil ai"', '"fonctionnalités ai"',
    ],
    "7. Video": [
        "Iptv", "vids", "télévision", "téléviseur", "TV", "omni", "pics", "vlc", "video", "éditeur", "3D", "plex", "LG", "hisense", "TCL", "Hdmi", "RGB", "videoprojecteur", "Awol", "xgimi", "Jmgo", "stick", "DJI", "osmo", "Insta360", "cast", "graphique", "édition", "image", "gopro", "affichage",
    ],
    "8. Audio": [
        "casque", "audio", '"casque audio"', '"casque gaming"', "pods", "buds", "headphone", "jbl", "shokz", "clip", "écouteurs", '"réduction de bruit"', "enceinte", '"barre de son"', "harman", "sennheiser", "bose", "sonos", "égaliseur", "musique", "voix", "platine", "vinyle", "oreille", "openfit", "aeroclip", "chanson",
    ],
    "9. Moiilité douce": [
        "mobilité", "vélo", "trottinette", "cargo", "sacoche", "antivol", "VAE", "navigo", "blablacar", "taxi", "rer", '"transport en commun"', "ebike", "vtt", "segway", "Navee", "veste", "avion", "aéroport", "GPS", "pompe", '"compteur vélo"',
    ],
    "10. Voiture": [
        "voiture", "automobile", "conduite", "conduire", "rouler", "essence", "carburant", "monospace", "berline", "citadine", "SUV", "permis", "volkswagen", "mercedes", "bmw", "byd", "mg", "renault", "peugeot", "toyota", "nissan", "stellantis", "citroen", "dacia", "kia", "fiat", "tesla", "ford", "jeep", "twingo", "volvo", "autonome", "FSD", "waze", '"google maps"', "auto", "carplay", "taxe", "moto", "scooter", "hybride", '"voiture électrique"', "borne", "véhicule", "recharge", "robotaxi", "porshe", "CATL",
    ],
    "11. Menage": [
        "ménage", "aspirateur", '"robot aspirateur"', '"aspirateurs robots"', '"aspirateur robot"', "dyson", "roborock", "narwal", "mova", "ecovacs", "dreame", "tineco", "shark", "laveur", "lavage", '"machine à laver"', "vitres", "pressing", "défroisseur", '"centrale vapeur"', '"fer à repasser"', "tambour", "maison", "linge", "vaisselle", '"lave-vaisselle"', "vitre", '"brosse à dent"', "vapeur", "ventilateur", "nettoyeur", "nettoyant", "purificateur", "serpillère",
    ],
    "12. Cuisine": [
        "Airfryer", "ninja", "moulinex", "thermomix", "cookeo", "coffee", "delonghi", "tefal", "cuisine", "plaque", "four", "cafetière", "café", "expresso", "cappuccino", "barista", "barbecue", "kenwood", "soda", "agrume", "frigo", "réfrigérateur", "congélateur", "broyeur", "krups", '"cuisson"', '"robot cuiseur"', "vin", "bouilloire",
    ],
    "13. Maison": [
        "box", "hue", "ikea", "freebox", "livebox", "home", "matter", "fibre", "caméra", "surveillance", "serrure", "nuki", "sonnette", "keypad", "clé", "thermostat", "linky", "tuya", "aqara", "prise", "interrupteurs",
    ],
    "14. Extérieur": [
        "solaire", "électricité", "compteur", "gaz", "énergie", "EDF", "solar", "starlink", "plante", "jardin", "chauffage", "chaudière", "climatiseur", '"robot tondeuse"',
    ],
    "15. Santé": [
        "Santé", "fitbit", "bracelet", "circa", "montre", "watch", "sommeil", "dormir", "ondes", "withings", "peau", "sport", "course", "natation", "coros", "coach", "outdoor", "garmin", "polar", "amazfit", "suunto", "muscle", "musculation", "randonnée", "band", "whoop", "ecg", "artérielle", "bague", "ring", "rasoir",
    ],
    "16. ludique": [
        "jeu", "jeux", "ludique", '"jeux de société"', '"jeux vidéo"', '"jeu vidéo"', "enfant", "cinéma", "film", '"court métrage"', '"long métrage"', "switch", "playstation", "xbox", "manette", "VR", '"casque VR"', "netflix", "canal", "TF1", "cinema", "série", "ps6", "ps5", "ps4", "nintendo", "steam", "gaming", "spotify", "valve", "zelda", "mario", '"super mario"', "mariokart", "Pokemon", "ubisoft", "console", "drone", '"neo geo"', "sega", "GTA", "streaming", "media", "youtube", "Disney+", "Disney", '"game boy"', "lego", '"hbo max"', '"apple TV"', "paramount", "saison", '"chef d\'oeuvre"', '"chef-d\'oeuvre"', "roman", "dessin", "dessine", "bande-dessinée", "manga", "teaser", "casting", "scène", "personnage", "ligue 1", '"ligue des champions"', '"coupe d\'europe"', "FIFA", "UEFA", "football", "tennis", "basket", "rugby", "handball",
    ],
    "17. actu jur et pol": [
        "UE", "Cnil", "souveraineté", "europe", "Etats-unis", "chine", "russie", "Iran", "condamnée", "condamné", "condamne", "gafam", '"conseil d\'état"', "arcom", "RGPD", '"IA Act"', "Bruxelles", '"union européenne"', "Otan", "procès", "plaintes", "règles", "réglementation", "loi", "juge", "justice", "tribunal", "sanction", "DSA", "politique", "onu", '"la France"', "Etat", "Trump", '"commission européenne"', "géopolitique", "diplomatie", "cyberattaque", "piratage", "taxe", "présidentielle", "amende", "décret", "législation", "démarchage", "fraude", '"vie privée"', "interdiction", '"réseaux sociaux"', "arnaque", '"fuite de données"', "régulateur", "régulation", "légal", "légalité", "illégal", "illégale", "ANSSI", "crypto", "police", "données", "coookies", "ministre", "ministère", "usurpation", "guerre", "impôts", "fisc", "hacker", "conforme",
    ],
    "18. Actualité éco": [
        "dépense", "bourse", "milliards", "rachat", "capitalisation", "acquisition", "argent", "OPA", "fusion", "bulle", "usine", "production", "productivité", "salaires", "salariés", "employé", "grève", "énergie", "énergétique", "rente", "rentable", "perte", "bénéfice", '"chiffre d\'affaires"', "comptable", "crise", "pénurie", "inflation", '"résultats financiers"', "chômage", "licenciement", "patron",
    ],
    "19. Sciences": [
        "Nasa", "planète", "spatiale", "fusée", "lune", "soleil", "mars", "CNRS", "téléscope", '"l\'espace"', "satellite", '"Système solaire"', "chercheur", "nucléaire", "astéroïde", "cratère", "iss", '"l\'iss"', "spacex", "lunaire", "implant", "cerveau", "astronome", "astronaute",
    ],
    "20. Sponso": [
        '"bon plan"', '"bons plans"', "bonplan", "bon-plan", "bons-plans", "brade", "promotion", "sponsorisé", "Aliexpress", "promo", "aliexpress",
    ],
}


# --- FONCTIONS UTILITAIRES ---
def remove_accents(text: str) -> str:
    """Supprime les accents et passe en minuscules."""
    if not text:
        return ""
    normalized = unicodedata.normalize("NFD", text)
    return "".join(c for c in normalized if unicodedata.category(c) != "Mn").lower()


def clean_url(url: str) -> str:
    """Nettoie les paramètres UTM de tracking des URLs."""
    parsed = urllib.parse.urlparse(url)
    qd = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    filtered = {k: v for k, v in qd.items() if not k.startswith("utm_")}
    return urllib.parse.urlunparse(
        parsed._replace(query=urllib.parse.urlencode(filtered, doseq=True))
    )


def build_regex_pattern(keyword: str) -> str:
    clean_kw = (
        remove_accents(keyword)
        .replace('"', "")
        .replace("«", "")
        .replace("»", "")
        .strip()
    )
    if not clean_kw:
        return None

    words = clean_kw.split()
    if len(words) > 1:
        first_part = r"\s+".join(re.escape(w) for w in words[:-1])
        last_part = re.escape(words[-1]) + r"(?:s|x)?\b"
        return r"\b" + first_part + r"\s+" + last_part
    else:
        return r"\b" + re.escape(clean_kw) + r"(?:s|x)?\b"


def categorize_article(title: str, url: str = "") -> str:
    text_title = remove_accents(title)
    text_url = remove_accents(url)

    sponso_category = next(
        (c for c in CATEGORIES if "sponso" in c.lower()), None
    )
    if sponso_category:
        full_text_sponso = f"{text_title} {text_url}"
        for keyword in CATEGORIES[sponso_category]:
            pattern = build_regex_pattern(keyword)
            if pattern and re.search(pattern, full_text_sponso):
                return sponso_category

    first_match_index = float("inf")
    best_category = "NON_CLASSE"

    for category, keywords in CATEGORIES.items():
        if category == sponso_category:
            continue

        for keyword in keywords:
            pattern = build_regex_pattern(keyword)
            if not pattern:
                continue

            match = re.search(pattern, text_title)
            if match:
                match_start = match.start()
                if match_start < first_match_index:
                    first_match_index = match_start
                    best_category = category

    return best_category


def parse_date(date_str: str):
    if not date_str:
        return None
    try:
        dt = datetime.fromisoformat(date_str)
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None


# --- SCRAPING LÉGER DU CONTENU POUR LE RSS ---
def fetch_article_description(link: str, original_summary: str = "") -> str:
    """Tente de récupérer un court extrait ou la description de l'article pour le RSS."""
    if original_summary and len(original_summary.strip()) > 30:
        return original_summary.strip()
    
    if not link or not link.startswith("http"):
        return ""

    try:
        res = requests.get(link, headers=HEADERS, timeout=5)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            # Cherche un paragraphe pertinent ou une méta description
            meta_desc = soup.find("meta", attrs={"name": "description"})
            if meta_desc and meta_desc.get("content"):
                return meta_desc["content"].strip()
            
            p_tag = soup.find("p")
            if p_tag:
                return p_tag.get_text(strip=True)
    except Exception:
        pass
    return original_summary or ""


# --- MODULES DE SCRAPING ---
def scrape_rss(source_name: str, feed_url: str) -> list:
    items = []
    try:
        response = requests.get(feed_url, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            return items
        
        feed = feedparser.parse(response.content)
        for entry in feed.entries:
            title = entry.get("title", "").strip()
            link = clean_url(entry.get("link", "").strip())
            if not title or not link:
                continue

            summary = entry.get("summary", "") or entry.get("description", "")
            # Nettoyage basique des balises HTML du résumé RSS d'origine
            if summary:
                summary = BeautifulSoup(summary, "html.parser").get_text(strip=True)

            dt_tuple = entry.get("published_parsed") or entry.get("updated_parsed")
            dt = (
                datetime(*dt_tuple[:6], tzinfo=timezone.utc) if dt_tuple else None
            )

            items.append({
                "source": source_name,
                "date": dt,
                "titre": title,
                "lien": link,
                "description": summary,
                "categorie": categorize_article(title, link),
            })
    except Exception as e:
        print(f"Erreur RSS {source_name}: {e}")
    return items


def scrape_frandroid_web(url: str, source_name: str) -> list:
    items = []
    try:
        res = requests.get(url, headers=HEADERS, timeout=10)
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
            if not title:
                continue

            time_tag = article.find("time")
            dt = None
            if time_tag and time_tag.has_attr("datetime"):
                dt = parse_date(time_tag["datetime"].replace("Z", "+00:00"))

            items.append({
                "source": source_name,
                "date": dt,
                "titre": title,
                "lien": link,
                "description": "",
                "categorie": categorize_article(title, link),
            })
    except Exception as e:
        print(f"Erreur Scraping Web {source_name}: {e}")
    return items


# --- HISTORIQUE & EXPORT ---
def load_existing_history(filepath: str) -> dict:
    history = {}
    if not os.path.exists(filepath):
        return history

    with open(filepath, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            link = row.get("lien")
            if link:
                history[link] = {
                    "source": row.get("source", ""),
                    "date": parse_date(row.get("date")),
                    "titre": row.get("titre", ""),
                    "lien": link,
                    "description": row.get("description", ""),
                    "categorie": row.get("categorie", "NON_CLASSE"),
                }
    return history


def generate_excel_from_csv():
    if not os.path.exists(CSV_FILE):
        print("CSV introuvable. Annulation de la génération Excel.")
        return

    df_all = pd.read_csv(CSV_FILE, delimiter=";")

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="Tous_les_articles", index=False)

        for cat_name in CATEGORIES.keys():
            sheet_title = cat_name[:31]
            df_filtered = df_all[df_all["categorie"] == cat_name]
            df_filtered.to_excel(writer, sheet_name=sheet_title, index=False)

        df_unclassed = df_all[df_all["categorie"] == "NON_CLASSE"]
        if not df_unclassed.empty:
            df_unclassed.to_excel(writer, sheet_name="NON_CLASSE", index=False)

    print(f"Fichier Excel mis à jour : '{EXCEL_FILE}' !")


def generate_rss_xml_from_csv():
    if not os.path.exists(CSV_FILE):
        return

    df_all = pd.read_csv(CSV_FILE, delimiter=";")
    rss = ET.Element("rss", version="2.0")
    channel = ET.SubElement(rss, "channel")

    ET.SubElement(channel, "title").text = "Mon Flux RSS Multi-Sites Catégorisé"
    ET.SubElement(channel, "link").text = "https://github.com"
    ET.SubElement(channel, "description").text = "Flux RSS personnel agrégé et trié"

    for _, row in df_all.head(200).iterrows():
        item = ET.SubElement(channel, "item")
        category_name = str(row["categorie"]) if pd.notna(row["categorie"]) else "NON_CLASSE"
        title_text = str(row["titre"]) if pd.notna(row["titre"]) else "Sans titre"

        ET.SubElement(item, "title").text = f"[{category_name}] {title_text}"
        ET.SubElement(item, "link").text = str(row["lien"]) if pd.notna(row["lien"]) else ""

        if pd.notna(row["date"]):
            ET.SubElement(item, "pubDate").text = str(row["date"])

        desc_text = str(row["description"]) if pd.notna(row["description"]) else ""
        ET.SubElement(item, "description").text = desc_text

        ET.SubElement(item, "category").text = category_name
        ET.SubElement(item, "source").text = str(row["source"]) if pd.notna(row["source"]) else "Inconnu"

    tree = ET.ElementTree(rss)
    with open(RSS_XML_FILE, "wb") as f:
        tree.write(f, encoding="utf-8", xml_declaration=True)


def generate_category_rss_feeds():
    if not os.path.exists(CSV_FILE):
        return

    df_all = pd.read_csv(CSV_FILE, delimiter=";")
    feeds_dir = os.path.join("output", "feeds")
    os.makedirs(feeds_dir, exist_ok=True)

    for cat_name in CATEGORIES.keys():
        df_cat = df_all[df_all["categorie"] == cat_name]
        if df_cat.empty:
            continue

        safe_cat_name = re.sub(r'[\s/]+', '_', cat_name)
        xml_filename = os.path.join(feeds_dir, f"{safe_cat_name}.xml")

        rss = ET.Element("rss", version="2.0")
        channel = ET.SubElement(rss, "channel")

        ET.SubElement(channel, "title").text = f"Veille - {cat_name}"
        ET.SubElement(channel, "link").text = "https://github.com"
        ET.SubElement(channel, "description").text = f"Flux RSS spécialisé pour {cat_name}"

        for _, row in df_cat.head(100).iterrows():
            item = ET.SubElement(channel, "item")
            ET.SubElement(item, "title").text = str(row["titre"]) if pd.notna(row["titre"]) else "Sans titre"
            ET.SubElement(item, "link").text = str(row["lien"]) if pd.notna(row["lien"]) else ""
            if pd.notna(row["date"]):
                ET.SubElement(item, "pubDate").text = str(row["date"])
            
            desc_text = str(row["description"]) if pd.notna(row["description"]) else ""
            ET.SubElement(item, "description").text = desc_text

            ET.SubElement(item, "category").text = cat_name
            ET.SubElement(item, "source").text = str(row["source"]) if pd.notna(row["source"]) else "Inconnu"

        tree = ET.ElementTree(rss)
        with open(xml_filename, "wb") as f:
            tree.write(f, encoding="utf-8", xml_declaration=True)

    # Gestion des articles non classés
    df_unclassed = df_all[df_all["categorie"] == "NON_CLASSE"]
    if not df_unclassed.empty:
        xml_filename = os.path.join(feeds_dir, "NON_CLASSE.xml")
        rss = ET.Element("rss", version="2.0")
        channel = ET.SubElement(rss, "channel")

        ET.SubElement(channel, "title").text = "Veille - NON_CLASSE"
        ET.SubElement(channel, "link").text = "https://github.com"
        ET.SubElement(channel, "description").text = "Flux RSS des articles non classés"

        for _, row in df_unclassed.head(100).iterrows():
            item = ET.SubElement(channel, "item")
            ET.SubElement(item, "title").text = str(row["titre"])
            ET.SubElement(item, "link").text = str(row["lien"])
            if pd.notna(row["date"]):
                ET.SubElement(item, "pubDate").text = str(row["date"])
            
            desc_text = str(row["description"]) if pd.notna(row["description"]) else ""
            ET.SubElement(item, "description").text = desc_text

            ET.SubElement(item, "category").text = "NON_CLASSE"
            ET.SubElement(item, "source").text = str(row["source"])

        tree = ET.ElementTree(rss)
        with open(xml_filename, "wb") as f:
            tree.write(f, encoding="utf-8", xml_declaration=True)

    print("Flux RSS individuels par catégorie générés dans 'output/feeds/' avec succès !")


# --- MAIN PIPELINE AVEC MULTITHREADING (PARALLÉLISATEUR) ---
def main():
    os.makedirs("output", exist_ok=True)

    history = load_existing_history(CSV_FILE)
    print(f"Articles dans l'historique initial : {len(history)}")

    for link, item in history.items():
        item["categorie"] = categorize_article(item["titre"], link)

    new_items_count = 0

    # 3. Scraping RSS en parallèle (ultra-rapide)
    def fetch_single_rss(source, url):
        print(f"Scraping RSS : {source}...")
        return scrape_rss(source, url)

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_rss = {
            executor.submit(fetch_single_rss, source, url): source 
            for source, url in RSS_FEEDS.items()
        }
        for future in as_completed(future_to_rss):
            items = future.result()
            for item in items:
                link = item["lien"]
                if link not in history:
                    # Enrichissement intelligent de la description si vide
                    if not item["description"]:
                        item["description"] = fetch_article_description(link)
                    history[link] = item
                    new_items_count += 1
                else:
                    history[link]["categorie"] = categorize_article(
                        history[link]["titre"], link
                    )

    # 4. Scraping Web Frandroid en parallèle
    frandroid_sections = [
        ("Frandroid - Actualités", "https://www.frandroid.com/actualites"),
        ("Frandroid - Bons plans", "https://www.frandroid.com/bons-plans"),
    ]

    def fetch_single_web(source_name, url):
        print(f"Scraping Web : {source_name}...")
        return scrape_frandroid_web(url, source_name)

    with ThreadPoolExecutor(max_workers=2) as executor:
        future_to_web = {
            executor.submit(fetch_single_web, source_name, url): source_name 
            for source_name, url in frandroid_sections
        }
        for future in as_completed(future_to_web):
            items = future.result()
            for item in items:
                link = item["lien"]
                if link not in history:
                    if not item["description"]:
                        item["description"] = fetch_article_description(link)
                    history[link] = item
                    new_items_count += 1
                else:
                    history[link]["categorie"] = categorize_article(
                        history[link]["titre"], link
                    )

    print(f"Nouveaux articles ajoutés : {new_items_count}")

    # 5. Tri antéchronologique
    cleaned_list = list(history.values())
    cleaned_list.sort(
        key=lambda x: (
            x["date"] is not None,
            x["date"] or datetime.min.replace(tzinfo=timezone.utc),
        ),
        reverse=True,
    )

    # 6. Écriture du fichier CSV mis à jour avec la colonne description
    with open(CSV_FILE, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["source", "date", "titre", "lien", "description", "categorie"])
        for item in cleaned_list:
            date_str = item["date"].isoformat() if item["date"] else ""
            writer.writerow([
                item["source"], 
                date_str, 
                item["titre"], 
                item["lien"], 
                item.get("description", ""), 
                item["categorie"]
            ])

    print("CSV d'historique mis à jour avec succès !")

    # 7. Génération des livrables
    generate_excel_from_csv()
    generate_rss_xml_from_csv()
    generate_category_rss_feeds()


if __name__ == "__main__":
    main()
