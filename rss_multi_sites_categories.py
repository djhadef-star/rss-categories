import os
import re
import csv
import time
import urllib.parse
from datetime import datetime, timezone
import feedparser
import requests
from bs4 import BeautifulSoup

# File path definition
CSV_FILE = os.path.join("output", "rss_history.csv")

# RSS Feeds List
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

# Categories dictionary with Sponso in 1st position
CATEGORIES = {
    "19. Sponso": [
        "bon plan", "bons plans", "brade", "promotion", "sponsorisé", "comparatif",
        "bons-plans", "humanoid"
    ],
    "1. SMARTPHONE": [
        "smartphone", "téléphone", "phone", "iphone", "pixel", "flip", "siri", "forfait",
        "android", "grapheneos", "ios", "5g", "6g", "réseau", "esim", "wifi", "wi-fi",
        "routeur", "netgear", "ethernet", "traceur", "tag", "lunette", "appli",
        "application", "anker", "redmi", "whatsapp", "spacex", "starlink", "sosh",
        "honor", "one ui", "poco", "realme", "galaxy s", "oppo"
    ],
    "2. PC": [
        "pc", "geekom", "framework", "ssd", "serveur", "nas", "portable", "macbook",
        "asus", "lenovo", "hp", "dell", "acer", "msi", "ryzen", "qualcomm", "mediatek",
        "snapdragon", "amd", "nvidia", "puce graphique", "galaxy book", "carte mère"
    ],
    "3. Periphérique": [
        "clavier", "écran", "moniteur", "souris", "imprimante", "logitech", "corsair",
        "steam deck", "chargeur", "charge", "batterie externe", "ugreen", "branche",
        "lexar", "périphérique", "hub", "dock", "usb", "usb-c", "ram", "chaise", "rj45"
    ],
    "4. Tablette": [
        "tablette", "fold", "pad", "tab", "kindle", "remarkable", "liseuse", "kobo",
        "ebooks", "pliable", "pliure", "pli", "galaxy tab"
    ],
    "5. Logiciels _ OS": [
        "logiciels", "software", "windows", "macos", "powertoys", "office", "linux",
        "bios", "os", "mail", "rss", "slack", "teams", "capture", "outlook", "onenote",
        "gmail", "drive", "cleaner", "vpn", "traduction", "chat", "automatisation",
        "edge", "chrome", "safari", "firefox", "mozilla", "opera", "vivaldi",
        "moteur de recherche", "github", "zip", "terminal", "torrent", "chatbot",
        "interface", "wordpress", "web", "navigateur", "google keep", "patch", "virus",
        "bloqueur", "dns", "expressvpn", "cyberghost", "surfshark", "pcloud", "cpanel",
        "gnome", "fedora", "ubuntu"
    ],
    "6. IA": [
        "ia", "ai", "intelligence artificielle", "data center", "llm", "token",
        "anthropic", "openai", "chatgpt", "gemini", "meta", "claude", "openclaw",
        "perplexity", "prompt", "copilot", "deepseek", "mistral", "groq", "midjourney", "grok"
    ],
    "7. Video": [
        "iptv", "vids", "télévision", "téléviseur", "tv", "omni", "pics", "vlc", "video",
        "éditeur", "3d", "plex", "lg", "hisense", "tcl", "hdmi", "rgb", "vidéoprojecteur",
        "awol", "xgimi", "jmgo", "stick", "dji", "osmo", "insta360", "cast", "graphique",
        "édition", "image", "gopro", "affichage"
    ],
    "8. Audio": [
        "casque", "audio", "casque audio", "casque gaming", "pods", "buds", "headphone",
        "jbl", "shockz", "clip", "écouteurs", "réduction de bruit", "enceinte",
        "barre de son", "harman", "sennheiser", "bose", "sonos", "égaliseur", "musique"
    ],
    "9. Moiilité douce": [
        "mobilité", "vélo", "trottinette", "cargo", "sacoche", "antivol", "vae", "navigo",
        "blablacar", "taxi", "rer", "transport en commun", "ebike", "vtt", "segway",
        "veste", "avion", "aéroport"
    ],
    "10. Voiture": [
        "voiture", "automobile", "conduite", "conduire", "rouler", "essence", "carburant",
        "monospace", "berline", "citadine", "suv", "permis", "volkswagen", "mercedes",
        "bmw", "byd", "mg", "renault", "peugeot", "toyota", "nissan", "stellantis",
        "citroen", "dacia", "kia", "fiat", "tesla", "ford", "jeep", "twingo", "autonome",
        "fsd", "waze", "google maps", "auto", "carplay", "taxe", "moto", "scooter",
        "hybride", "voiture électrique", "borne"
    ],
    "11. Menage": [
        "aspirateur", "robot aspirateur", "dyson", "roborock", "narwal", "mova", "ecovacs",
        "dreame", "tineco", "shark", "laveur", "lavage", "machine à laver", "vitres",
        "pressing", "défroisseur", "centrale vapeur", "fer à repasser", "tambour",
        "maison", "linge", "vaisselle", "lave-vaisselle", "vitre", "brosse à dent",
        "facture", "argent", "vapeur", "ventilateur", "nettoyeur", "nettoyant", "purificateur"
    ],
    "12. Cuisine": [
        "airfryer", "ninja", "moulinex", "thermomix", "cookeo", "coffee", "delonghi",
        "tefal", "cuisine", "plaque", "four", "cafetière", "café", "expresso", "barbecue",
        "kenwood", "soda", "agrume", "frigo", "réfrigérateur", "congélateur", "plante",
        "jardin", "chauffage", "chaudière", "climatiseur", "prise", "home"
    ],
    "13. Maison": [
        "solaire", "électricité", "compteur", "gaz", "énergie", "edf", "solar", "starlink",
        "box", "hue", "ikea", "robot", "freebox", "livebox", "home", "matter", "fibre",
        "caméra", "surveillance", "serrure", "nuki", "sonnette", "keypad", "clé",
        "thermostat", "linky", "tuya", "aqara", "vin"
    ],
    "14. Santé": [
        "santé", "fitbit", "bracelet", "circa", "montre", "watch", "sommeil", "dormir",
        "ondes", "withings", "peau", "sport", "course", "natation", "coros", "coach",
        "outdoor", "garmin", "polar", "amazfit", "suunto", "muscle", "musculation",
        "randonnée", "band", "whoop", "ecg", "artérielle", "bague"
    ],
    "15. ludique": [
        "jeu", "jeux", "ludique", "jeux de société", "jeux vidéo", "jeu vidéo", "enfant",
        "cinéma", "film", "switch", "playstation", "xbox", "manette", "vr", "casque vr",
        "netflix", "canal", "tf1", "série", "ps5", "ps4", "nintendo", "steam", "gaming",
        "spotify", "valve", "zelda", "mario", "super mario", "mariokart", "pokemon",
        "ubisoft", "console", "drone", "neo geo", "sega", "gta", "streaming", "media"
    ],
    "16. actu jur et pol": [
        "ue", "cnil", "souveraineté", "europe", "chine", "russie", "iran", "condamnée",
        "condamné", "condamne", "gafam", "conseil d'état", "arcom", "rgpd", "ia act",
        "bruxelles", "union européenne", "otan", "procès", "plaintes", "règles",
        "réglementation", "loi", "juge", "justice", "tribunal", "sanction", "dsa",
        "politique", "onu", "la france", "etat", "trump", "commission européenne",
        "géopolitique", "diplomatie", "cyberattaque", "piratage", "taxe", "présidentielle",
        "amende", "décret", "législation", "démarchage", "fraude", "vie privée",
        "interdiction", "réseaux sociaux", "arnaque", "fuite de données", "régulateur",
        "régulation", "légal", "légalité", "illégal", "anssi", "crypto", "police",
        "données", "cookies", "ministre", "ministère", "usurpation"
    ],
    "17. Actualité éco": [
        "dépense", "bourse", "milliards", "rachat", "capitalisation", "acquisition",
        "argent", "opa", "fusion", "bulle", "usine", "production", "productivité",
        "salaires", "salariés", "employé", "grève", "énergie", "énergétique", "rente",
        "rentable", "perte", "bénéfice", "chiffre d'affaires", "comptable", "crise",
        "pénurie", "inflation", "résultats financiers", "chômage", "licenciement",
        "patron"
    ],
    "18. Sciences": [
        "nasa", "planète", "spatiale", "fusée", "lune", "soleil", "mars", "cnrs",
        "télescope", "l'espace", "satellite"
    ]
}


def clean_url(url):
    """Clean tracking parameters from URLs."""
    parsed = urllib.parse.urlparse(url)
    qd = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    filtered = {k: v for k, v in qd.items() if not k.startswith("utm_")}
    new_query = urllib.parse.urlencode(filtered, doseq=True)
    return urllib.parse.urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path, parsed.params, new_query, parsed.fragment)
    )


def categorize_article(title, url):
    """Classify articles based on title and URL against the CATEGORIES dictionary."""
    text_to_check = f"{title.lower()} {url.lower()}"

    for category, keywords in CATEGORIES.items():
        for keyword in keywords:
            pattern = r'\b' + re.escape(keyword) + r'\b'
            if re.search(pattern, text_to_check):
                return category

    return "NON_CLASSE"


def parse_date(date_str):
    """Parse ISO formatted dates."""
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str)
    except Exception:
        return None


def scrape_rss(source_name, feed_url):
    """Fetch items from a standard RSS feed."""
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
                "title": title,
                "link": link,
                "category": category
            })
    except Exception as e:
        print(f"Error scraping {source_name}: {e}")

    return items


def scrape_frandroid_web(url, source_name):
    """Scrape web pages specifically for Frandroid sections."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
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
                    "title": title,
                    "link": link,
                    "category": category
                })

    except Exception as e:
        print(f"Error scraping Frandroid {source_name}: {e}")

    return items


def load_existing_history(filepath):
    """Load existing articles from CSV to avoid duplicates."""
    history = {}
    if not os.path.exists(filepath):
        return history

    with open(filepath, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            link = row.get("lien")
            if link:
                history[link] = {
                    "source": row.get("source", ""),
                    "date": parse_date(row.get("date")),
                    "title": row.get("titre", ""),
                    "link": link,
                    "category": row.get("categorie", "NON_CLASSE")
                }
    return history


def main():
    os.makedirs("output", exist_ok=True)
    history = load_existing_history(CSV_FILE)
    print(f"Existing articles in history: {len(history)}")

    new_items_count = 0

    # 1. Gather standard RSS feeds
    for source, url in RSS_FEEDS.items():
        print(f"Scraping RSS: {source}...")
        items = scrape_rss(source, url)
        for item in items:
            link = item["link"]
            if link not in history:
                history[link] = item
                new_items_count += 1
            else:
                # Recategorize existing items with the updated dictionary
                history[link]["category"] = categorize_article(history[link]["title"], link)

    # 2. Gather Frandroid Web sections
    frandroid_sections = [
        ("Frandroid - Actualités", "https://www.frandroid.com/actualites"),
        ("Frandroid - Bons plans", "https://www.frandroid.com/bons-plans")
    ]

    for source_name, url in frandroid_sections:
        print(f"Scraping Web: {source_name}...")
        items = scrape_frandroid_web(url, source_name)
        for item in items:
            link = item["link"]
            if link not in history:
                history[link] = item
                new_items_count += 1
            else:
                history[link]["category"] = categorize_article(history[link]["title"], link)

    print(f"New articles added: {new_items_count}")

    # Convert dictionary back to list
    cleaned_list = list(history.values())

    # Sort: Valid dates in descending order on top; items without date at the end
    cleaned_list.sort(
        key=lambda x: (x["date"] is not None, x["date"] or datetime.min.replace(tzinfo=timezone.utc)),
        reverse=True
    )

    # Save to CSV
    with open(CSV_FILE, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(["source", "date", "titre", "lien", "categorie"])

        for item in cleaned_list:
            date_str = item["date"].isoformat() if item["date"] else ""
            writer.writerow([
                item["source"],
                date_str,
                item["title"],
                item["link"],
                item["category"]
            ])

    print("CSV updated successfully!")


if __name__ == "__main__":
    main()
