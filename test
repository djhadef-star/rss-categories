import csv
from datetime import datetime, timezone
import os
import re
import unicodedata
import urllib.parse
from bs4 import BeautifulSoup
import feedparser
import pandas as pd
import requests

# --- CONFIGURATION DES CHEMINS ---
CSV_FILE = os.path.join("output", "rss_history.csv")
EXCEL_FILE = os.path.join("output", "rss_history.xlsx")
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

# --- DICTIONNAIRE MIS À JOUR (20 CATÉGORIES SELON CAPTURE) ---
CATEGORIES = {
    "1. SMARTPHONE": [
        "smartphone",
        "téléphone",
        "phone",
        "l'iphone",
        "iphone",
        "pixel",
        "flip",
        "siri",
        "forfait",
        "android",
        "GrapheneOS",
        "ios",
        "5G",
        "6G",
        "réseau",
        "Esim",
        "traceur",
        "Tag",
        "lunette",
        "appli",
        "application",
        "redmi",
        "whatsapp",
        "sosh",
        "Honor",
        '"One Ui"',
        "poco",
        "realme",
        '"S25"',
        '"S26"',
        '"S27"',
        '"S28"',
        "oneplus",
        "oppo",
        "pliant",
        "pliable",
        "BlackBerry",
        "motorola",
    ],
    "2. PC_OS": [
        "PC",
        "Geekom",
        "Framework",
        "SSD",
        "serveur",
        "Nas",
        "portable",
        "macbook",
        "asus",
        "lenovo",
        "HP",
        "Dell",
        "acer",
        "MSI",
        "ryzen",
        "qualcomm",
        '"Media Tek"',
        "mediatek",
        "snapdragon",
        "amd",
        "nvidia",
        '"puce graphique"',
        '"Galaxy book"',
        '"carte mère"',
        '"cartes mères"',
        "GPU",
        "arm",
        "laptop",
        "linux",
        "Bios",
        "windows",
        "Macos",
        "gnome",
        "fedora",
        "ubuntu",
        "patch",
        "DNS",
    ],
    "3. Periphérique": [
        "Clavier",
        "écran",
        "moniteur",
        "souris",
        "l'imprimante",
        "imprimante",
        "logitech",
        "corsair",
        "steam deck",
        "chargeur",
        "charge",
        '"batterie externe"',
        "ugreen",
        "branche",
        "lexar",
        "périphérique",
        "hub",
        "dock",
        "usb",
        "usb-c",
        "chaise",
        '"RJ 45"',
        "Wifi",
        "wi-fi",
        "routeur",
        "netgear",
        "ethernet",
        "répéteur",
        '"TP-link"',
    ],
    "4. Tablette": [
        "Tablette",
        "Fold",
        "pad",
        "tab",
        "kindle",
        "remarkable",
        "liseuse",
        "kobo",
        "ebooks",
        "pliable",
        "pliure",
        "pli",
        '"Galaxy tab"',
        "livre",
    ],
    "5. Logiciels _ OS": [
        "Logiciels",
        "software",
        '"power Toys"',
        "office",
        "mail",
        "rss",
        "slack",
        "teams",
        "capture",
        "outlook",
        "onenote",
        "excel",
        "word",
        "gmail",
        "drive",
        "cleaner",
        "VPN",
        "traduction",
        "chat",
        "automatisation",
        "edge",
        "chrome",
        "safari",
        "firefox",
        "mozilla",
        "opera",
        "vivaldi",
        "onglet",
        '"moteur de recherche"',
        "github",
        "zip",
        "terminal",
        "torrent",
        "chatbot",
        "interface",
        "wordpress",
        "web",
        "navigateur",
        '"google keep"',
        "virus",
        "bloqueur",
        "expressvpn",
        "cyberghost",
        "surfshark",
        "pcloud",
        "cpanel",
    ],
    "6. IA": [
        '"l\'IA"',
        "IA",
        '"intelligence artificielle"',
        '"data center"',
        "LLM",
        "token",
        "anthropic",
        '"d\'openAI"',
        '"openAI"',
        '"chatgpt"',
        "gemini",
        "meta",
        "Claude",
        "openclaw",
        "perplexity",
        "prompt",
        "copilot",
        "deepseek",
        "mistral",
        "groq",
        "midjourney",
        "grok",
        '"llm"',
        '"agent ai"',
        '"outil ai"',
        '"fonctionnalités ai"',
    ],
    "7. Video": [
        "Iptv",
        "vids",
        "télévision",
        "téléviseur",
        "TV",
        "omni",
        "pics",
        "vlc",
        "video",
        "éditeur",
        "3D",
        "plex",
        "LG",
        "hisense",
        "TCL",
        "Hdmi",
        "RGB",
        "videoprojecteur",
        "Awol",
        "xgimi",
        "Jmgo",
        "stick",
        "DJI",
        "osmo",
        "Insta360",
        "cast",
        "graphique",
        "édition",
        "image",
        "gopro",
        "affichage",
    ],
    "8. Audio": [
        "casque",
        "audio",
        '"casque audio"',
        '"casque gaming"',
        "pods",
        "buds",
        "headphone",
        "jbl",
        "shokz",
        "clip",
        "écouteurs",
        '"réduction de bruit"',
        "enceinte",
        '"barre de son"',
        "harman",
        "sennheiser",
        "bose",
        "sonos",
        "égaliseur",
        "musique",
        "voix",
        "platine",
        "vinyle",
        "oreille",
        "openfit",
        "aeroclip",
        "chanson",
    ],
    "9. Moiilité douce": [
        "mobilité",
        "vélo",
        "trottinette",
        "cargo",
        "sacoche",
        "antivol",
        "VAE",
        "navigo",
        "blablacar",
        "taxi",
        "rer",
        '"transport en commun"',
        "ebike",
        "vtt",
        "segway",
        "Navee",
        "veste",
        "avion",
        "aéroport",
        "GPS",
        "pompe",
        '"compteur vélo"',
    ],
    "10. Voiture": [
        "voiture",
        "automobile",
        "conduite",
        "conduire",
        "rouler",
        "essence",
        "carburant",
        "monospace",
        "berline",
        "citadine",
        "SUV",
        "permis",
        "volkswagen",
        "mercedes",
        "bmw",
        "byd",
        "mg",
        "renault",
        "peugeot",
        "toyota",
        "nissan",
        "stellantis",
        "citroen",
        "dacia",
        "kia",
        "fiat",
        "tesla",
        "ford",
        "jeep",
        "twingo",
        "volvo",
        "autonome",
        "FSD",
        "waze",
        '"google maps"',
        "auto",
        "carplay",
        "taxe",
        "moto",
        "scooter",
        "hybride",
        '"voiture électrique"',
        "borne",
        "véhicule",
        "recharge",
        "robotaxi",
        "porshe",
        "CATL",
    ],
    "11. Menage": [
        "ménage",
        "aspirateur",
        '"robot aspirateur"',
        '"aspirateurs robots"',
        '"aspirateur robot"',
        "dyson",
        "roborock",
        "narwal",
        "mova",
        "ecovacs",
        "dreame",
        "tineco",
        "shark",
        "laveur",
        "lavage",
        '"machine à laver"',
        "vitres",
        "pressing",
        "défroisseur",
        '"centrale vapeur"',
        '"fer à repasser"',
        "tambour",
        "maison",
        "linge",
        "vaisselle",
        '"lave-vaisselle"',
        "vitre",
        '"brosse à dent"',
        "vapeur",
        "ventilateur",
        "nettoyeur",
        "nettoyant",
        "purificateur",
        "serpillère",
    ],
    "12. Cuisine": [
        "Airfryer",
        "ninja",
        "moulinex",
        "thermomix",
        "cookeo",
        "coffee",
        "delonghi",
        "tefal",
        "cuisine",
        "plaque",
        "four",
        "cafetière",
        "café",
        "expresso",
        "cappuccino",
        "barista",
        "barbecue",
        "kenwood",
        "soda",
        "agrume",
        "frigo",
        "réfrigérateur",
        "congélateur",
        "broyeur",
        "krups",
        '"cuisson"',
        '"robot cuiseur"',
        "vin",
        "bouilloire",
    ],
    "13. Maison": [
        "box",
        "hue",
        "ikea",
        "freebox",
        "livebox",
        "home",
        "matter",
        "fibre",
        "caméra",
        "surveillance",
        "serrure",
        "nuki",
        "sonnette",
        "keypad",
        "clé",
        "thermostat",
        "linky",
        "tuya",
        "aqara",
        "prise",
        "interrupteurs",
    ],
    "14. Extérieur": [
        "solaire",
        "électricité",
        "compteur",
        "gaz",
        "énergie",
        "EDF",
        "solar",
        "starlink",
        "plante",
        "jardin",
        "chauffage",
        "chaudière",
        "climatiseur",
        '"robot tondeuse"',
    ],
    "15. Santé": [
        "Santé",
        "fitbit",
        "bracelet",
        "circa",
        "montre",
        "watch",
        "sommeil",
        "dormir",
        "ondes",
        "withings",
        "peau",
        "sport",
        "course",
        "natation",
        "coros",
        "coach",
        "outdoor",
        "garmin",
        "polar",
        "amazfit",
        "suunto",
        "muscle",
        "musculation",
        "randonnée",
        "band",
        "whoop",
        "ecg",
        "artérielle",
        "bague",
        "ring",
        "rasoir",
    ],
    "16. ludique": [
        "jeu",
        "jeux",
        "ludique",
        '"jeux de société"',
        '"jeux vidéo"',
        '"jeu vidéo"',
        "enfant",
        "cinéma",
        "film",
        '"court métrage"',
        '"long métrage"',
        "switch",
        "playstation",
        "xbox",
        "manette",
        "VR",
        '"casque VR"',
        "netflix",
        "canal",
        "TF1",
        "cinema",
        "série",
        "ps6",
        "ps5",
        "ps4",
        "nintendo",
        "steam",
        "gaming",
        "spotify",
        "valve",
        "zelda",
        "mario",
        '"super mario"',
        "mariokart",
        "Pokemon",
        "ubisoft",
        "console",
        "drone",
        '"neo geo"',
        "sega",
        "GTA",
        "streaming",
        "media",
        "youtube",
        "Disney+",
        "Disney",
        '"game boy"',
        "lego",
        '"hbo max"',
        '"apple TV"',
        "paramount",
        "saison",
        '"chef d\'oeuvre"',
        '"chef-d\'oeuvre"',
        "roman",
        "dessin",
        "dessine",
        "bande-dessinée",
        "manga",
        "teaser",
        "casting",
        "scène",
        "personnage",
        "ligue 1",
        '"ligue des champions"',
        '"coupe d\'europe"',
        "FIFA",
        "UEFA",
        "football",
        "tennis",
        "basket",
        "rugby",
        "handball",
    ],
    "17. actu jur et pol": [
        "UE",
        "Cnil",
        "souveraineté",
        "europe",
        "Etats-unis",
        "chine",
        "russie",
        "Iran",
        "condamnée",
        "condamné",
        "condamne",
        "gafam",
        '"conseil d\'état"',
        "arcom",
        "RGPD",
        '"IA Act"',
        "Bruxelles",
        '"union européenne"',
        "Otan",
        "procès",
        "plaintes",
        "règles",
        "réglementation",
        "loi",
        "juge",
        "justice",
        "tribunal",
        "sanction",
        "DSA",
        "politique",
        "onu",
        '"la France"',
        "Etat",
        "Trump",
        '"commission européenne"',
        "géopolitique",
        "diplomatie",
        "cyberattaque",
        "piratage",
        "taxe",
        "présidentielle",
        "amende",
        "décret",
        "législation",
        "démarchage",
        "fraude",
        '"vie privée"',
        "interdiction",
        '"réseaux sociaux"',
        "arnaque",
        '"fuite de données"',
        "régulateur",
        "régulation",
        "légal",
        "légalité",
        "illégal",
        "illégale",
        "ANSSI",
        "crypto",
        "police",
        "données",
        "coookies",
        "ministre",
        "ministère",
        "usurpation",
        "guerre",
        "impôts",
        "fisc",
        "hacker",
        "conforme",
    ],
    "18. Actualité éco": [
        "dépense",
        "bourse",
        "milliards",
        "rachat",
        "capitalisation",
        "acquisition",
        "argent",
        "OPA",
        "fusion",
        "bulle",
        "usine",
        "production",
        "productivité",
        "salaires",
        "salariés",
        "employé",
        "grève",
        "énergie",
        "énergétique",
        "rente",
        "rentable",
        "perte",
        "bénéfice",
        '"chiffre d\'affaires"',
        "comptable",
        "crise",
        "pénurie",
        "inflation",
        '"résultats financiers"',
        "chômage",
        "licenciement",
        "patron",
    ],
    "19. Sciences": [
        "Nasa",
        "planète",
        "spatiale",
        "fusée",
        "lune",
        "soleil",
        "mars",
        "CNRS",
        "téléscope",
        '"l\'espace"',
        "satellite",
        '"Système solaire"',
        "chercheur",
        "nucléaire",
        "astéroïde",
        "cratère",
        "iss",
        '"l\'iss"',
        "spacex",
        "lunaire",
        "implant",
        "cerveau",
        "astronome",
        "astronaute",
    ],
    "20. Sponso": [
        '"bon plan"',
        '"bons plans"',
        "bonplan",
        "bon-plan",
        "bons-plans",
        "brade",
        "promotion",
        "sponsorisé",
        "Aliexpress",
        "promo",
        "aliexpress",
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
    """Construit une regex gérant les blocs entre guillemets et les pluriels (s/x)
    sans capturer les verbes ou mots dérivés plus longs.
    """
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
    """Classe l'article selon tes règles strictes :
    1. Priorité ABSOLUE à la catégorie Sponso (Titre + URL).
    2. Pour les autres catégories : recherche exclusivement dans le Titre,
       avec priorité au mot-clé / bloc qui apparaît le plus tôt.
    """
    text_title = remove_accents(title)
    text_url = remove_accents(url)

    # --- 1. RÈGLE SPONSO : PRIORITÉ ABSOLUE (Titre + URL) ---
    sponso_category = next(
        (c for c in CATEGORIES if "sponso" in c.lower()), None
    )
    if sponso_category:
        full_text_sponso = f"{text_title} {text_url}"
        for keyword in CATEGORIES[sponso_category]:
            pattern = build_regex_pattern(keyword)
            if pattern and re.search(pattern, full_text_sponso):
                return sponso_category

    # --- 2. CATÉGORIES STANDARD : Index du premier mot/bloc dans le Titre uniquement ---
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
    """Parse les dates ISO et assure une compatibilité des fuseaux horaires (UTC)."""
    if not date_str:
        return None
    try:
        dt = datetime.fromisoformat(date_str)
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None


# --- MODULES DE SCRAPING ---
def scrape_rss(source_name: str, feed_url: str) -> list:
    """Scrape un flux RSS standard."""
    items = []
    try:
        feed = feedparser.parse(feed_url)
        for entry in feed.entries:
            title = entry.get("title", "").strip()
            link = clean_url(entry.get("link", "").strip())
            if not title or not link:
                continue

            dt_tuple = entry.get("published_parsed") or entry.get("updated_parsed")
            dt = (
                datetime(*dt_tuple[:6], tzinfo=timezone.utc) if dt_tuple else None
            )

            items.append({
                "source": source_name,
                "date": dt,
                "titre": title,
                "lien": link,
                "categorie": categorize_article(title, link),
            })
    except Exception as e:
        print(f"Erreur RSS {source_name}: {e}")
    return items


def scrape_frandroid_web(url: str, source_name: str) -> list:
    """Scrape spécifiquement la structure HTML des pages Web de Frandroid."""
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
                "categorie": categorize_article(title, link),
            })
    except Exception as e:
        print(f"Erreur Scraping Web {source_name}: {e}")
    return items


# --- HISTORIQUE & EXPORT ---
def load_existing_history(filepath: str) -> dict:
    """Charge le fichier CSV d'historique s'il existe pour éviter la perte de données."""
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
                    "categorie": row.get("categorie", "NON_CLASSE"),
                }
    return history


def generate_excel_from_csv():
    """Génère le fichier Excel multi-onglets (1 global + catégories + 1 non-classé)."""
    if not os.path.exists(CSV_FILE):
        print("CSV introuvable. Annulation de la génération Excel.")
        return

    df_all = pd.read_csv(CSV_FILE, delimiter=";")

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        # 1. Onglet Tous les articles
        df_all.to_excel(writer, sheet_name="Tous_les_articles", index=False)

        # 2. Onglets par catégorie
        for cat_name in CATEGORIES.keys():
            sheet_title = cat_name[:31]  # Contrainte max d'Excel (31 caractères)
            df_filtered = df_all[df_all["categorie"] == cat_name]
            df_filtered.to_excel(writer, sheet_name=sheet_title, index=False)

        # 3. Onglet Articles non classés
        df_unclassed = df_all[df_all["categorie"] == "NON_CLASSE"]
        if not df_unclassed.empty:
            df_unclassed.to_excel(writer, sheet_name="NON_CLASSE", index=False)

    print(f"Fichier Excel mis à jour : '{EXCEL_FILE}' !")


# --- MAIN PIPELINE ---
def main():
    os.makedirs("output", exist_ok=True)

    # 1. Chargement de l'historique existant
    history = load_existing_history(CSV_FILE)
    print(f"Articles dans l'historique initial : {len(history)}")

    # 2. Re-catégorisation systématique de l'historique
    for link, item in history.items():
        item["categorie"] = categorize_article(item["titre"], link)

    new_items_count = 0

    # 3. Scraping des flux RSS
    for source, url in RSS_FEEDS.items():
        print(f"Scraping RSS : {source}...")
        items = scrape_rss(source, url)
        for item in items:
            link = item["lien"]
            if link not in history:
                history[link] = item
                new_items_count += 1
            else:
                history[link]["categorie"] = categorize_article(
                    history[link]["titre"], link
                )

    # 4. Scraping Web Frandroid (Actualités & Bons plans)
    frandroid_sections = [
        ("Frandroid - Actualités", "https://www.frandroid.com/actualites"),
        ("Frandroid - Bons plans", "https://www.frandroid.com/bons-plans"),
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

    # 6. Écriture du fichier CSV mis à jour
    with open(CSV_FILE, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["source", "date", "titre", "lien", "categorie"])
        for item in cleaned_list:
            date_str = item["date"].isoformat() if item["date"] else ""
            writer.writerow(
                [item["source"], date_str, item["titre"], item["lien"], item["categorie"]]
            )

    print("CSV d'historique mis à jour avec succès !")

    # 7. Génération du fichier Excel multi-onglets
    generate_excel_from_csv()


if __name__ == "__main__":
    main()
