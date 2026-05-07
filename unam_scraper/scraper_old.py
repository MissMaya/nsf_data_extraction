import requests
from bs4 import BeautifulSoup, Tag
from urllib.parse import urljoin
import pandas as pd
import time
import string
import unicodedata
import urllib3
import re

# --------------------
# config
# --------------------
BASE = "https://www.iifilologicas.unam.mx"
LANDING = BASE + "/dicabenovo/"
POST_URL = BASE + "/dicabenovo/index.php"
OUTFILE = "unam_dicabenovo_scrape.xlsx"

LETTERS = list(string.ascii_uppercase) + ["Ñ"]
TIMEOUT = 60

# --------------------
# SSL (UNAM cert chain)
# --------------------
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

session = requests.Session()
session.verify = False
session.headers.update({"User-Agent": "Mozilla/5.0"})

# --------------------
# superscript map
# --------------------
SUP_MAP = str.maketrans({
    "0":"⁰","1":"¹","2":"²","3":"³","4":"⁴",
    "5":"⁵","6":"⁶","7":"⁷","8":"⁸","9":"⁹",
    "a":"ᵃ","b":"ᵇ","c":"ᶜ","d":"ᵈ","e":"ᵉ",
    "f":"ᶠ","g":"ᵍ","h":"ʰ","i":"ⁱ","j":"ʲ",
    "k":"ᵏ","l":"ˡ","m":"ᵐ","n":"ⁿ","o":"ᵒ",
    "p":"ᵖ","r":"ʳ","s":"ˢ","t":"ᵗ","u":"ᵘ",
    "v":"ᵛ","w":"ʷ","x":"ˣ","y":"ʸ","z":"ᶻ",
    "+":"⁺","-":"⁻","=":"⁼","(":"⁽",")":"⁾"
})

# --------------------
# helpers
# --------------------
def html_to_unicode(html: str) -> str:
    """Convert inner HTML to plain text, turning <sup>..</sup> into Unicode superscripts."""
    soup = BeautifulSoup(html, "html.parser")
    for sup in soup.find_all("sup"):
        sup.replace_with(sup.get_text().translate(SUP_MAP))
    return soup.get_text("", strip=True)

def first_letter_bucket(abbr: str) -> str:
    """
    Decide which letter bucket an abbreviation belongs to.
    - Find first alphabetic char (incl ñ)
    - Strip diacritics (ã -> a) for bucketing only
    - Keep original abbr text unchanged elsewhere
    """
    # find first alphabetic character
    for ch in abbr:
        if ch.isalpha():
            # normalize diacritics for bucketing (ã -> a)
            decomp = unicodedata.normalize("NFD", ch.lower())
            base = "".join(c for c in decomp if unicodedata.category(c) != "Mn")
            if base == "ñ" or ch.lower() == "ñ":
                return "ñ"
            if base and base[0] in string.ascii_lowercase:
                return base[0]
            # fallback: return original lower
            return ch.lower()
    return "?"  # should be rare

FONPAL_RE = re.compile(r"\bfonPalabra\b|\bfonPalabram\b|fonPalabra", re.IGNORECASE)

def parse_entries(soup: BeautifulSoup):
    """
    Parse entries as:
      tituloh1 (abbr)
      then until next tituloh1:
        any span whose class contains 'fonPalabra' -> expansion
        any img tag -> image
    """
    entries = []
    heads = soup.find_all("span", class_="tituloh1")

    for h in heads:
        abbr = html_to_unicode("".join(str(x) for x in h.contents)).strip()
        if not abbr:
            continue

        expansions = []
        images = []

        for node in h.next_elements:
            if not isinstance(node, Tag):
                continue

            # stop at next headword
            if node is not h and node.name == "span" and "tituloh1" in node.get("class", []):
                break

            # expansions: any fonPalabra-ish span
            if node.name == "span":
                cls = " ".join(node.get("class", []))
                if cls and FONPAL_RE.search(cls):
                    txt = html_to_unicode("".join(str(x) for x in node.contents)).strip()
                    if txt:
                        expansions.append(txt)

            # images: any img in the block
            if node.name == "img" and node.get("src"):
                images.append(urljoin(BASE, node["src"]))

        entries.append({
            "abbr": abbr,
            "bucket": first_letter_bucket(abbr),
            "expansions": expansions,
            "images": images,
        })

    return entries

def fetch_post(letter: str):
    payload = {"page": "muestra-lista2", "indice": letter}
    # polite retries in case of 403
    for attempt in range(6):
        r = session.post(POST_URL, data=payload, timeout=TIMEOUT)
        if r.status_code == 403:
            # backoff
            time.sleep(6 + attempt * 4)
            continue
        r.raise_for_status()
        return r
    raise RuntimeError(f"Repeated 403 for letter {letter}")

# --------------------
# STEP 1: parse landing page entries (THIS includes ã. etc.)
# --------------------
print("Fetching landing page…")
r0 = session.get(LANDING, timeout=TIMEOUT)
r0.raise_for_status()

landing_soup = BeautifulSoup(r0.content, "html5lib", from_encoding="latin-1")
landing_entries = parse_entries(landing_soup)

print("Landing entries parsed:", len(landing_entries))
if not landing_entries:
    raise RuntimeError("Landing page parsed 0 entries — check connectivity/HTML parser.")

# Index landing entries by bucket (order preserved)
landing_by_bucket = {}
for e in landing_entries:
    landing_by_bucket.setdefault(e["bucket"], []).append(e)

# Also index landing data by abbr for later merge
landing_data = {e["abbr"]: e for e in landing_entries}

# --------------------
# STEP 2: parse POST pages for each letter (the continuations)
# --------------------
post_by_bucket = {c.lower(): [] for c in string.ascii_uppercase}
post_by_bucket["ñ"] = []

print("Fetching POST pages A–Z + Ñ…")
for L in LETTERS:
    print("  POST", L)
    r = fetch_post(L)
    soup = BeautifulSoup(r.content, "html5lib", from_encoding="latin-1")
    entries = parse_entries(soup)

    # keep only items that truly belong to that bucket (prevents spillover)
    bucket = L.lower()
    entries = [e for e in entries if e["bucket"] == bucket]

    post_by_bucket[bucket].extend(entries)
    time.sleep(2.5)  # crucial for avoiding 403

# --------------------
# STEP 3: build TRUE master order per letter:
#   landing-prefix (in order) then post continuation (in order), deduping by abbr
# --------------------
master_order = []
seen_abbr = set()

for bucket in list(string.ascii_lowercase) + ["ñ"]:
    # landing prefix for this letter
    for e in landing_by_bucket.get(bucket, []):
        if e["abbr"] not in seen_abbr:
            master_order.append(e["abbr"])
            seen_abbr.add(e["abbr"])

    # post continuation for this letter
    for e in post_by_bucket.get(bucket, []):
        if e["abbr"] not in seen_abbr:
            master_order.append(e["abbr"])
            seen_abbr.add(e["abbr"])

print("Master order size (full dictionary):", len(master_order))

# --------------------
# STEP 4: merge data from landing + post into final rows
# --------------------
merged = {}

def merge_into(abbr, expansions, images):
    rec = merged.setdefault(abbr, {"abbr": abbr, "expansions": [], "images": []})
    rec["expansions"].extend(expansions or [])
    rec["images"].extend(images or [])

# landing contributes expansions/images (important!)
for e in landing_entries:
    merge_into(e["abbr"], e["expansions"], e["images"])

# post contributes expansions/images
for bucket, lst in post_by_bucket.items():
    for e in lst:
        merge_into(e["abbr"], e["expansions"], e["images"])

# finalize rows in master order
final_rows = []
for abbr in master_order:
    rec = merged.get(abbr, {"abbr": abbr, "expansions": [], "images": []})
    rec["letter"] = first_letter_bucket(abbr)
    final_rows.append(rec)

# --------------------
# STEP 5: dataframe + expand lists into columns
# --------------------
df = pd.DataFrame(final_rows)

max_exp = int(df["expansions"].apply(len).max() or 0)
for i in range(max_exp):
    df[f"expansion_{i+1}"] = df["expansions"].apply(lambda x: x[i] if len(x) > i else None)
df = df.drop(columns=["expansions"])

max_imgs = int(df["images"].apply(len).max() or 0)
for i in range(max_imgs):
    df[f"image_{i+1}"] = df["images"].apply(lambda x: x[i] if len(x) > i else None)
df = df.drop(columns=["images"])

# --------------------
# STEP 6: write Excel with sensible widths
# --------------------
with pd.ExcelWriter(OUTFILE, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)
    ws = writer.book.active

    for col in ws.columns:
        col_letter = col[0].column_letter
        maxlen = 0
        for cell in col:
            if cell.value is not None:
                maxlen = max(maxlen, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(maxlen + 2, 70)

print("Wrote:", OUTFILE)

# --------------------
# sanity stats
# --------------------
exp_filled = (df.filter(like="expansion_").notna().any(axis=1)).sum() if max_exp else 0
img_filled = (df.filter(like="image_").notna().any(axis=1)).sum() if max_imgs else 0

print("Rows:", len(df))
print("Rows with >=1 expansion:", exp_filled)
print("Rows with >=1 image:", img_filled)
print("First 5 abbreviations:", df["abbr"].head(5).tolist())
