import glob
import re
import string
from bs4 import BeautifulSoup, Tag
from urllib.parse import urljoin
import pandas as pd
import unicodedata
from email import message_from_bytes
from openpyxl.utils import get_column_letter

BASE = "https://www.iifilologicas.unam.mx"
OUTFILE = "unam_dicabenovo_scrape_with_source.xlsx"

# ---------------- superscripts ----------------

# There is no q here because unicode superscript q doesn't exist
# Using a fallback character here for q because it's important to differentiate superscripts

SUP_MAP = str.maketrans({
    "0":"⁰","1":"¹","2":"²","3":"³","4":"⁴",
    "5":"⁵","6":"⁶","7":"⁷","8":"⁸","9":"⁹",
    "a":"ᵃ","b":"ᵇ","c":"ᶜ","d":"ᵈ","e":"ᵉ",
    "f":"ᶠ","g":"ᵍ","h":"ʰ","i":"ⁱ","j":"ʲ",
    "k":"ᵏ","l":"ˡ","m":"ᵐ","n":"ⁿ","o":"ᵒ",
    "p":"ᵖ","q":"ᑫ",  # no real superscript q exists
    "r":"ʳ","s":"ˢ","t":"ᵗ","u":"ᵘ",
    "v":"ᵛ","w":"ʷ","x":"ˣ","y":"ʸ","z":"ᶻ"
})

FONPAL_RE = re.compile("fonPalabra", re.I)

# ---------------- helpers ----------------

def extract_html_from_mhtml(path):
    with open(path, "rb") as f:
        msg = message_from_bytes(f.read())

    for part in msg.walk():
        if part.get_content_type() == "text/html":
            return part.get_payload(decode=True)

    raise RuntimeError(f"No HTML found in {path}")

def html_to_unicode(html):
    soup = BeautifulSoup(html, "html.parser")
    for sup in soup.find_all("sup"):
        sup.replace_with(sup.get_text().translate(SUP_MAP))
    return soup.get_text("", strip=True)

def first_letter_bucket(abbr):
    for ch in abbr:
        if ch.isalpha():
            d = unicodedata.normalize("NFD", ch.lower())
            base = "".join(c for c in d if unicodedata.category(c) != "Mn")
            if base == "ñ":
                return "ñ"
            if base and base[0] in string.ascii_lowercase:
                return base[0]
            return ch.lower()
    return "?"

def parse_entries(soup):
    rows = []

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

            if node is not h and node.name == "span" and "tituloh1" in node.get("class", []):
                break

            if node.name == "span":
                cls = " ".join(node.get("class", []))
                if FONPAL_RE.search(cls):
                    t = html_to_unicode("".join(str(x) for x in node.contents)).strip()
                    if t:
                        expansions.append(t)

            if node.name == "img" and node.get("src"):
                images.append(urljoin(BASE, node["src"]))

        rows.append({
            "abbr": abbr,
            "letter": first_letter_bucket(abbr),
            "source": "UNAM",
            "expansions": expansions,
            "images": images
        })

    return rows

# ---------------- MAIN ----------------

rows = []

files = sorted(glob.glob("html_pages/*.mhtml"))
if not files:
    raise RuntimeError("No mhtml files found")

for fn in files:
    print("Parsing:", fn)
    html = extract_html_from_mhtml(fn)
    soup = BeautifulSoup(html, "html5lib")
    rows.extend(parse_entries(soup))

print("Raw rows:", len(rows))

if not rows:
    raise RuntimeError("ZERO rows parsed — stop")

df = pd.DataFrame(rows)

# expansions

df[["old_spanish", "modern_spanish"]] = pd.DataFrame(
    df["expansions"].tolist(),
    index=df.index
)
df = df.drop(columns=["expansions"])

# images

max_imgs = int(df["images"].apply(len).max() or 0)
for i in range(max_imgs):
    df[f"image_{i+1}"] = df["images"].apply(lambda x: x[i] if len(x) > i else None)
df = df.drop(columns=["images"])

from openpyxl.utils import get_column_letter

with pd.ExcelWriter(OUTFILE, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)
    ws = writer.book.active

    # find image columns
    img_cols = [
        i+1 for i, c in enumerate(df.columns)
        if c.startswith("image_")
    ]

    # convert URLs to real Excel hyperlinks
    for row in ws.iter_rows(min_row=2):
        for idx in img_cols:
            cell = row[idx-1]
            if cell.value:
                url = str(cell.value)
                cell.hyperlink = url
                cell.value = "image"
                cell.style = "Hyperlink"

    # resize columns
    for col in ws.columns:
        col_letter = col[0].column_letter
        maxlen = 0
        for cell in col:
            if cell.value:
                maxlen = max(maxlen, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(maxlen + 2, 60)

print("Wrote:", OUTFILE)
print("Rows:", len(df))
print("First 5:", df["abbr"].head().tolist())
