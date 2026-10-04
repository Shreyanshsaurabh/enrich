"""
Developer pipeline:
  1. Keep only apps that contain ads
  2. Deduplicate by Developer ID
       - Rating     -> average across the developer's apps
       - Downloads  -> sum across the developer's apps
       - Genre, Category, App Name, App Link, Privacy Policy -> taken from the
         developer's app with the most downloads
  3. Visit each developer's privacy policy page and extract privacy_email / privacy_number
  4. Write the result to an Excel file

Install:
    pip install pandas openpyxl requests beautifulsoup4 phonenumbers

Usage:
    python developer_pipeline.py all_apps.xlsx developers_enriched.xlsx
    python developer_pipeline.py all_apps.xlsx out.xlsx --workers 16
    python developer_pipeline.py all_apps.xlsx out.xlsx --limit 200      # test on 200 developers
    python developer_pipeline.py all_apps.xlsx out.xlsx --no-scrape      # steps 1-2 only

Scraping results are cached in <output>.cache.json, so if the run stops you can
run the same command again and it resumes where it left off.
"""
import argparse
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import pandas as pd
import phonenumbers
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; PrivacyContactBot/1.0)"}
TIMEOUT = 15
MAX_BYTES = 2_000_000
HOST_DELAY = 1.0          # min seconds between requests to the same host
MAX_EMAILS = 3
MAX_PHONES = 3

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*\.[a-zA-Z]{2,}")
OBFUSCATED_RE = re.compile(
    r"([a-zA-Z0-9._%+-]+)\s*[\[\(]\s*(?:at|AT)\s*[\]\)]\s*"
    r"([a-zA-Z0-9-]+(?:\s*[\[\(]?\s*(?:dot|DOT)\s*[\]\)]?\s*[a-zA-Z0-9-]+)+)")
JUNK_EMAIL = re.compile(
    r"\.(png|jpe?g|gif|svg|webp|css|js)$|sentry|wixpress|example\.(com|org)|"
    r"yourcompany|yourdomain|domain\.com|email\.com$|^user@|^name@|^you@", re.I)
PRIVACY_WORDS = ("privacy", "dpo", "gdpr", "data", "legal", "dataprotection")


# ---------------------------------------------------------------- cleaning ---
def to_bool(v):
    return str(v).strip().lower() in {"true", "yes", "y", "1", "1.0"}


def parse_downloads(v):
    """'10,000,000+' -> 10000000 ; '5M+' -> 5000000 ; 1200 -> 1200"""
    if pd.isna(v):
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper().replace(",", "").replace("+", "").replace(" ", "")
    mult = {"K": 1e3, "M": 1e6, "B": 1e9}.get(s[-1:], 1)
    if mult != 1:
        s = s[:-1]
    try:
        return float(s) * mult
    except ValueError:
        return 0


def clean_and_dedupe(df):
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Developer ID", "Rating", "Downloads", "Contains Ads", "Genre",
                "Category", "App Name", "App Link", "Privacy Policy"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f"Missing columns in input file: {missing}")

    total = len(df)
    df = df[df["Contains Ads"].apply(to_bool)].copy()
    print(f"[1] Kept {len(df):,} of {total:,} apps that contain ads")

    # fall back to developer name if the ID is blank
    if "Developer Name" in df.columns:
        df["Developer ID"] = df["Developer ID"].fillna(df["Developer Name"])
    df = df.dropna(subset=["Developer ID"])

    df["_downloads"] = df["Downloads"].apply(parse_downloads)
    df["Rating"] = pd.to_numeric(df["Rating"], errors="coerce")

    # per-developer aggregates
    agg = df.groupby("Developer ID").agg(
        Avg_Rating=("Rating", "mean"),
        Total_Downloads=("_downloads", "sum"),
        Apps_Count=("App Name", "count"),
    )

    # row of the highest-downloaded app per developer
    top = df.sort_values("_downloads", ascending=False).drop_duplicates("Developer ID", keep="first")
    top = top.set_index("Developer ID")

    out = top.drop(columns=["Rating", "Downloads", "_downloads"]).join(agg)
    out["Avg_Rating"] = out["Avg_Rating"].round(2)
    out["Total_Downloads"] = out["Total_Downloads"].astype("int64")
    out = out.rename(columns={"Avg_Rating": "Avg Rating",
                              "Total_Downloads": "Total Downloads",
                              "Apps_Count": "Apps Count"})
    out = out.reset_index()
    out = out.sort_values("Total Downloads", ascending=False).reset_index(drop=True)
    print(f"[2] {len(out):,} unique developers after deduplication")
    return out


# ---------------------------------------------------------------- scraping ---
_host_lock = threading.Lock()
_host_last = {}


def polite_wait(host):
    with _host_lock:
        now = time.time()
        wait = _host_last.get(host, 0) + HOST_DELAY - now
        _host_last[host] = max(now, _host_last.get(host, 0) + HOST_DELAY)
    if wait > 0:
        time.sleep(wait)


def fetch(url):
    """Returns (html_or_None, status_string)."""
    host = urlparse(url).netloc
    polite_wait(host)
    try:
        with requests.get(url, headers=HEADERS, timeout=TIMEOUT, stream=True) as r:
            if r.status_code >= 400:
                return None, f"http_{r.status_code}"
            ctype = r.headers.get("content-type", "").lower()
            if "pdf" in ctype or url.lower().endswith(".pdf"):
                return None, "pdf_skipped"
            raw = r.raw.read(MAX_BYTES, decode_content=True)
            r.encoding = r.encoding or "utf-8"
            return raw.decode(r.encoding, errors="ignore"), "ok"
    except requests.RequestException as e:
        return None, f"fetch_failed_{type(e).__name__}"


def region_from_country(country):
    c = str(country or "").strip().upper()
    return c if len(c) == 2 and c.isalpha() else None


def rank_email(email, page_host, dev_email_domain):
    local, domain = email.split("@")
    score = 0
    if any(w in local for w in PRIVACY_WORDS):
        score += 100
    if dev_email_domain and domain == dev_email_domain:
        score += 20
    if page_host and (domain in page_host or page_host.endswith(domain)):
        score += 10
    if local in {"noreply", "no-reply", "donotreply"}:
        score -= 100
    return -score


def extract_contacts(html, page_url, region, dev_email):
    soup = BeautifulSoup(html, "html.parser")
    emails, phones = set(), []

    for a in soup.select('a[href^="mailto:"]'):
        emails.add(a["href"][7:].split("?")[0].strip().lower())
    for a in soup.select('a[href^="tel:"]'):
        try:
            n = phonenumbers.parse(a["href"][4:], region)
            if phonenumbers.is_valid_number(n):
                phones.append(phonenumbers.format_number(n, phonenumbers.PhoneNumberFormat.INTERNATIONAL))
        except phonenumbers.NumberParseException:
            pass

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text(" ")

    emails |= {e.lower() for e in EMAIL_RE.findall(text)}
    for user, host in OBFUSCATED_RE.findall(text):
        host = re.sub(r"\s*[\[\(]?\s*(?:dot|DOT)\s*[\]\)]?\s*", ".", host).replace(" ", "")
        emails.add(f"{user}@{host}".lower())
    emails = {e.strip(".,;:") for e in emails if not JUNK_EMAIL.search(e)}

    for m in phonenumbers.PhoneNumberMatcher(text, region, leniency=phonenumbers.Leniency.VALID):
        phones.append(phonenumbers.format_number(m.number, phonenumbers.PhoneNumberFormat.INTERNATIONAL))

    page_host = urlparse(page_url).netloc.lower().replace("www.", "")
    dev_domain = dev_email.split("@")[1].lower() if isinstance(dev_email, str) and "@" in dev_email else None
    ranked = sorted(emails, key=lambda e: rank_email(e, page_host, dev_domain))
    phones = list(dict.fromkeys(phones))        # dedupe, keep order
    return ranked[:MAX_EMAILS], phones[:MAX_PHONES]


def scrape_one(url, country, dev_email):
    if not isinstance(url, str) or not url.strip().lower().startswith("http"):
        return {"privacy_email": "", "privacy_number": "", "privacy_status": "no_privacy_url"}
    html, status = fetch(url.strip())
    if html is None:
        return {"privacy_email": "", "privacy_number": "", "privacy_status": status}
    emails, phones = extract_contacts(html, url, region_from_country(country), dev_email)
    return {
        "privacy_email": "; ".join(emails),
        "privacy_number": "; ".join(phones),
        "privacy_status": "ok" if (emails or phones) else "no_contact_found",
    }


def load_cache(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(path, cache):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    os.replace(tmp, path)


def enrich_privacy(df, cache_path, workers):
    cache = load_cache(cache_path)
    email_col = "Developer Email" if "Developer Email" in df.columns else None
    country_col = "Country" if "Country" in df.columns else None

    jobs = {}
    for _, row in df.iterrows():
        url = row["Privacy Policy"]
        if isinstance(url, str) and url.strip() and url.strip() not in cache and url.strip() not in jobs:
            jobs[url.strip()] = (row[country_col] if country_col else None,
                                 row[email_col] if email_col else None)
    print(f"[3] Scraping {len(jobs):,} privacy pages ({len(cache):,} already cached)")

    done = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(scrape_one, u, c, e): u for u, (c, e) in jobs.items()}
        for fut in as_completed(futures):
            url = futures[fut]
            try:
                cache[url] = fut.result()
            except Exception as e:  # never let one bad page kill the run
                cache[url] = {"privacy_email": "", "privacy_number": "",
                              "privacy_status": f"error_{type(e).__name__}"}
            done += 1
            if done % 200 == 0:
                save_cache(cache_path, cache)
                print(f"    {done:,}/{len(jobs):,} done")
    save_cache(cache_path, cache)

    empty = {"privacy_email": "", "privacy_number": "", "privacy_status": "no_privacy_url"}
    res = df["Privacy Policy"].apply(lambda u: cache.get(str(u).strip(), empty) if isinstance(u, str) else empty)
    for col in ("privacy_email", "privacy_number", "privacy_status"):
        df[col] = res.apply(lambda d: d[col])
    return df


# ------------------------------------------------------------------ output ---
def write_excel(df, path):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Developers")
        ws = xw.sheets["Developers"]
        head_fill = PatternFill("solid", fgColor="1F3864")
        for c in ws[1]:
            c.font = Font(name="Arial", bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
        for i, col in enumerate(df.columns, start=1):
            width = min(max(len(str(col)), int(df[col].astype(str).str.len().quantile(0.9))) + 2, 50)
            ws.column_dimensions[get_column_letter(i)].width = max(width, 12)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        if "Total Downloads" in df.columns:
            col = get_column_letter(list(df.columns).index("Total Downloads") + 1)
            for c in ws[col][1:]:
                c.number_format = "#,##0"
    print(f"[4] Saved {path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--workers", type=int, default=12)
    ap.add_argument("--limit", type=int, default=0, help="only process the top N developers (testing)")
    ap.add_argument("--no-scrape", action="store_true")
    args = ap.parse_args()

    df = pd.read_excel(args.input)
    df = clean_and_dedupe(df)
    if args.limit:
        df = df.head(args.limit)
    if not args.no_scrape:
        df = enrich_privacy(df, args.output + ".cache.json", args.workers)
        print(df["privacy_status"].value_counts().to_string())
    write_excel(df, args.output)


if __name__ == "__main__":
    main()
