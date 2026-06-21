import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------

BASE_URL = "https://www.myjobmag.co.za"

SCRAPE_PAGES = int(os.environ.get("SCRAPE_PAGES", "3"))      # how many listing pages to crawl
REQUEST_DELAY = float(os.environ.get("REQUEST_DELAY", "1.0"))  # polite delay between requests, seconds
MAX_JOBS = int(os.environ.get("MAX_JOBS", "0"))               # 0 = no cap, otherwise stop after N jobs printed

# Apply-link resolution: follow the myjobmag.co.za/apply-now/<id> redirect to
# find the real destination. Each resolution is an extra HTTP request, so it
# can be turned off (e.g. for a quick test run) by setting RESOLVE_APPLY_URLS=0.
RESOLVE_APPLY_URLS = os.environ.get("RESOLVE_APPLY_URLS", "1") != "0"
RESOLVE_DELAY = float(os.environ.get("RESOLVE_DELAY", "0.5"))  # polite delay between resolve requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Charset": "utf-8",
}

REQUEST_TIMEOUT = 25

# Reuse one TCP/TLS connection where possible for every request this run makes.
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# Caches the raw apply-now URL -> resolved real URL, so if the same apply-now
# link is ever seen twice we don't hit it again.
_apply_url_cache = {}

# Generic "collection" titles we don't want to treat as a real job title.
# Only used in the single-job fallback parser, where the title comes from <h1>.
GENERIC_TITLE_PATTERNS = [
    re.compile(r"^fresh jobs? at ", re.I),
    re.compile(r"^latest (jobs?|recruitment|openings?) at ", re.I),
    re.compile(r"^(job )?(openings?|opportunities|vacancies|positions?) at ", re.I),
    re.compile(r"^top (positions|roles) at ", re.I),
    re.compile(r"^open roles? at ", re.I),
    re.compile(r"^careers? at ", re.I),
    re.compile(r"^(excellent|new) career openings? at ", re.I),
    re.compile(r"^(new )?recruitment at ", re.I),
    re.compile(r"^(trending|active|current|hot) (jobs?|roles?|openings?|vacancies) at ", re.I),
]

# Matches a plain email address inside free text job descriptions.
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")

# Matches <meta http-equiv="refresh" content="0;url=https://...">
META_REFRESH_PATTERN = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]*content=["\'][^"\'>]*url=([^"\'>]+)',
    re.I,
)

# Matches common JS redirect idioms:
#   window.location = "..."   window.location.href = "..."
#   window.location.replace("...")   location.replace("...")
JS_REDIRECT_PATTERN = re.compile(
    r'(?:window\.)?location(?:\.href)?\s*(?:=\s*|\.replace\(\s*)["\']([^"\']+)["\']',
    re.I,
)


# -----------------------------------------------------------------------------
# HELPERS
# -----------------------------------------------------------------------------

def log(msg):
    print(msg, flush=True)


def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    return BeautifulSoup(resp.text, "lxml")


def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()


# UI boilerplate phrases that leak into job-details text and aren't part of
# the actual job content. Stripped before printing.
BOILERPLATE_PATTERNS = [
    re.compile(r"go to method of application\s*[»>]*", re.I),
    re.compile(r"Read more about this company", re.I),
]


def clean_description(text):
    if not text:
        return text
    for pattern in BOILERPLATE_PATTERNS:
        text = pattern.sub("", text)
    return re.sub(r"\s+", " ", text).strip()


def strip_company_suffix(title):
    """Strip ' at Company Name' suffix, used only by the single-job fallback parser."""
    if not title:
        return title
    m = re.match(r"^(.+?)\s+at\s+.+$", title, re.I)
    return m.group(1).strip() if m else title.strip()


def is_generic_title(title):
    return any(p.search(title) for p in GENERIC_TITLE_PATTERNS)


def parse_posted_date(date_str):
    """Parses strings like 'Jun 20, 2026' -> datetime, or None on failure."""
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str.strip(), "%b %d, %Y")
    except ValueError:
        return None


def add_three_months(dt):
    if not dt:
        return ""
    month = dt.month - 1 + 3
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, 28)  # safe day to avoid month-length issues
    return datetime(year, month, day).strftime("%Y-%m-%d")


def absolute_url(href):
    if not href:
        return ""
    return href if href.startswith("http") else BASE_URL + href


def extract_email(text):
    """Returns the first email address found in the given text, or ''."""
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""


def _same_site(url):
    try:
        return urlparse(url).netloc == urlparse(BASE_URL).netloc
    except Exception:
        return False


def _find_html_redirect_target(html, current_url):
    """Looks for a meta-refresh or simple JS redirect inside an HTML page and
    returns the absolute target URL, or '' if none is found."""
    m = META_REFRESH_PATTERN.search(html)
    if not m:
        m = JS_REDIRECT_PATTERN.search(html)
    if not m:
        return ""
    target = m.group(1).strip().strip("'\"")
    return urljoin(current_url, target)


def resolve_apply_url(raw_apply_url):
    """Follows a myjobmag.co.za/apply-now/<id> link to find the real employer
    application page. Returns the resolved URL, or '' if it can't be resolved
    (no link, request failure, or it never leaves myjobmag.co.za)."""
    if not raw_apply_url:
        return ""

    if raw_apply_url in _apply_url_cache:
        return _apply_url_cache[raw_apply_url]

    resolved = ""
    try:
        resp = SESSION.get(raw_apply_url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        final_url = resp.url

        if final_url and not _same_site(final_url):
            # Plain HTTP redirect already landed us off myjobmag.co.za - done.
            resolved = final_url
        else:
            # Either no redirect happened, or it redirected to another page
            # still on myjobmag.co.za. Check the HTML for a meta-refresh or
            # JS redirect that requests.get() wouldn't have followed.
            html_target = _find_html_redirect_target(resp.text, final_url)
            if html_target:
                resolved = html_target
            elif final_url and final_url != raw_apply_url:
                # It moved somewhere, even if still on-site - report it rather
                # than silently dropping it, since it's still more useful
                # than the bare tracking link.
                resolved = final_url
            else:
                resolved = ""
    except requests.RequestException as e:
        log(f"    WARNING: could not resolve apply URL {raw_apply_url}: {e}")
        resolved = ""

    _apply_url_cache[raw_apply_url] = resolved
    if RESOLVE_DELAY:
        time.sleep(RESOLVE_DELAY)
    return resolved


def resolve_application_contact(raw_apply_url, description):
    """Returns a dict with the best application info we can find:
      {'apply_url': <resolved real link or ''>,
       'apply_email': <email scraped from description or ''>,
       'apply_raw': <original myjobmag.co.za apply-now link, for reference>}
    Resolution order: real redirect target first, then an email address found
    in the job description, otherwise both come back empty."""
    result = {"apply_url": "", "apply_email": "", "apply_raw": raw_apply_url}

    if raw_apply_url and RESOLVE_APPLY_URLS:
        result["apply_url"] = resolve_apply_url(raw_apply_url)

    if not result["apply_url"]:
        result["apply_email"] = extract_email(description)

    return result


# -----------------------------------------------------------------------------
# STEP 1 — COLLECT COMPANY/JOB PAGE URLS FROM LISTING PAGES
# -----------------------------------------------------------------------------

def collect_company_page_urls(pages=SCRAPE_PAGES):
    urls = []
    seen = set()

    for i in range(1, pages + 1):
        page_url = BASE_URL if i == 1 else f"{BASE_URL}/page/{i}"
        log(f"\n{'=' * 80}\nFETCHING LISTING PAGE {i}: {page_url}\n{'=' * 80}")

        try:
            soup = get_soup(page_url)
        except Exception as e:
            log(f"  ERROR fetching listing page {i}: {e}")
            continue

        blocks = soup.select("li.job-list-li")
        log(f"  Found {len(blocks)} company/job blocks on this page")

        for block in blocks:
            h2_a = block.select_one("li.job-info h2 a")
            if not h2_a or not h2_a.get("href"):
                continue
            full_url = absolute_url(h2_a["href"])
            if full_url not in seen:
                seen.add(full_url)
                urls.append(full_url)

        time.sleep(REQUEST_DELAY)

    log(f"\nTotal unique company/job page URLs collected: {len(urls)}")
    return urls


# -----------------------------------------------------------------------------
# STEP 2 — PARSE A COMPANY/JOB PAGE (handles both multi-job and single-job
#           templates by looking for ul.job-key-info blocks generically)
# -----------------------------------------------------------------------------

def extract_company_blurb(printable, first_h2):
    """Best-effort extraction of the intro company description text that sits
    above the first job block inside #printable."""
    if printable is None:
        return ""
    texts = []
    for el in printable.contents:
        if first_h2 is not None and el is first_h2:
            break
        name = getattr(el, "name", None)
        if name == "div" and el.get("id") == "adbox":
            continue
        if name == "a" and "view-all2" in (el.get("class") or []):
            continue
        if name == "ul" and "table-of-content" in (el.get("class") or []):
            continue
        if name is None:
            txt = str(el).strip()
        else:
            txt = el.get_text(" ", strip=True)
        if txt:
            texts.append(txt)
    return re.sub(r"\s+", " ", " ".join(texts)).strip()


def parse_job_page(url):
    """Returns a list of job dicts scraped from a single company/job page URL.
    Handles the confirmed multi-job template (multiple <h2 id="jobNNN">
    blocks) and falls back to a single-job interpretation if no such blocks
    are found."""

    soup = get_soup(url)
    printable = soup.select_one("#printable")

    h1 = soup.select_one("div.read-left-section h1") or soup.select_one("h1")
    page_title_raw = clean_text(h1)

    company_a = soup.select_one("li.job-industry a[href^='/jobs-at/']")
    company_name = clean_text(company_a) if company_a else ""
    company_name = re.sub(r"^View Jobs at\s*", "", company_name, flags=re.I).strip()
    company_url = absolute_url(company_a["href"]) if company_a and company_a.get("href") else ""

    posted_date_raw = clean_text(soup.select_one("#posted-date"))
    posted_date_raw = re.sub(r"^Posted:\s*", "", posted_date_raw, flags=re.I).strip()
    posted_dt = parse_posted_date(posted_date_raw)
    estimated_deadline = add_three_months(posted_dt)

    deadline = ""
    for li in soup.select(".read-date-sec-li"):
        txt = li.get_text(" ", strip=True)
        if "Deadline" in txt:
            deadline = re.sub(r"^.*Deadline:\s*", "", txt).strip()
    if not deadline or deadline.lower() == "not specified":
        deadline = estimated_deadline or deadline

    apply_links = {}
    for a in soup.select(".application-links a"):
        href = a.get("href", "")
        m = re.search(r"/apply-now/(\d+)", href)
        if m:
            apply_links[m.group(1)] = absolute_url(href)
    single_apply_fallback = next(iter(apply_links.values()), "")

    keyinfo_blocks = (printable.select("ul.job-key-info") if printable else [])

    company_blurb = ""
    job_h2s = (printable.select("h2[id^='job']") if printable else [])
    job_h2s = [h for h in job_h2s if h.get("id") != "application-method"]
    if job_h2s:
        company_blurb = extract_company_blurb(printable, job_h2s[0])

    jobs = []

    if keyinfo_blocks:
        for ul in keyinfo_blocks:
            h2 = ul.find_previous_sibling("h2")
            numeric_id = ""
            if h2 is not None and h2.get("id", "").startswith("job") and h2.get("id") != "application-method":
                a = h2.select_one("a.subjob-title") or h2.select_one("a")
                title = clean_text(a) if a else clean_text(h2)
                job_url = absolute_url(a["href"]) if a and a.get("href") else url
                numeric_id = re.sub(r"\D", "", h2.get("id", ""))
            else:
                # Fallback: no per-job <h2>, treat the whole page as one job
                title = strip_company_suffix(page_title_raw)
                job_url = url
                if is_generic_title(page_title_raw):
                    log(f"  Skipping generic page title with no per-job blocks: {page_title_raw}")
                    continue

            if not title:
                continue

            info = {}
            for li in ul.select("li"):
                key = clean_text(li.select_one(".jkey-title"))
                val = clean_text(li.select_one(".jkey-info"))
                if key:
                    info[key] = val

            details_div = ul.find_next_sibling("div", class_="job-details")
            description = clean_description(clean_text(details_div))

            raw_apply_url = apply_links.get(numeric_id, "") if numeric_id else single_apply_fallback

            log(f"    Resolving apply link for '{title}'...")
            application = resolve_application_contact(raw_apply_url, description)

            jobs.append({
                "title": title,
                "job_url": job_url,
                "job_type": info.get("Job Type", ""),
                "qualification": info.get("Qualification", ""),
                "experience": info.get("Experience", ""),
                "location": info.get("Location", ""),
                "city": info.get("City", ""),
                "field": info.get("Job Field", ""),
                "posted_date": posted_date_raw,
                "deadline": deadline,
                "description": description,
                "apply_url": application["apply_url"],
                "apply_email": application["apply_email"],
                "apply_raw": application["apply_raw"],
                "company_name": company_name,
                "company_url": company_url,
                "company_blurb": company_blurb,
                "source_page": url,
            })
    else:
        log(f"  No ul.job-key-info blocks found at all on: {url} (template not recognized, skipping)")

    return jobs


# -----------------------------------------------------------------------------
# STEP 3 — VERBOSE PRINTING
# -----------------------------------------------------------------------------

def print_job_verbose(index, job):
    print("\n" + "-" * 80)
    print(f"JOB #{index}")
    print("-" * 80)
    print(f"Title           : {job['title']}")
    print(f"Company         : {job['company_name']}")
    print(f"Company URL     : {job['company_url']}")
    print(f"Job URL         : {job['job_url']}")
    print(f"Job Type        : {job['job_type']}")
    print(f"Qualification   : {job['qualification']}")
    print(f"Experience      : {job['experience']}")
    print(f"Location        : {job['location']}")
    print(f"City            : {job['city']}")
    print(f"Field           : {job['field']}")
    print(f"Posted Date     : {job['posted_date']}")
    print(f"Deadline        : {job['deadline']}")

    if job["apply_url"]:
        print(f"Apply URL       : {job['apply_url']}")
    elif job["apply_email"]:
        print(f"Apply Email     : {job['apply_email']}")
    else:
        print("Apply URL       : (not found - check source page)")
    if job["apply_raw"]:
        print(f"  (MyJobMag tracking link was: {job['apply_raw']})")

    print(f"Source Page     : {job['source_page']}")
    if job["company_blurb"]:
        print(f"Company Blurb   : {job['company_blurb']}")
    print("Description     :")
    print(job["description"] if job["description"] else "(none extracted)")


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

def main():
    start_time = datetime.now()
    log(f"SCRAPE RUN STARTED: {start_time.isoformat()}")
    log(f"SCRAPE_PAGES={SCRAPE_PAGES}  REQUEST_DELAY={REQUEST_DELAY}  MAX_JOBS={MAX_JOBS or 'unlimited'}")
    log(f"RESOLVE_APPLY_URLS={RESOLVE_APPLY_URLS}  RESOLVE_DELAY={RESOLVE_DELAY}")

    page_urls = collect_company_page_urls(SCRAPE_PAGES)

    total_jobs = 0
    errors = 0

    for i, page_url in enumerate(page_urls, start=1):
        log(f"\nScraping page {i}/{len(page_urls)}: {page_url}")
        try:
            jobs = parse_job_page(page_url)
        except Exception as e:
            errors += 1
            log(f"  ERROR scraping {page_url}: {e}")
            continue

        log(f"  -> {len(jobs)} job(s) extracted from this page")

        for job in jobs:
            total_jobs += 1
            print_job_verbose(total_jobs, job)
            if MAX_JOBS and total_jobs >= MAX_JOBS:
                log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
                break

        if MAX_JOBS and total_jobs >= MAX_JOBS:
            break

        time.sleep(REQUEST_DELAY)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    log(f"\n{'=' * 80}")
    log("SCRAPE COMPLETE")
    log(f"  Company/job pages visited : {len(page_urls)}")
    log(f"  Total individual jobs     : {total_jobs}")
    log(f"  Errors                    : {errors}")
    log(f"  Duration                  : ~{duration:.1f} min")
    log("=" * 80)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
