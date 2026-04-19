# -*- coding: utf-8 -*-
# Combined SEO and Keyword Audit Script

import argparse
import concurrent.futures as futures
import re
import sys
import time
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlparse, urldefrag
from collections import deque
import pandas as pd
from bs4 import BeautifulSoup

# --- Selenium and Requests for Web Crawling ---
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


INPUT_XLSX = r"C:\path\to\your\input_file.xlsx"
INPUT_COL = "organization_website_url"
OUTPUT_XLSX = r"C:\path\to\your\output_file.xlsx"

MAX_WORKERS = 5

RESIDENTIAL_KEYWORDS = ["Residential", "Res A/C", "HAC Residential", "Residential AC"]
COMMERCIAL_KEYWORDS = ["Commercial", "Comm A/C", "Commercial AC", "Commercial HVAC"]
INDUSTRIAL_KEYWORDS = ["Industrial", "Ind A/C", "Industrial AC", "Industrial HVAC"]
 
REQUEST_TIMEOUT = (10, 20)
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0 Safari/537.36"
MAX_PAGES = 20  
SLEEP_BETWEEN_REQUESTS = 1



PATTERNS = {
    "Google Tag Manager": [
        re.compile(r"gooogletagmanager\.com/gtm\.js\?id=GTM-[A-Z0-9]+", re.I),
        re.compile(r"googletagmanager\.com/gtm\.js", re.I),
        re.compile(r"GTM-[A-Z0-9]+", re.I),
        re.compile(r"<noscript>.*googletagmanager\.com/ns\.html\?id=GTM-", re.I | re.S),
    ],
    "Google Analytics 4": [
        re.compile(r"gtag\(\s*'config'\s*,\s*'G-(?=[A-Z0-9]+)", re.I),
        re.compile(r"G-[A-Z0-9]{6,}", re.I),
        re.compile(r"googletagmanager\.com/gtag/js\?id=G-", re.I),
    ],
    "Google Analytics (UA)": [
        re.compile(r"UA-\d{4,}-\d+", re.I),
        re.compile(r"analytics\.js", re.I),
    ],
    "Google Ads (AW)": [
        re.compile(r"AW-\d{6,}", re.I),
        re.compile(r"googletagmanager\.com/gtag/js\?id=AW-", re.I),
    ],
    "Meta Pixel": [
        re.compile(r"connect\.facebook\.net/.*/fbevents\.js", re.I),
        re.compile(r"fbq\('init'", re.I),
        re.compile(r"fbq\('track'", re.I),
    ],
    "TikTok Pixel": [
        re.compile(r"analytics\.tiktok\.com/i18n/pixel/events\.js", re.I),
        re.compile(r"ttq\s*=\s*window\.ttq", re.I),
        re.compile(r"ttq\.track\(", re.I),
    ],
    "LinkedIn Insight": [
        re.compile(r"snap\.licdn\.com/li\.lms-analytics/insight\.min\.js", re.I),
        re.compile(r"linkedin_partner_id", re.I),
    ],
    "Hotjar": [
        re.compile(r"static\.hotjar\.com/c/hotjar-", re.I),
        re.compile(r"hotjar\.com", re.I),
        re.compile(r"hj\('trigger'", re.I),
    ],
    "Microsoft Clarity": [
        re.compile(r"clarity/ms\.js", re.I),
        re.compile(r"clarity\(", re.I),
    ],
    "HubSpot": [
        re.compile(r"js\.hs-analytics\.net", re.I),
        re.compile(r"hs-scripts\.com", re.I),
    ],
    "Segment": [
        re.compile(r"cdn\.segment\.com/analytics", re.I),
        re.compile(r"analytics\.load\(", re.I),
    ],
}


@dataclass
class AuditRow:
    url: str
    final_url: str
    http_status: Optional[int]
    meta_title: str
    meta_description: str
    title_length: int
    description_length: int
    total_images: int
    images_missing_alt: int
    alt_tag_coverage: str
    internal_links_count: int
    trackers_found: str
    meta_quality: str
    image_seo_status: str
    overall_seo_status: str
    suggestion: str
    residential_found: str
    residential_source: str
    commercial_found: str
    commercial_source: str
    industrial_found: str
    industrial_source: str
    found_keywords_list: str
    address: str
    notes: str


def normalize_url(u: str) -> Optional[str]:
    if not isinstance(u, str) or not u.strip():
        return None
    u = u.strip()
    if not re.match(r"^https?://", u, re.I):
        u = "http://" + u
    return u


def normalize_link_for_crawl(base_url, link):
    link, _ = urldefrag(urljoin(base_url, link))
    return link.rstrip("/")


def fetch_and_parse(url: str) -> Tuple[Optional[requests.Response], Optional[str]]:
    try:
        headers = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
        resp = requests.get(
            url,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            verify=True,
        )
        resp.raise_for_status()
        return resp, None
    except requests.exceptions.SSLError:
        try:
            resp = requests.get(
                url,
                headers={"User-Agent": USER_AGENT},
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
                verify=False,
            )
            resp.raise_for_status()
            return resp, "SSL issue (proceeded without verification)"
        except Exception as e:
            return None, str(e)
    except Exception as e:
        return None, str(e)


def extract_meta(html: str) -> Tuple[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    description = ""
    m = soup.find("meta", attrs={"name": "description"})
    if m and m.get("content"):
        description = m["content"].strip()
    return title, description


def analyze_images(soup: BeautifulSoup) -> Tuple[int, int, str]:
    images = soup.find_all("img")
    total_images = len(images)
    if total_images == 0:
        return 0, 0, "No images found"
    missing_alt = sum(1 for img in images if not img.get("alt", "").strip())
    coverage_pct = ((total_images - missing_alt) / total_images) * 100
    if coverage_pct == 100:
        status = "Excellent (100%)"
    elif coverage_pct >= 80:
        status = f"Good ({coverage_pct:.0f}%)"
    else:
        status = f"Poor ({coverage_pct:.0f}%)"
    return total_images, missing_alt, status


def count_internal_links(soup: BeautifulSoup, base_url: str) -> int:
    parsed_base = urlparse(base_url)
    base_domain = parsed_base.netloc.lower()
    links = soup.find_all("a", href=True)
    internal_count = 0
    for link in links:
        href = link.get("href", "").strip()
        if not href:
            continue
        absolute_url = urljoin(base_url, href)
        parsed_url = urlparse(absolute_url)
        if parsed_url.netloc.lower() == base_domain:
            internal_count += 1
    return internal_count


def detect_trackers(html: str) -> List[str]:
    found = []
    for label, regex_list in PATTERNS.items():
        for rgx in regex_list:
            if rgx.search(html):
                found.append(label)
                break
    seen = set()
    return [x for x in found if not (x in seen or seen.add(x))]


def score_meta_quality(title: str, desc: str) -> str:
    tl, dl = len(title or ""), len(desc or "")
    title_ok = 30 <= tl <= 65
    desc_ok = 70 <= dl <= 170
    if not title and not desc:
        return "Poor (missing title & description)"
    if title_ok and desc_ok:
        return "Good"
    return "Weak (length issues or missing)"


def score_image_seo(total_images: int, missing_alt: int) -> str:
    if total_images == 0:
        return "No images"
    coverage_pct = ((total_images - missing_alt) / total_images) * 100
    if coverage_pct == 100:
        return "Excellent"
    elif coverage_pct >= 80:
        return "Good"
    elif coverage_pct >= 50:
        return "Needs improvement"
    else:
        return "Poor"


def calculate_overall_seo_status(meta_quality: str, image_seo: str, internal_links: int) -> str:
    meta_score = 3 if "Good" in meta_quality else 1
    image_score = 3 if image_seo == "Excellent" else 1 if image_seo == "Good" else 0
    if image_seo == "No images": image_score = 2
    links_score = 3 if internal_links >= 20 else 2 if internal_links >= 10 else 1
    total_score = meta_score + image_score + links_score
    if total_score >= 7: return "Excellent"
    elif total_score >= 5: return "Good"
    elif total_score >= 3: return "Fair"
    else: return "Poor"


def suggest_service(meta_quality: str, image_seo: str, overall_status: str, trackers: List[str]) -> str:
    has_trackers = len(trackers) > 0
    if overall_status == "Poor":
        return "Comprehensive SEO Overhaul + Digital Marketing"
    elif overall_status == "Fair":
        return "SEO Optimization Package" if not has_trackers else "SEO + Analytics Review"
    elif overall_status == "Good":
        return "Digital Marketing Setup" if not has_trackers else "Advanced Growth Strategy"
    elif overall_status == "Excellent":
        return "Digital Marketing & Growth" if not has_trackers else "Premium Growth & Analytics"
    return "SEO + Digital Marketing Consultation"


def create_driver():
    options = Options()
    options.headless = True
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)
    return driver


def crawl_for_keywords_and_address(url: str) -> Dict:
    result = {
        "residential_found": "No",
        "residential_source": "",
        "commercial_found": "No",
        "commercial_source": "",
        "industrial_found": "No",
        "industrial_source": "",
        "found_keywords_list": "",
        "address": ""
    }
    visited = set()
    queue = deque([url])
    found_keywords = set()
    addresses = set()

    patterns = {
        "Residential": re.compile(r"\b(" + "|".join(map(re.escape, RESIDENTIAL_KEYWORDS)) + r")\b", re.IGNORECASE),
        "Commercial": re.compile(r"\b(" + "|".join(map(re.escape, COMMERCIAL_KEYWORDS)) + r")\b", re.IGNORECASE),
        "Industrial": re.compile(r"\b(" + "|".join(map(re.escape, INDUSTRIAL_KEYWORDS)) + r")\b", re.IGNORECASE)
    }

    driver = None
    try:
        driver = create_driver()
        while queue and len(visited) < MAX_PAGES:
            current_url = queue.popleft()
            if current_url in visited:
                continue
            visited.add(current_url)

            try:
                driver.get(current_url)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(SLEEP_BETWEEN_REQUESTS)
                html = driver.page_source
                soup = BeautifulSoup(html, "lxml")
                full_text = soup.get_text(" ", strip=True)

                for key, pattern in patterns.items():
                    if pattern.search(full_text):
                        found_keywords.add(key)
                        result[f"{key.lower()}_found"] = "Yes"
                        result[f"{key.lower()}_source"] = current_url

                for addr_tag in soup.find_all("address"):
                    text = addr_tag.get_text(" ", strip=True)
                    if text: addresses.add(text)

                address_keywords = ["Street", "St.", "Road", "Rd.", "Ave", "Avenue", "Suite", "City", "State"]
                for tag in soup.find_all(["p", "div", "span"]):
                    text = tag.get_text(" ", strip=True)
                    if any(k in text for k in address_keywords) and len(text) < 200:
                        addresses.add(text)

                for a in soup.find_all("a", href=True):
                    link = normalize_link_for_crawl(url, a["href"])
                    if link.startswith(url) and link not in visited and len(visited) + len(queue) < MAX_PAGES:
                        queue.append(link)

            except (WebDriverException, TimeoutException) as e:
                print(f"Error crawling {current_url}: {e}")
                continue
    except Exception as e:
        print(f"Failed to initialize driver or unexpected error: {e}")
    finally:
        if driver:
            driver.quit()

    result["found_keywords_list"] = ", ".join(found_keywords)
    result["address"] = " ; ".join(addresses)
    return result


def audit_one_combined(raw_url: str) -> AuditRow:
    url = normalize_url(raw_url)
    if not url:
        return AuditRow(
            url=str(raw_url), final_url="", http_status=None, meta_title="", meta_description="", title_length=0,
            description_length=0, total_images=0, images_missing_alt=0, alt_tag_coverage="Invalid URL",
            internal_links_count=0, trackers_found="", meta_quality="Invalid URL", image_seo_status="N/A",
            overall_seo_status="Invalid", suggestion="Review Manually",
            residential_found="No", residential_source="", commercial_found="No", commercial_source="",
            industrial_found="No", industrial_source="", found_keywords_list="", address="", notes="Empty or invalid URL"
        )

    # Part 1: Perform SEO audit using requests
    resp, note = fetch_and_parse(url)
    if not resp:
        return AuditRow(
            url=url, final_url="", http_status=None, meta_title="", meta_description="", title_length=0,
            description_length=0, total_images=0, images_missing_alt=0, alt_tag_coverage="Unreachable",
            internal_links_count=0, trackers_found="", meta_quality="Unreachable", image_seo_status="N/A",
            overall_seo_status="Unreachable", suggestion="Review Manually",
            residential_found="No", residential_source="", commercial_found="No", commercial_source="",
            industrial_found="No", industrial_source="", found_keywords_list="", address="", notes=note or "Request failed"
        )

    html_content = resp.text
    soup = BeautifulSoup(html_content, "html.parser")
    title, description = extract_meta(html_content)
    trackers = detect_trackers(html_content)
    meta_quality = score_meta_quality(title, description)
    total_images, missing_alt, alt_coverage = analyze_images(soup)
    image_seo = score_image_seo(total_images, missing_alt)
    internal_links = count_internal_links(soup, resp.url)
    overall_status = calculate_overall_seo_status(meta_quality, image_seo, internal_links)
    suggestion = suggest_service(meta_quality, image_seo, overall_status, trackers)

    # Part 2: Perform keyword crawl using Selenium
    keyword_results = crawl_for_keywords_and_address(url)

    return AuditRow(
        url=raw_url,
        final_url=resp.url,
        http_status=resp.status_code,
        meta_title=title,
        meta_description=description,
        title_length=len(title or ""),
        description_length=len(description or ""),
        total_images=total_images,
        images_missing_alt=missing_alt,
        alt_tag_coverage=alt_coverage,
        internal_links_count=internal_links,
        trackers_found=", ".join(trackers) if trackers else "",
        meta_quality=meta_quality,
        image_seo_status=image_seo,
        overall_seo_status=overall_status,
        suggestion=suggestion,
        residential_found=keyword_results["residential_found"],
        residential_source=keyword_results["residential_source"],
        commercial_found=keyword_results["commercial_found"],
        commercial_source=keyword_results["commercial_source"],
        industrial_found=keyword_results["industrial_found"],
        industrial_source=keyword_results["industrial_source"],
        found_keywords_list=keyword_results["found_keywords_list"],
        address=keyword_results["address"],
        notes=note or ""
    )


def read_input_excel(path: str, url_col: str) -> List[str]:
    df = pd.read_excel(path, engine="openpyxl")
    if url_col not in df.columns:
        raise ValueError(f"Column '{url_col}' not found in {path}. Available columns: {list(df.columns)}")
    urls = df[url_col].dropna().astype(str).tolist()
    return urls


def write_output_excel(rows: List[AuditRow], output_path: str) -> None:
    df = pd.DataFrame([asdict(r) for r in rows])
    cols = [
        "url", "final_url", "http_status", "meta_title", "title_length", "meta_description",
        "description_length", "total_images", "images_missing_alt", "alt_tag_coverage",
        "internal_links_count", "trackers_found", "meta_quality", "image_seo_status",
        "overall_seo_status", "suggestion", "residential_found", "residential_source",
        "commercial_found", "commercial_source", "industrial_found", "industrial_source",
        "found_keywords_list", "address", "notes"
    ]
    df = df[cols]
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SEO_Audit")
        summary_data = {
            "Metric": [
                "Total Websites Audited", "Successfully Crawled", "Failed to Crawl",
                "Excellent SEO Status", "Good SEO Status", "Fair SEO Status",
                "Poor SEO Status", "Sites with Perfect Alt Tags", "Sites Missing Alt Tags",
                "Average Internal Links"
            ],
            "Count": [
                len(df), len(df[df['http_status'].notna()]), len(df[df['http_status'].isna()]),
                len(df[df['overall_seo_status'] == 'Excellent']), len(df[df['overall_seo_status'] == 'Good']),
                len(df[df['overall_seo_status'] == 'Fair']), len(df[df['overall_seo_status'] == 'Poor']),
                len(df[df['images_missing_alt'] == 0]), len(df[df['images_missing_alt'] > 0]),
                round(df['internal_links_count'].mean(), 1) if len(df) > 0 else 0
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
    print(f"✅ Wrote report: {output_path} ({len(df)} rows)")


def main():
    try:
        urls = read_input_excel(INPUT_XLSX, INPUT_COL)
    except FileNotFoundError:
        print(f"❌ Error: The file '{INPUT_XLSX}' was not found.")
        sys.exit(1)
    except ValueError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)

    print(f"🔍 Auditing {len(urls)} websites for comprehensive SEO and keywords...")
    print("📊 Checking: Meta tags, Alt tags, Internal links, Tracking tools, and keyword presence")
    t0 = time.time()

    rows: List[AuditRow] = []
    with futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        future_to_url = {ex.submit(audit_one_combined, url): url for url in urls}
        for future in futures.as_completed(future_to_url):
            url = future_to_url[future]
            try:
                row = future.result()
                rows.append(row)
            except Exception as e:
                print(f"Error processing URL {url}: {e}")
                rows.append(AuditRow(
                    url=url, final_url="", http_status=None, meta_title="", meta_description="",
                    title_length=0, description_length=0, total_images=0, images_missing_alt=0,
                    alt_tag_coverage="Unreachable", internal_links_count=0, trackers_found="", meta_quality="Unreachable",
                    image_seo_status="N/A", overall_seo_status="Unreachable", suggestion="Review Manually",
                    residential_found="No", residential_source="", commercial_found="No", commercial_source="",
                    industrial_found="No", industrial_source="", found_keywords_list="", address="",
                    notes=f"Error: {e}"
                ))

    write_output_excel(rows, OUTPUT_XLSX)
    dt = time.time() - t0
    print(f"⏱️ Done in {dt:.1f}s")


if __name__ == "__main__":
    main()