from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import TYPE_CHECKING
from urllib.parse import parse_qs, quote_plus, unquote, urlparse
from zipfile import ZIP_DEFLATED, ZipFile

if TYPE_CHECKING:
    from scrapling.fetchers import FetcherSession, StealthySession

KEYWORDS_FILE = Path("keywords.txt")
ALL_OUTPUT_FILE = Path("all_results.xlsx")

HEADLESS = False  # اجعلها True إذا أردت التشغيل بدون إظهار المتصفح
REAL_CHROME = True  # اجعلها True إذا أردت استخدام Chrome المثبت على جهازك
MAX_RESULTS_PER_KEYWORD = 700
SCROLL_PAUSE_MS = 2000
NO_NEW_RESULTS_LIMIT = 8
REQUEST_TIMEOUT = 60000

EMAIL_REGEX = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
PHONE_REGEX = re.compile(r"(?:\+?\d[\d\s().-]{7,}\d)")
EMAIL_GUESS_PREFIXES = ("Info", "Sales")

CONSENT_TEXTS = [
    "Accept all",
    "I agree",
    "Alle akzeptieren",
    "Tout accepter",
    "Aceptar todo",
    "Accetta tutto",
    "قبول الكل",
]

CONTACT_HINTS = (
    "contact",
    "contact-us",
    "about",
    "about-us",
    "support",
    "help",
    "tawasol",
    "اتصل",
    "تواصل",
    "اتصل-بنا",
)


@dataclass
class BusinessRecord:
    keyword: str
    name: str = "N/A"
    email: str = "N/A"
    phone: str = "N/A"
    website: str = "N/A"
    maps_url: str = "N/A"
    registered_email_found: str = "No"
    email_status: str = "No public email found"


def load_scrapling_fetchers():
    try:
        from scrapling.fetchers import FetcherSession, StealthySession
        return FetcherSession, StealthySession
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "مكتبة Scrapling fetchers غير مكتملة على جهازك.\n"
            "ثبّت المتطلبات بهذا الأمر أولًا:\n"
            "python -m pip install \"scrapling[fetchers]\"\n"
            "ثم نفّذ:\n"
            "scrapling install\n\n"
            f"السبب الحالي: {exc}"
        ) from exc


def read_keywords(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"الملف غير موجود: {path}")

    keywords = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not keywords:
        raise ValueError("ملف keywords.txt فارغ، أضف كلمة مفتاحية في كل سطر.")
    return keywords


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    return " ".join(str(value).replace("\u200b", " ").split()).strip()


def is_google_domain(url: str) -> bool:
    host = (urlparse(url).netloc or "").lower()
    return "google." in host or "gstatic.com" in host


def normalize_maps_url(url: str | None) -> str:
    url = clean_text(url)
    if not url:
        return ""

    if url.startswith("/"):
        url = f"https://www.google.com{url}"

    if "google.com/url?" in url:
        parsed = urlparse(url)
        query = parse_qs(parsed.query)
        target = query.get("q") or query.get("url")
        if target:
            url = unquote(target[0])

    return url


def normalize_website_url(url: str | None) -> str:
    url = normalize_maps_url(url)
    if not url:
        return ""

    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    for key in ("q", "url", "continue"):
        if key in query:
            candidate = unquote(query[key][0])
            if candidate.startswith("http"):
                url = candidate
                parsed = urlparse(url)
                break

    if not parsed.scheme.startswith("http"):
        return ""

    return url


def sanitize_filename(name: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*]+', '_', clean_text(name))
    safe = safe.replace(' ', '_')
    safe = re.sub(r'_+', '_', safe).strip('._ ')
    return safe or 'keyword_results'


def safe_sheet_name(name: str) -> str:
    safe = re.sub(r'[\[\]:*?/\\]+', '_', clean_text(name))
    return (safe or 'Results')[:31]


def get_domain_from_website(website_url: str) -> str:
    if not website_url or website_url == 'N/A':
        return ''

    host = urlparse(website_url).netloc.lower().strip()
    if host.startswith('www.'):
        host = host[4:]
    return host


def generate_email_guesses(website_url: str) -> str:
    domain = get_domain_from_website(website_url)
    if not domain:
        return 'N/A'

    guesses = [f'{prefix}@{domain}' for prefix in EMAIL_GUESS_PREFIXES]
    return ' | '.join(dict.fromkeys(guesses))


def record_unique_key(record: BusinessRecord) -> tuple[str, str, str, str]:
    return (
        clean_text(record.name).lower(),
        clean_text(record.phone).lower(),
        clean_text(record.website).lower(),
        clean_text(record.maps_url).lower(),
    )


def dedupe_records(records: list[BusinessRecord]) -> list[BusinessRecord]:
    unique_records: list[BusinessRecord] = []
    seen: set[tuple[str, str, str, str]] = set()

    for record in records:
        key = record_unique_key(record)
        if key in seen:
            continue
        seen.add(key)
        unique_records.append(record)

    return unique_records


def excel_column_name(index: int) -> str:
    result = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def write_xlsx(rows: list[list[str]], output_path: Path, sheet_name: str = 'Results') -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not rows:
        rows = [['No results found']]

    max_cols = max(len(row) for row in rows)
    last_col = excel_column_name(max_cols)
    last_row = len(rows)
    dimension = f'A1:{last_col}{last_row}'

    row_xml_parts: list[str] = []
    for row_idx, row in enumerate(rows, start=1):
        cells: list[str] = []
        style_index = 1 if row_idx == 1 else 0
        padded_row = row + [''] * (max_cols - len(row))

        for col_idx, value in enumerate(padded_row, start=1):
            cell_ref = f'{excel_column_name(col_idx)}{row_idx}'
            cell_value = escape('' if value is None else str(value))
            cells.append(
                f'<c r="{cell_ref}" t="inlineStr" s="{style_index}">' \
                f'<is><t xml:space="preserve">{cell_value}</t></is></c>'
            )

        row_xml_parts.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<sheetData>{"".join(row_xml_parts)}</sheetData>'
        '</worksheet>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{escape(safe_sheet_name(sheet_name))}" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="2">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )

    with ZipFile(output_path, 'w', ZIP_DEFLATED) as workbook:
        workbook.writestr('[Content_Types].xml', content_types_xml)
        workbook.writestr('_rels/.rels', root_rels_xml)
        workbook.writestr('xl/workbook.xml', workbook_xml)
        workbook.writestr('xl/_rels/workbook.xml.rels', workbook_rels_xml)
        workbook.writestr('xl/worksheets/sheet1.xml', worksheet_xml)
        workbook.writestr('xl/styles.xml', styles_xml)


def records_to_rows(records: list[BusinessRecord]) -> list[list[str]]:
    rows: list[list[str]] = [[
        'Keyword',
        'Company Name',
        'Email',
        'Phone',
        'Website',
        'Google Maps URL',
        'Registered Email Found',
        'Email Status',
    ]]

    for record in records:
        rows.append([
            record.keyword,
            record.name,
            record.email,
            record.phone,
            record.website,
            record.maps_url,
            record.registered_email_found,
            record.email_status,
        ])

    return rows


def save_keyword_excel_files(results_by_keyword: dict[str, list[BusinessRecord]]) -> None:
    for keyword, records in results_by_keyword.items():
        file_path = Path(f'{sanitize_filename(keyword)}.xlsx')
        write_xlsx(records_to_rows(records), file_path, sheet_name=keyword)
        print(f'✅ تم حفظ نتائج الكلمة المفتاحية في: {file_path.resolve()}')


def first_value(page, selectors: list[str]) -> str:
    for selector in selectors:
        try:
            value = page.css(selector).get()
        except Exception:
            value = None
        cleaned = clean_text(value)
        if cleaned:
            return cleaned
    return ""


def all_values(page, selectors: list[str]) -> list[str]:
    values: list[str] = []
    for selector in selectors:
        try:
            values.extend(page.css(selector).getall())
        except Exception:
            continue
    return [clean_text(value) for value in values if clean_text(value)]


def extract_dumped_place_links(page) -> list[str]:
    dumped_text = first_value(page, ['#__scrapling_place_links::text'])
    if not dumped_text:
        return []
    return [clean_text(line) for line in dumped_text.splitlines() if clean_text(line)]


def try_accept_google_consent(playwright_page, target_url: str) -> None:
    if "consent.google.com" not in (playwright_page.url or ""):
        return

    for text in CONSENT_TEXTS:
        selectors = [
            f'button:has-text("{text}")',
            f'[role="button"]:has-text("{text}")',
        ]
        for selector in selectors:
            try:
                locator = playwright_page.locator(selector)
                if locator.count():
                    locator.first.click(timeout=3000, force=True, no_wait_after=True)
                    playwright_page.wait_for_timeout(2500)
                    if "consent.google.com" not in (playwright_page.url or ""):
                        return
            except Exception:
                continue

    if "consent.google.com" in (playwright_page.url or ""):
        try:
            playwright_page.goto(target_url, wait_until="domcontentloaded")
            playwright_page.wait_for_timeout(3000)
        except Exception:
            pass


def build_search_action(target_url: str, max_results: int):
    def action(playwright_page):
        try_accept_google_consent(playwright_page, target_url)
        playwright_page.wait_for_timeout(2000)

        last_count = 0
        no_new_results_rounds = 0

        while True:
            try:
                state = playwright_page.evaluate(
                    """
                    () => {
                        const selectors = ['a.hfpxzc', 'a[href*="/maps/place/"]', 'a[href*="/place/"]'];
                        const holderId = '__scrapling_place_links';
                        const existing = Array.isArray(window.__scrapling_place_links)
                            ? window.__scrapling_place_links
                            : [];
                        const knownLinks = new Set(existing);

                        for (const selector of selectors) {
                            for (const node of document.querySelectorAll(selector)) {
                                const href = node.href || node.getAttribute('href') || '';
                                if (href) {
                                    knownLinks.add(href);
                                }
                            }
                        }

                        window.__scrapling_place_links = Array.from(knownLinks);

                        let holder = document.getElementById(holderId);
                        if (!holder) {
                            holder = document.createElement('div');
                            holder.id = holderId;
                            holder.style.display = 'none';
                            document.body.appendChild(holder);
                        }
                        holder.textContent = window.__scrapling_place_links.join('\\n');

                        const pageText = document.body ? document.body.innerText : '';
                        const endReached = /You've reached the end of the list|You\\'ve reached the end of the list|end of the list|نهاية القائمة|لقد وصلت إلى نهاية القائمة/i.test(pageText);

                        return {
                            count: window.__scrapling_place_links.length,
                            endReached,
                        };
                    }
                    """
                )
            except Exception:
                state = {"count": last_count, "endReached": False}

            current_count = int(state.get("count", 0))

            if current_count >= max_results:
                break

            if current_count > last_count:
                last_count = current_count
                no_new_results_rounds = 0
            else:
                no_new_results_rounds += 1

            if state.get("endReached") or no_new_results_rounds >= NO_NEW_RESULTS_LIMIT:
                break

            try:
                feed = playwright_page.locator('div[role="feed"]')
                if feed.count():
                    feed.first.hover(timeout=2000)
                    feed.first.evaluate(
                        "(el) => { el.scrollTop = el.scrollHeight; el.dispatchEvent(new Event('scroll')); }"
                    )
                else:
                    playwright_page.mouse.wheel(0, 4000)
            except Exception:
                try:
                    playwright_page.mouse.wheel(0, 4000)
                except Exception:
                    pass

            playwright_page.wait_for_timeout(SCROLL_PAUSE_MS)

    return action


def collect_place_links(maps_session: StealthySession, keyword: str) -> list[str]:
    search_url = f"https://www.google.com/maps/search/{quote_plus(keyword)}"
    print(f"\n[+] البحث عن: {keyword}")

    try:
        page = maps_session.fetch(
            search_url,
            timeout=REQUEST_TIMEOUT,
            wait=1500,
            network_idle=True,
            wait_selector='a.hfpxzc, a[href*="/place/"]',
            page_action=build_search_action(search_url, MAX_RESULTS_PER_KEYWORD),
        )
    except Exception as exc:
        print(f"[!] فشل فتح صفحة البحث: {exc}")
        return []

    raw_links: list[str] = []
    raw_links.extend(extract_dumped_place_links(page))
    raw_links.extend(all_values(page, ["a.hfpxzc::attr(href)"]))
    raw_links.extend(all_values(page, ['a[href*="/maps/place/"]::attr(href)']))
    raw_links.extend(all_values(page, ['a[href*="/place/"]::attr(href)']))

    seen: set[str] = set()
    cleaned_links: list[str] = []

    for link in raw_links:
        normalized = normalize_maps_url(link)
        if not normalized:
            continue
        if "/place/" not in normalized and "/maps/place/" not in normalized:
            continue
        if normalized in seen:
            continue
        seen.add(normalized)
        cleaned_links.append(normalized)

        if len(cleaned_links) >= MAX_RESULTS_PER_KEYWORD:
            break

    print(f"[+] تم تجميع {len(cleaned_links)} رابط مكان من أصل المطلوب {MAX_RESULTS_PER_KEYWORD}")
    return cleaned_links


def extract_phone(page) -> str:
    candidates = all_values(
        page,
        [
            'a[href^="tel:"]::attr(href)',
            'button[data-item-id*="phone"]::attr(aria-label)',
            'button[aria-label*="+"]::attr(aria-label)',
            'a[aria-label*="+"]::attr(aria-label)',
        ],
    )

    for candidate in candidates:
        cleaned = candidate.replace("tel:", "")
        cleaned = re.sub(r"^(phone|call|telephone|telefon|الهاتف|رقم الهاتف)\s*:?\s*", "", cleaned, flags=re.I)
        match = PHONE_REGEX.search(cleaned)
        if match:
            return clean_text(match.group(0))

    return "N/A"


def extract_website(page) -> str:
    candidates = all_values(
        page,
        [
            'a[data-item-id="authority"]::attr(href)',
            'a[aria-label*="Website"]::attr(href)',
            'a[aria-label*="site web"]::attr(href)',
            'a[aria-label*="الموقع"]::attr(href)',
            'a[href^="http"]::attr(href)',
        ],
    )

    for candidate in candidates:
        url = normalize_website_url(candidate)
        if not url:
            continue
        if is_google_domain(url):
            continue
        return url

    return "N/A"


def extract_name(page) -> str:
    name = first_value(
        page,
        [
            "h1.DUwDvf::text",
            "h1::text",
            'meta[property="og:title"]::attr(content)',
        ],
    )
    return name or "N/A"


def extract_email_from_response(page) -> str:
    mailto_links = all_values(page, ['a[href^="mailto:"]::attr(href)'])
    for mail in mailto_links:
        cleaned = mail.replace("mailto:", "").split("?")[0].strip().lower()
        if cleaned:
            return cleaned

    text_chunks = all_values(page, ["body *::text"])
    if text_chunks:
        joined_text = " ".join(text_chunks)
        matches = EMAIL_REGEX.findall(joined_text)
        for email in matches:
            email = email.strip(" <>.,;:'\"").lower()
            if email:
                return email

    return "N/A"


def find_contact_page_urls(page, base_url: str) -> list[str]:
    parsed_base = urlparse(base_url)
    base_root = f"{parsed_base.scheme}://{parsed_base.netloc}"

    found: list[str] = []
    for href in all_values(page, ["a::attr(href)"]):
        lowered = href.lower()
        if not any(hint in lowered for hint in CONTACT_HINTS):
            continue

        if href.startswith("/"):
            href = f"{base_root}{href}"
        elif not href.startswith("http"):
            continue

        if href not in found:
            found.append(href)

    return found[:3]


def fetch_email_from_website(http_session: FetcherSession, website_url: str) -> str:
    if not website_url or website_url == "N/A":
        return "N/A"

    queue = [website_url]
    visited: set[str] = set()

    while queue:
        current_url = queue.pop(0)
        if current_url in visited:
            continue
        visited.add(current_url)

        try:
            page = http_session.get(current_url, stealthy_headers=True)
        except Exception:
            continue

        email = extract_email_from_response(page)
        if email != "N/A":
            return email

        if len(visited) == 1:
            queue.extend(find_contact_page_urls(page, current_url))

    return "N/A"


def scrape_business(maps_session: StealthySession, http_session: FetcherSession, keyword: str, place_url: str) -> BusinessRecord | None:
    try:
        page = maps_session.fetch(
            place_url,
            timeout=REQUEST_TIMEOUT,
            wait=1000,
            network_idle=True,
        )
    except Exception as exc:
        print(f"[!] تعذر فتح صفحة النشاط: {exc}")
        return None

    name = extract_name(page)
    phone = extract_phone(page)
    website = extract_website(page)
    email = fetch_email_from_website(http_session, website)
    registered_email_found = "Yes"
    email_status = "Public email found on website"

    if email == "N/A" and website != "N/A":
        email = generate_email_guesses(website)
        registered_email_found = "No"
        email_status = "Generated from domain because no public email was found"
    elif email == "N/A":
        registered_email_found = "No"
        email_status = "No email and no website"

    return BusinessRecord(
        keyword=keyword,
        name=name,
        email=email,
        phone=phone,
        website=website,
        maps_url=place_url,
        registered_email_found=registered_email_found,
        email_status=email_status,
    )


def main() -> None:
    FetcherSession, StealthySession = load_scrapling_fetchers()

    keywords = read_keywords(KEYWORDS_FILE)
    all_results: list[BusinessRecord] = []
    results_by_keyword: dict[str, list[BusinessRecord]] = defaultdict(list)

    with StealthySession(
        headless=HEADLESS,
        real_chrome=REAL_CHROME,
        disable_resources=False,
        network_idle=True,
        google_search=True,
    ) as maps_session, FetcherSession(impersonate="chrome") as http_session:
        for keyword in keywords:
            place_links = collect_place_links(maps_session, keyword)
            seen_in_keyword: set[tuple[str, str, str, str]] = set()

            for idx, place_url in enumerate(place_links, start=1):
                print(f"    [{idx}/{len(place_links)}] {place_url}")
                record = scrape_business(maps_session, http_session, keyword, place_url)
                if not record:
                    continue

                dedupe_key = record_unique_key(record)
                if dedupe_key in seen_in_keyword:
                    continue

                seen_in_keyword.add(dedupe_key)
                results_by_keyword[keyword].append(record)
                all_results.append(record)
                print(f"        -> {record.name} | {record.phone} | {record.email}")

    save_keyword_excel_files(results_by_keyword)

    unique_all_results = dedupe_records(all_results)
    write_xlsx(records_to_rows(unique_all_results), ALL_OUTPUT_FILE, sheet_name="all")

    print(f"\n✅ تم حفظ {len(unique_all_results)} نتيجة غير مكررة في الملف: {ALL_OUTPUT_FILE.resolve()}")

if __name__ == "__main__":
    main()
