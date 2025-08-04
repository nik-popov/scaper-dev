import logging
from typing import List, Tuple
from urllib.parse import urlparse

from bs4 import BeautifulSoup

from icon_image_lib.google_parser import (
    clean_image_url,
    clean_source_url,
    extract_true_url_from_wrapper,
)


def get_results_page_results(html_bytes: bytes, logger: logging.Logger | None = None) -> Tuple[List[str], List[str], List[str], List[str]]:
    """Extract organic search results from Google search results HTML.

    Args:
        html_bytes: Raw HTML bytes of the Google search results page.
        logger: Optional logger instance for debugging.

    Returns:
        Tuple of (urls, titles, sources, thumbnails) as lists.
    """
    logger = logger or logging.getLogger(__name__)
    try:
        html_content = html_bytes.decode("utf-8", errors="ignore")
    except Exception as exc:  # pragma: no cover - defensive
        logger.error(f"Failed to decode HTML bytes: {exc}")
        return [], [], [], []

    soup = BeautifulSoup(html_content, "html.parser")
    final_urls: List[str] = []
    final_titles: List[str] = []
    final_sources: List[str] = []
    final_thumbs: List[str] = []

    result_divs = soup.find_all("div", class_="MjjYud")
    if not result_divs:
        logger.warning("No primary result divs (MjjYud) found, trying alternative selectors")
        result_divs = soup.find_all("div", class_=lambda c: c and any(cls in c for cls in ["g", "tF2Cxc"]))

    logger.info(f"Found {len(result_divs)} potential result items.")

    for item_container in result_divs:
        if len(final_urls) >= 100:
            logger.debug("Reached result cap of 100, stopping.")
            break

        link_tag = item_container.find("a", href=True)
        raw_href = link_tag.get("href") if link_tag else None
        url = extract_true_url_from_wrapper(raw_href, logger) if raw_href else "No URL"
        final_urls.append(url)

        title_tag = item_container.find("h3")
        title = title_tag.get_text(strip=True) if title_tag else "No title"
        final_titles.append(title)

        source_tag = item_container.find("cite")
        source = (
            source_tag.get_text(strip=True).split(" › ")[0]
            if source_tag
            else urlparse(url).netloc or "No source"
        )
        final_sources.append(clean_source_url(source))

        img_tag = item_container.find(
            "img",
            class_=lambda c: c and any(cls in c for cls in ["rg_i", "n3VNCb", "YQ4gaf"]),
        )
        thumb_url = img_tag.get("src") or img_tag.get("data-src") if img_tag else None
        thumb = (
            clean_image_url(thumb_url)
            if thumb_url and "data:image" not in thumb_url
            else "No thumbnail"
        )
        final_thumbs.append(thumb)

    logger.debug(
        "Parsed %d results: URLs=%d, Titles=%d, Sources=%d, Thumbnails=%d",
        len(final_urls),
        len(final_titles),
        len(final_sources),
        len(final_thumbs),
    )
    return final_urls, final_titles, final_sources, final_thumbs
