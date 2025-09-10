import urllib.parse
import re
import chardet
from bs4 import BeautifulSoup
import logging
import pandas as pd
import requests
import json
import asyncio
import time
import os
from .LR import LR  # Assuming LR is in icon_image_lib
from app.email_utils import send_email  # Adjust import based on your project structure
# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_source_url(s):
    """Clean and decode URL string by replacing encoded characters."""
    if not isinstance(s, str):
        return s
    simplified_str = s.replace('\\\\', '')
    replacements = {
        'u0026': '&', 'u003d': '=', 'u003f': '?', 'u0020': ' ', 'u0025': '%', 'u002b': '+', 'u003c': '<',
        'u003e': '>', 'u0023': '#', 'u0024': '$', 'u002f': '/', 'u005c': '\\', 'u007c': '|', 'u002d': '-',
        'u003a': ':', 'u003b': ';', 'u002c': ',', 'u002e': '.', 'u0021': '!', 'u0040': '@', 'u005e': '^',
        'u0060': '`', 'u007b': '{', 'u007d': '}', 'u005b': '[', 'u005d': ']', 'u002a': '*', 'u0028': '(',
        'u0029': ')'
    }
    for encoded, decoded in replacements.items():
        simplified_str = simplified_str.replace(encoded, decoded)
    return simplified_str

def clean_image_url(url):
    """Extract the base image URL without query parameters if it matches image extensions."""
    if not isinstance(url, str):
        return url
    pattern = re.compile(r'(.*\.(?:png|jpg|jpeg|gif|webp|avif|svg))(?:\?.*)?', re.IGNORECASE)
    match = pattern.match(url)
    return match.group(1) if match else url

def extract_true_url_from_wrapper(url_string: str, logger=None) -> str:
    """Extracts the true target URL from known wrapper URL patterns."""
    logger = logger or logging.getLogger(__name__)
    if not url_string or not isinstance(url_string, str):
        return url_string

    try:
        parsed = urllib.parse.urlparse(url_string)
        query_params = urllib.parse.parse_qs(parsed.query)

        if parsed.path.endswith("/_next/image") and 'url' in query_params:
            inner_url = query_params['url'][0]
            logger.debug(f"Extracted from _next/image wrapper: {inner_url} (from {url_string})")
            return urllib.parse.unquote_plus(inner_url)

        is_google_domain = parsed.netloc.startswith(('www.google.', 'images.google.'))
        is_relative_google_path = parsed.netloc == '' and parsed.path in ('/url', '/imgres')

        if is_google_domain or is_relative_google_path:
            target_param_key = None
            if 'url' in query_params:
                target_param_key = 'url'
            elif 'q' in query_params:
                target_param_key = 'q'
            elif 'imgurl' in query_params and parsed.path == '/imgres':
                target_param_key = 'imgurl'

            if target_param_key and query_params[target_param_key]:
                inner_url = query_params[target_param_key][0]
                logger.debug(f"Extracted from Google wrapper (param '{target_param_key}'): {inner_url} (from {url_string})")
                return urllib.parse.unquote_plus(inner_url)

        final_url = urllib.parse.unquote_plus(url_string)
        if final_url != url_string and not (is_google_domain or is_relative_google_path):
            logger.debug(f"Unquoted non-wrapper URL: {final_url} (from {url_string})")
        return final_url

    except Exception as e:
        logger.error(f"🔴 Error in extract_true_url_from_wrapper for '{url_string}': {e}")
        try:
            return urllib.parse.unquote_plus(url_string)
        except Exception:
            return url_string

def decode_html_bytes(html_bytes, logger):
    """Decode HTML bytes to string with fallback encoding."""
    try:
        detected_encoding = chardet.detect(html_bytes)['encoding'] or 'utf-8'
        logger.debug(f"Detected encoding: {detected_encoding}")
        return html_bytes.decode(detected_encoding, errors='replace')
    except Exception as e:
        logger.error(f"🔴 Decoding error: {e}", exc_info=True)
        return html_bytes.decode('utf-8', errors='replace')

def get_original_images(html_bytes, logger=None):
    """Extract image data from Google image search HTML."""
    logger = logger or logging.getLogger(__name__)
    html_content = decode_html_bytes(html_bytes, logger)

    start_tag = 'FINANCE",[22,1]]]]]'
    end_tag = ':[null,null,null,"glbl'
    matched_google_image_data = LR().get(html_content, start_tag, end_tag)
    
    if 'Error' in matched_google_image_data or not matched_google_image_data:
        logger.warning('Main results tags not found or no data extracted (LR step)')
        return [], [], [], []
    
    thumbnails_data_str = str(matched_google_image_data).replace('\u003d', '=').replace('\u0026', '&')
    if '"2003":' not in thumbnails_data_str:
        logger.warning('No "2003" tag found in main thumbnails data')
        return [], [], [], []
    
    raw_thumbs = re.findall(r'\[\"(https:\/\/encrypted-tbn0\.gstatic\.com\/images\?.*?)\",\d+,\d+\]', thumbnails_data_str)
    main_thumbs = [clean_source_url(url) for url in raw_thumbs]

    main_descriptions = re.findall(r'"2003":\[null,"[^"]*","[^"]*","(.*?)"', thumbnails_data_str)
    
    raw_source_urls = re.findall(r'"2003":\[null,"[^"]*","(.*?)"', thumbnails_data_str)
    main_source_urls = [extract_true_url_from_wrapper(clean_source_url(url), logger) for url in raw_source_urls]
    
    removed_thumbs_data = re.sub(r'\[\"(https:\/\/encrypted-tbn0\.gstatic\.com\/images\?.*?)\",\d+,\d+\]', "", thumbnails_data_str)
    raw_main_image_urls = re.findall(r"(?:|,),\[\"(https:|http.*?)\",\d+,\d+\]", removed_thumbs_data)
    main_image_urls = [clean_image_url(extract_true_url_from_wrapper(clean_source_url(url), logger)) for url in raw_main_image_urls]
    
    if not main_image_urls and main_thumbs:
        logger.info("No main_image_urls found, using main_thumbs as fallback for image URLs.")
        main_image_urls = [clean_image_url(thumb_url) for thumb_url in main_thumbs]

    min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
    main_image_urls = main_image_urls[:min_length][:5]
    main_descriptions = main_descriptions[:min_length][:5]
    main_source_urls = main_source_urls[:min_length][:5]
    main_thumbs = main_thumbs[:min_length][:5]

    logger.debug(f"GetOriginal: URLs={len(main_image_urls)}, Desc={len(main_descriptions)}, Sources={len(main_source_urls)}, Thumbs={len(main_thumbs)}")
    return main_image_urls, main_descriptions, main_source_urls, main_thumbs

def process_search_result(image_html_bytes: bytes, entry_id: int, logger=None) -> pd.DataFrame:
    """Process HTML bytes from a search result to extract image data into a DataFrame."""
    logger = logger or logging.getLogger(__name__)
    if not logger.handlers:
        logger.addHandler(logging.StreamHandler())
        logger.setLevel(logging.INFO)

    if not isinstance(image_html_bytes, bytes) or not image_html_bytes:
        logger.error(f"EntryID {entry_id}: Invalid or empty image_html_bytes.")
        return pd.DataFrame()
    if not isinstance(entry_id, int):
        logger.error(f"Invalid entry_id type: {type(entry_id)}. Expected int.")
        return pd.DataFrame()

    try:
        final_urls, final_descriptions, final_sources, final_thumbs = get_original_images(image_html_bytes, logger)
        all_lists = [final_urls, final_descriptions, final_sources, final_thumbs]

        if not all(isinstance(lst, list) for lst in all_lists):
            logger.error(f"EntryID {entry_id}: get_original_images returned invalid types: {[type(lst) for lst in all_lists]}")
            return pd.DataFrame()

        if not any(all_lists):
            logger.warning(f"EntryID {entry_id}: No data extracted.")
            return pd.DataFrame()

        max_len = max(len(lst) for lst in all_lists)
        if any(len(lst) != max_len for lst in [final_urls, final_descriptions, final_sources, final_thumbs]):
            logger.warning(f"EntryID {entry_id}: List lengths: URLs={len(final_urls)}, "
                          f"Descriptions={len(final_descriptions)}, Sources={len(final_sources)}, "
                          f"Thumbnails={len(final_thumbs)}. Padding to {max_len}.")

        df = pd.DataFrame({
            'EntryID': [entry_id] * max_len,
            'ImageUrl': (final_urls + [None] * max_len)[:max_len],
            'ImageDesc': (final_descriptions + [None] * max_len)[:max_len],
            'ImageSource': (final_sources + [None] * max_len)[:max_len],
            'ImageUrlThumbnail': (final_thumbs + [None] * max_len)[:max_len]
        })

        logger.info(f"Processed EntryID {entry_id} with {len(df)} images.")
        return df

    except Exception as e:
        logger.error(f"EntryID {entry_id}: Unexpected error: {str(e)}")
        return pd.DataFrame()

def process_api_image_results(json_data, entry_id: int, logger=None) -> pd.DataFrame:
    """Process JSON from Google Custom Search API into a DataFrame."""
    logger = logger or logging.getLogger(__name__)
    if not logger.handlers:
        logger.addHandler(logging.StreamHandler())
        logger.setLevel(logging.INFO)

    if not isinstance(entry_id, int):
        logger.error(f"Invalid entry_id type: {type(entry_id)}. Expected int.")
        return pd.DataFrame()

    try:
        if isinstance(json_data, str):
            json_data = json.loads(json_data)

        if not isinstance(json_data, (list, dict)):
            logger.error(f"EntryID {entry_id}: Invalid JSON data type: {type(json_data)}")
            return pd.DataFrame()

        items = json_data if isinstance(json_data, list) else json_data.get("items", [])

        if not items:
            logger.warning(f"EntryID {entry_id}: No items found in JSON data.")
            return pd.DataFrame()

        final_urls = []
        final_descriptions = []
        final_sources = []
        final_thumbs = []

        for item in items[:5]:
            image_url = clean_image_url(extract_true_url_from_wrapper(clean_source_url(item.get("ImageUrl", "No image URL")), logger))
            description = item.get("ImageDesc", "No description")
            source = extract_true_url_from_wrapper(clean_source_url(item.get("ImageSource", "No source")), logger)
            thumb = clean_source_url(item.get("ImageUrlThumbnail", "No thumbnail URL"))

            final_urls.append(image_url)
            final_descriptions.append(description)
            final_sources.append(source)
            final_thumbs.append(thumb)

        max_len = max(len(final_urls), len(final_descriptions), len(final_sources), len(final_thumbs))
        if any(len(lst) != max_len for lst in [final_urls, final_descriptions, final_sources, final_thumbs]):
            logger.warning(f"EntryID {entry_id}: List lengths: URLs={len(final_urls)}, "
                          f"Descriptions={len(final_descriptions)}, Sources={len(final_sources)}, "
                          f"Thumbnails={len(final_thumbs)}. Padding to {max_len}.")

        df = pd.DataFrame({
            'EntryID': [entry_id] * max_len,
            'ImageUrl': (final_urls + [None] * max_len)[:max_len],
            'ImageDesc': (final_descriptions + [None] * max_len)[:max_len],
            'ImageSource': (final_sources + [None] * max_len)[:max_len],
            'ImageUrlThumbnail': (final_thumbs + [None] * max_len)[:max_len]
        })

        logger.info(f"Processed EntryID {entry_id} with {len(df)} images.")
        return df

async def fetch_and_process_images(query: str, entry_id: int, search_type: str = "image", worker_url: str = "https://browser-worker.nik-97d.workers.dev") -> pd.DataFrame:
    """Fetch image search results from Cloudflare Worker, process into a DataFrame, and send email."""
    logger = logger or logging.getLogger(__name__)
    try:
        # Construct URL
        encoded_query = urllib.parse.quote_plus(query)
        url = f"{worker_url}?q={encoded_query}&searchType={search_type}&entryId={entry_id}"
        logger.info(f"Fetching from Worker: {url}")

        # Fetch JSON
        response = requests.get(url, timeout=10)
        if not response.ok:
            logger.error(f"Failed to fetch from Worker: {response.status_code} {response.text}")
            return pd.DataFrame()

        json_data = response.json()
        df = process_api_image_results(json_data, entry_id, logger)

        if not df.empty:
            # Save DataFrame to CSV
            file_name = f"image_results_{query}_{int(time.time())}.csv"
            local_filename = f"/tmp/{file_name}"
            df.to_csv(local_filename, index=False)
            logger.info(f"Saved DataFrame to {local_filename}")

            # Send email (assuming send_email is imported elsewhere)
            file_id = entry_id
            success = await send_email(
                to_emails='nik@iconluxurygroup.com',
                subject=f'MSRP File Processed: {file_name}',
                file_path=local_filename,
                job_id=file_id
            )
            if success:
                logger.info(f"Email sent successfully for job ID {file_id}")
            else:
                logger.error(f"Failed to send email for job ID {file_id}")

        return df

    except Exception as e:
        logger.error(f"Error fetching/processing query '{query}': {str(e)}")
        return pd.DataFrame()