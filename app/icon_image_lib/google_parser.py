import urllib.parse
import re
import chardet
from bs4 import BeautifulSoup
import logging
import pandas as pd
import requests
import json
import asyncio
import httpx
import time
import os
from .LR import LR  # Assuming LR is in icon_image_lib
#onfigure logging
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
    """Extract image data from Google image search HTML. Deprecated: Use fetch_and_process_images instead."""
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

def get_results_page_results(html_bytes, final_urls, final_descriptions, final_sources, final_thumbs, logger=None):
    """Extract additional image data from results page HTML without base64. Deprecated: Use fetch_and_process_images instead."""
    logger = logger or logging.getLogger(__name__)
    html_content = decode_html_bytes(html_bytes, logger)

    soup = BeautifulSoup(html_content, 'html.parser')
    result_divs = soup.find_all('div', class_='H8Rx8c')

    if not result_divs:
        result_divs = soup.find_all('div', class_='isv motiva_DRHBO')
        if not result_divs:
            result_divs = soup.find_all('a', class_='isv-r')
            if not result_divs:
                logger.warning("No primary result item divs (H8Rx8c, isv motiva_DRHBO, isv-r) found")
                return final_urls, final_descriptions, final_sources, final_thumbs

    logger.info(f"Found {len(result_divs)} potential items in additional results page.")

    for item_container in result_divs:
        if len(final_urls) >= 100:
            break

        img_tag = item_container.find('img', class_=lambda c: c and any(cls in c for cls in ['rg_i', 'n3VNCb', 'YQ4gaf', 'gdOPf', 'uhHOwf', 'ez24Df']))
        raw_thumb_url = None
        if img_tag:
            raw_thumb_url = img_tag.get('src') or img_tag.get('data-src')
        
        if not raw_thumb_url or 'data:image' in raw_thumb_url:
            logger.debug("No valid thumbnail URL or base64 src in result item, skipping.")
            continue
        
        thumb = clean_source_url(raw_thumb_url)
        final_thumbs.append(thumb)

        desc_element = item_container.find(['div', 'span', 'a'], class_=lambda c: c and any(cls in c for cls in ['bytUYc', 'VFACy', 'mVDVAe', 'VwiC3b']))
        description = desc_element.get_text(strip=True) if desc_element else 'No description'
        final_descriptions.append(description)

        source_element = item_container.find(['cite', 'div', 'span'], class_=lambda c: c and any(cls in c for cls in ['VuuXrf', 'SW5pqf', 'qLRx3b']))
        raw_source_text = source_element.get_text(strip=True) if source_element else 'No source'
        
        cleaned_source_text = clean_source_url(raw_source_text)
        if cleaned_source_text.startswith(('http', '/', 'www.')):
            source = extract_true_url_from_wrapper(cleaned_source_text, logger)
        else:
            source = cleaned_source_text
        final_sources.append(source)

        link_tag = None
        if item_container.name == 'a':
            link_tag = item_container
        else:
            link_tag = item_container.find('a', class_=lambda c: c and any(cls in c for cls in ['zReHs', 'VFACy', 'isv-r']))

        raw_href = link_tag.get('href') if link_tag and link_tag.get('href') else None
        if not raw_href and img_tag:
            raw_href = img_tag.get('data-actualn3r')

        image_url_final = 'No image URL'
        if raw_href:
            url1 = clean_source_url(raw_href)
            url2 = extract_true_url_from_wrapper(url1, logger)
            image_url_final = clean_image_url(url2)
        final_urls.append(image_url_final)

    if len(final_urls) > 100:
        final_urls = final_urls[:100]
        final_descriptions = final_descriptions[:100]
        final_sources = final_sources[:100]
        final_thumbs = final_thumbs[:100]
        logger.debug("Capped total results from additional page processing at 100")

    logger.debug(f"GetResultsPage: Total after append: URLs={len(final_urls)}")
    return final_urls, final_descriptions, final_sources, final_thumbs

def process_search_result(image_html_bytes: bytes, entry_id: int, logger=None) -> pd.DataFrame:
    """Process HTML bytes from a search result to extract image data into a DataFrame. Deprecated: Use fetch_and_process_images instead."""
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
        logger.error(f"Error processing search result for EntryID {entry_id}: {str(e)}")
        return pd.DataFrame()

async def submit_tunnel_job(client, url: str, logger=None) -> str:
    """Submit URL to tunnel and return job_id."""
    if not url or not url.startswith(('http', 'https')):
        return None
    
    submit_url = "https://tunnel.publicwebarchive.com/tunnel"
    try:
        if logger: logger.debug(f"Submitting {url[:50]}... to tunnel")
        resp = await client.get(submit_url, params={"url": url})
        if resp.status_code != 200:
            if logger: logger.warning(f"Tunnel submit failed: {resp.status_code}")
            return None
        
        data = resp.json()
        job_id = data.get("job_id")
        if not job_id:
            # Check if it returned a direct url (in case server logic changed)
            if data.get("url"): 
                return "COMPLETED:" + data.get("url")
            
            if logger: logger.warning(f"No job_id in tunnel response for {url}")
            return None
        
        if logger: logger.info(f"Tunnel job started: {job_id}")
        return job_id

    except Exception as e:
        if logger: logger.error(f"Tunnel submit error: {e}")
        return None

async def poll_tunnel_job(client, job_id: str, logger=None) -> str:
    """Poll tunnel for job completion."""
    if not job_id:
        return None
    
    if job_id.startswith("COMPLETED:"):
        return job_id.replace("COMPLETED:", "")
        
    status_url = f"https://tunnel.publicwebarchive.com/tunnel/status/{job_id}"
    
    for i in range(12): # Poll for connection
        await asyncio.sleep(3)
        try:
            poll_resp = await client.get(status_url)
            if poll_resp.status_code == 200:
                poll_data = poll_resp.json()
                status = poll_data.get("status")
                
                if status == "completed":
                    result_url = poll_data.get("html_source_url") or poll_data.get("result_url") or poll_data.get("final_url") or poll_data.get("url")
                    if logger: logger.info(f"Tunnel job {job_id} completed. Result: {result_url}")
                    return result_url
                elif status == "failed":
                        if logger: logger.warning(f"Tunnel job {job_id} failed")
                        return None
            elif poll_resp.status_code == 404:
                    pass
        except Exception as e:
            if logger: logger.debug(f"Polling error for {job_id}: {e}")
            pass
            
    if logger: logger.warning(f"Tunnel timed out for job {job_id}")
    return None

async def process_api_image_results(json_data, entry_id: int, logger=None) -> pd.DataFrame:
    """Process JSON from image API into a DataFrame."""
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
        
        async with httpx.AsyncClient(verify=False, timeout=60.0) as client:
            items_to_process = items[:5]
            
            # Phase 1: Submit
            submit_tasks = []
            for item in items_to_process:
                raw_source = extract_true_url_from_wrapper(clean_source_url(item.get("ImageSource", "No source")), logger)
                submit_tasks.append(submit_tunnel_job(client, raw_source, logger))
            
            job_ids = await asyncio.gather(*submit_tasks)
            
            # Phase 2: Poll
            poll_tasks = []
            for job_id in job_ids:
                poll_tasks.append(poll_tunnel_job(client, job_id, logger))
            
            resolved_results = await asyncio.gather(*poll_tasks)
            
            for i, item in enumerate(items_to_process):
                image_url = clean_image_url(extract_true_url_from_wrapper(clean_source_url(item.get("ImageUrl", "No image URL")), logger))
                description = item.get("ImageDesc", "No description")
                raw_source = extract_true_url_from_wrapper(clean_source_url(item.get("ImageSource", "No source")), logger)
                thumb = clean_source_url(item.get("ImageUrlThumbnail", "No thumbnail URL"))
                
                resolved_source = resolved_results[i]
                final_source = resolved_source if resolved_source else raw_source

                final_urls.append(image_url)
                final_descriptions.append(description)
                final_sources.append(final_source)
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
    except Exception as e:
        logger.error(f"Error processing API results for EntryID {entry_id}: {str(e)}")
        return pd.DataFrame()

import logging
import requests
import pandas as pd

import logging
import requests
import pandas as pd

async def fetch_and_process_images(
    query: str,
    entry_id: int,
    brand: str = None,
    search_type: str = "image",
    logger: logging.Logger = None,
) -> pd.DataFrame:
    logger = logger or logging.getLogger(__name__)
    if not logger.handlers:
        logger.addHandler(logging.StreamHandler())
        logger.setLevel(logging.INFO)
    
    try:
        worker_url: str = "https://browser-worker.nik-97d.workers.dev"
        params = {
            "q": query.strip('"'),  # Remove quotes if present
            "searchType": search_type,
            "entryId": str(entry_id)  # Convert to string for query params
        }
        if brand:
            params["brand"] = brand  # Include brand if provided
        logger.info(f"GET request to Worker: {worker_url} with params: {params}")
        
        response = requests.get(worker_url, params=params, timeout=10)
        response.raise_for_status()  # Raise for non-2xx status codes
        
        json_data = response.json()
        df = await process_api_image_results(json_data, entry_id, logger)
        return df
    
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP error fetching query '{query}': {response.status_code} {response.text[:500]}")
        return pd.DataFrame()
    except requests.exceptions.RequestException as e:
        logger.error(f"Network error fetching query '{query}': {str(e)}")
        return pd.DataFrame()
    except ValueError as e:
        logger.error(f"Error parsing JSON response for query '{query}': {str(e)}")
        return pd.DataFrame()