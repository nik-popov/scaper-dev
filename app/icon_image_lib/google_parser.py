import urllib.parse
import logging
import json
from bs4 import BeautifulSoup
import re
import chardet
import pandas as pd
from typing import Optional, Tuple, List
from icon_image_lib.LR import LR  # Assuming this import is correct

# Configure logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def clean_source_url(s: Optional[str]) -> Optional[str]:
    """Clean and decode URL string by replacing encoded characters."""
    if not isinstance(s, str):
        logger.warning(f"Invalid input type for clean_source_url: {type(s)}. Expected str, got {s}")
        return s
    try:
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
        logger.debug(f"Cleaned URL: {simplified_str}")
        return simplified_str
    except Exception as e:
        logger.error(f"Error cleaning URL '{s}': {e}", exc_info=True)
        return s

def clean_image_url(url: Optional[str]) -> Optional[str]:
    """Extract the base image URL without query parameters if it matches image extensions."""
    if not isinstance(url, str):
        logger.warning(f"Invalid input type for clean_image_url: {type(url)}. Expected str, got {url}")
        return url
    try:
        pattern = re.compile(r'(.*\.(?:png|jpg|jpeg|gif|webp|avif|svg))(?:\?.*)?', re.IGNORECASE)
        match = pattern.match(url)
        cleaned_url = match.group(1) if match else url
        if cleaned_url != url:
            logger.debug(f"Cleaned image URL: {cleaned_url} (from {url})")
        return cleaned_url
    except Exception as e:
        logger.error(f"Error cleaning image URL '{url}': {e}", exc_info=True)
        return url

def extract_true_url_from_wrapper(url_string: Optional[str], logger_instance: Optional[logging.Logger] = None) -> Optional[str]:
    """Extracts the true target URL from known wrapper URL patterns."""
    logger_instance = logger_instance or logger
    if not isinstance(url_string, str) or not url_string:
        logger_instance.warning(f"Invalid or empty URL in extract_true_url_from_wrapper: {url_string}")
        return url_string

    try:
        parsed = urllib.parse.urlparse(url_string)
        query_params = urllib.parse.parse_qs(parsed.query)

        # Next.js _next/image wrapper
        if parsed.path.endswith("/_next/image") and 'url' in query_params:
            inner_url = query_params['url'][0]
            logger_instance.debug(f"Extracted from _next/image wrapper: {inner_url} (from {url_string})")
            return urllib.parse.unquote_plus(inner_url)

        # Google wrapper
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
                logger_instance.debug(f"Extracted from Google wrapper (param '{target_param_key}'): {inner_url} (from {url_string})")
                return urllib.parse.unquote_plus(inner_url)

        # Unquote non-wrapper URL
        final_url = urllib.parse.unquote_plus(url_string)
        if final_url != url_string:
            logger_instance.debug(f"Unquoted non-wrapper URL: {final_url} (from {url_string})")
        return final_url
    except Exception as e:
        logger_instance.error(f"Error extracting URL from wrapper '{url_string}': {e}", exc_info=True)
        try:
            return urllib.parse.unquote_plus(url_string)
        except Exception:
            return url_string

def decode_html_bytes(html_bytes: bytes, logger_instance: Optional[logging.Logger] = None) -> str:
    """Decode HTML bytes to string with fallback encoding."""
    logger_instance = logger_instance or logger
    if not isinstance(html_bytes, bytes):
        logger_instance.error(f"Invalid input type for decode_html_bytes: {type(html_bytes)}. Expected bytes")
        return ""
    try:
        detected = chardet.detect(html_bytes)
        encoding = detected['encoding'] or 'utf-8'
        confidence = detected.get('confidence', 0.0)
        logger_instance.debug(f"Detected encoding: {encoding} (confidence: {confidence:.2f})")
        decoded = html_bytes.decode(encoding, errors='replace')
        return decoded
    except Exception as e:
        logger_instance.error(f"Error decoding HTML bytes: {e}", exc_info=True)
        return html_bytes.decode('utf-8', errors='replace')

def get_original_images(html_bytes: bytes, logger_instance: Optional[logging.Logger] = None) -> Tuple[List[str], List[str], List[str], List[str]]:
    """Extract image data from Google image search HTML in grid order."""
    logger_instance = logger_instance or logger
    if not html_bytes:
        logger_instance.warning("Empty HTML bytes provided to get_original_images")
        return [], [], [], []

    try:
        html_content = decode_html_bytes(html_bytes, logger_instance)
        soup = BeautifulSoup(html_content, 'html.parser')
        
        main_image_urls = []
        main_descriptions = []
        main_source_urls = []
        main_thumbs = []
        
        meta_divs = soup.select('div.rg_meta.notranslate')
        logger_instance.debug(f"Found {len(meta_divs)} metadata entries")
        
        for idx, div in enumerate(meta_divs):
            try:
                meta_json_str = div.text.strip()
                meta_json_str = re.sub(r'\\u([0-9a-fA-F]{4})', lambda m: chr(int(m.group(1), 16)), meta_json_str)
                meta_data = json.loads(meta_json_str)
                
                full_url = meta_data.get('ou', '')
                if full_url:
                    full_url = clean_source_url(full_url)
                    full_url = extract_true_url_from_wrapper(full_url, logger_instance)
                    full_url = clean_image_url(full_url)
                    main_image_urls.append(full_url)
                
                desc = meta_data.get('pt', '') or meta_data.get('s', '')
                main_descriptions.append(desc)
                
                source_url = meta_data.get('ru', '')
                if source_url:
                    source_url = clean_source_url(source_url)
                    source_url = extract_true_url_from_wrapper(source_url, logger_instance)
                    main_source_urls.append(source_url)
                
                thumb_url = meta_data.get('tu', '') or meta_data.get('st', '')
                if thumb_url and thumb_url.startswith('http'):
                    thumb_url = clean_source_url(thumb_url)
                    thumb_url = extract_true_url_from_wrapper(thumb_url, logger_instance)
                    main_thumbs.append(thumb_url)
                else:
                    main_thumbs.append('')
                    
            except json.JSONDecodeError as e:
                logger_instance.warning(f"Failed to parse JSON in metadata div {idx}: {e}")
                main_image_urls.append('')
                main_descriptions.append('')
                main_source_urls.append('')
                main_thumbs.append('')
            except Exception as e:
                logger_instance.error(f"Error processing metadata div {idx}: {e}", exc_info=True)
                main_image_urls.append('')
                main_descriptions.append('')
                main_source_urls.append('')
                main_thumbs.append('')
        
        # Fallback to img tags
        if not any(main_thumbs) or len([t for t in main_thumbs if t]) < len(main_image_urls):
            logger_instance.info("Thumbnails incomplete; falling back to img tags")
            img_tags = soup.select('img.rg_i.Q4LuWd')
            main_thumbs = []
            for idx, img in enumerate(img_tags):
                try:
                    thumb_src = img.get('src') or img.get('data-src', '')
                    if thumb_src:
                        thumb_src = clean_source_url(thumb_src)
                        thumb_src = extract_true_url_from_wrapper(thumb_src, logger_instance)
                        main_thumbs.append(thumb_src)
                    else:
                        main_thumbs.append('')
                except Exception as e:
                    logger_instance.error(f"Error processing img tag {idx}: {e}", exc_info=True)
                    main_thumbs.append('')
        
        min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
        main_image_urls = main_image_urls[:min_length][:5]
        main_descriptions = main_descriptions[:min_length][:5]
        main_source_urls = main_source_urls[:min_length][:5]
        main_thumbs = main_thumbs[:min_length][:5]
        
        logger_instance.debug(f"Extracted: URLs={len(main_image_urls)}, Desc={len(main_descriptions)}, Sources={len(main_source_urls)}, Thumbs={len(main_thumbs)}")
        return main_image_urls, main_descriptions, main_source_urls, main_thumbs
    except Exception as e:
        logger_instance.error(f"Critical error in get_original_images: {e}", exc_info=True)
        return [], [], [], []

def process_search_result(image_html_bytes: bytes, entry_id: int, logger_instance: Optional[logging.Logger] = None) -> pd.DataFrame:
    """Process search result HTML bytes and return a DataFrame with image data."""
    logger_instance = logger_instance or logger
    if not isinstance(image_html_bytes, bytes):
        logger_instance.error(f"Invalid input type for image_html_bytes: {type(image_html_bytes)}. Expected bytes")
        return pd.DataFrame()
    if not isinstance(entry_id, int):
        logger_instance.error(f"Invalid entry_id type: {type(entry_id)}. Expected int")
        return pd.DataFrame()

    try:
        final_urls, final_descriptions, final_sources, final_thumbs = get_original_images(image_html_bytes, logger_instance)
        
        if not final_urls:
            logger_instance.warning(f"EntryID {entry_id}: No images extracted")
            return pd.DataFrame()
        
        min_len = min(len(lst) for lst in [final_urls, final_descriptions, final_sources, final_thumbs])
        if any(len(lst) > min_len for lst in [final_urls, final_descriptions, final_sources, final_thumbs]):
            logger_instance.warning(
                f"EntryID {entry_id}: Data lists had different lengths "
                f"({[len(lst) for lst in [final_urls, final_descriptions, final_sources, final_thumbs]]}). "
                f"Truncated to min length: {min_len}"
            )
        
        df = pd.DataFrame({
            'EntryID': [entry_id] * min_len,
            'ImageUrl': final_urls[:min_len],
            'ImageDesc': final_descriptions[:min_len],
            'ImageSource': final_sources[:min_len],
            'ImageUrlThumbnail': final_thumbs[:min_len]
        })
        
        logger_instance.info(f"Processed EntryID {entry_id} with {len(df)} images")
        return df
    except Exception as e:
        logger_instance.error(f"Critical error processing EntryID {entry_id}: {e}", exc_info=True)
        return pd.DataFrame()