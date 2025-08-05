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
import json
import re
import logging
from typing import List, Tuple, Optional
from bs4 import BeautifulSoup

import json
import re
import logging
from typing import List, Tuple, Optional
from bs4 import BeautifulSoup
import urllib.parse
from typing import Optional, Tuple, List
from bs4 import BeautifulSoup
import logging
import re
import json

def get_original_images(html_bytes: bytes, logger_instance: Optional[logging.Logger] = None) -> Tuple[List[str], List[str], List[str], List[str]]:
    """Extract image data from Google image search HTML in grid order."""
    logger_instance = logger_instance or logging.getLogger(__name__)
    if not html_bytes:
        logger_instance.warning("Empty HTML bytes provided to get_original_images")
        return [], [], [], []

    try:
        # Decode HTML bytes to string
        html_content = html_bytes.decode('utf-8', errors='ignore')
        soup = BeautifulSoup(html_content, 'html.parser')
        
        main_image_urls = []
        main_descriptions = []
        main_source_urls = []
        main_thumbs = []
        
        # Process metadata divs
        meta_divs = soup.select('div[data-ved]')
        logger_instance.debug(f"Found {len(meta_divs)} potential metadata entries")
        
        for idx, div in enumerate(meta_divs):
            try:
                # Find imgres link for metadata
                img_link = div.select_one('a[href^="/imgres"]')
                if not img_link:
                    logger_instance.debug(f"No imgres link in div {idx}")
                    main_image_urls.append('')
                    main_descriptions.append('')
                    main_source_urls.append('')
                    main_thumbs.append('')
                    continue
                
                # Extract URL parameters
                href = img_link.get('href', '')
                params = urllib.parse.parse_qs(urllib.parse.urlparse(href).query)
                
                # Image URL
                full_url = params.get('imgurl', [''])[0]
                if full_url:
                    full_url = clean_source_url(full_url)
                    full_url = extract_true_url_from_wrapper(full_url, logger_instance)
                    full_url = clean_image_url(full_url)
                main_image_urls.append(full_url if full_url else '')
                
                # Source URL
                source_url = params.get('imgrefurl', [''])[0]
                if source_url:
                    source_url = clean_source_url(source_url)
                    source_url = extract_true_url_from_wrapper(source_url, logger_instance)
                main_source_urls.append(source_url if source_url else '')
                
                # Description from div or img alt
                desc_div = div.select_one('div.toI8Rb.OSrXXb')
                desc = desc_div.text.strip() if desc_div else (img_link.select_one('img').get('alt', '') if img_link.select_one('img') else '')
                main_descriptions.append(desc if desc else '')
                
                # Thumbnail
                img_tag = img_link.select_one('img.YQ4gaf')
                thumb_src = img_tag.get('src') or img_tag.get('data-src') or img_tag.get('data-iurl') or '' if img_tag else ''
                if thumb_src and not thumb_src.startswith('data:image'):
                    thumb_src = clean_source_url(thumb_src)
                    thumb_src = extract_true_url_from_wrapper(thumb_src, logger_instance)
                main_thumbs.append(thumb_src if thumb_src else '')
                
            except Exception as e:
                logger_instance.error(f"Error processing metadata div {idx}: {e}", exc_info=True)
                main_image_urls.append('')
                main_descriptions.append('')
                main_source_urls.append('')
                main_thumbs.append('')
        
        # Fallback to img tags if thumbnails are incomplete
        if not any(main_thumbs) or len([t for t in main_thumbs if t]) < len([u for u in main_image_urls if u]):
            logger_instance.info("Thumbnails incomplete; falling back to img tags")
            img_tags = soup.select('img.YQ4gaf, img[class*="rg_i"], img[data-src], img[src], img[data-iurl]')
            for idx, img in enumerate(img_tags):
                try:
                    thumb_src = img.get('src') or img.get('data-src') or img.get('data-iurl') or ''
                    if thumb_src and not thumb_src.startswith('data:image') and any(thumb_src.endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif']):
                        thumb_src = clean_source_url(thumb_src)
                        thumb_src = extract_true_url_from_wrapper(thumb_src, logger_instance)
                        main_thumbs.append(thumb_src)
                    else:
                        main_thumbs.append('')
                except Exception as e:
                    logger_instance.error(f"Error processing img tag {idx}: {e}", exc_info=True)
                    main_thumbs.append('')
        
        # Ensure consistent list lengths and limit to top 5 results
        min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
        main_image_urls = main_image_urls[:min_length][:5]
        main_descriptions = main_descriptions[:min_length][:5]
        main_source_urls = main_source_urls[:min_length][:5]
        main_thumbs = main_thumbs[:min_length][:5]
        
        # Log the source of the results
        valid_thumbs = len([t for t in main_thumbs if t])
        logger_instance.debug(f"Extracted: URLs={len(main_image_urls)}, Desc={len(main_descriptions)}, Sources={len(main_source_urls)}, Thumbs={valid_thumbs}")
  
        return main_image_urls, main_descriptions, main_source_urls, main_thumbs
    except Exception as e:
        logger_instance.error(f"Critical error in get_original_images: {e}", exc_info=True)
        return [], [], [], []

def clean_source_url(url: str) -> str:
    """Clean URL by removing unnecessary parameters."""
    try:
        return re.sub(r'\?.*$', '', url).strip()
    except Exception:
        return url

def clean_image_url(url: str) -> str:
    """Clean image URL by ensuring it points to a valid image."""
    try:
        if not url.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')):
            return ''
        return url
    except Exception:
        return url

def extract_true_url_from_wrapper(url: str, logger_instance: logging.Logger) -> str:
    """Extract the true URL from Google's wrapper URLs."""
    try:
        match = re.search(r'imgurl=([^&]+)', url)
        if match:
            from urllib.parse import unquote
            return unquote(match.group(1))
        return url
    except Exception as e:
        logger_instance.warning(f"Failed to extract true URL: {e}")
        return url

def process_search_result(image_html_bytes: bytes, entry_id: int, logger_instance: Optional[logging.Logger] = None, save_to_file: Optional[str] = None) -> pd.DataFrame:
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
        # Ensure consistent lengths
        lengths = [len(final_urls), len(final_descriptions), len(final_sources), len(final_thumbs)]
        min_len = min(lengths)
        if any(l != min_len for l in lengths):
            logger_instance.warning(
                f"EntryID {entry_id}: Data lists had different lengths {lengths}. Truncated to min length: {min_len}"
            )
        df = pd.DataFrame({
            'EntryID': [entry_id] * min_len,
            'ImageUrl': final_urls[:min_len],
            'ImageDesc': final_descriptions[:min_len],
            'ImageSource': final_sources[:min_len],
            'ImageUrlThumbnail': final_thumbs[:min_len]
        })
        # Remove rows with empty ImageUrl
        df = df[df['ImageUrl'].str.strip() != '']
        if df.empty:
            logger_instance.warning(f"EntryID {entry_id}: No valid images after filtering")
            return pd.DataFrame()
        # Save to CSV if requested
        if save_to_file:
            try:
                df.to_csv(save_to_file, mode='a', index=False, header=not pd.io.common.file_exists(save_to_file))
                logger_instance.info(f"EntryID {entry_id}: Saved {len(df)} images to {save_to_file}")
            except Exception as e:
                logger_instance.error(f"EntryID {entry_id}: Failed to save to {save_to_file}: {e}", exc_info=True)
        logger_instance.info(f"Processed EntryID {entry_id} with {len(df)} images")
        return df
    except Exception as e:
        logger_instance.error(f"Critical error processing EntryID {entry_id}: {e}", exc_info=True)
        return pd.DataFrame()