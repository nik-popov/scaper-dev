import urllib.parse
import logging
from bs4 import BeautifulSoup
import re
import chardet
import pandas as pd
from typing import Optional, Tuple, List

# Configure logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def clean_source_url(s: Optional[str]) -> Optional[str]:
    """Clean and decode URL string by replacing encoded characters and removing query parameters."""
    if not isinstance(s, str):
        logger.warning(f"Invalid input type for clean_source_url: {type(s)}. Expected str, got {s}")
        return s
    try:
        s = re.sub(r'\?.*$', '', s).strip()
        replacements = {
            'u0026': '&', 'u003d': '=', 'u003f': '?', 'u0020': ' ', 'u0025': '%', 'u002b': '+', 'u003c': '<',
            'u003e': '>', 'u0023': '#', 'u0024': '$', 'u002f': '/', 'u005c': '\\', 'u007c': '|', 'u002d': '-',
            'u003a': ':', 'u003b': ';', 'u002c': ',', 'u002e': '.', 'u0021': '!', 'u0040': '@', 'u005e': '^',
            'u0060': '`', 'u007b': '{', 'u007d': '}', 'u005b': '[', 'u005d': ']', 'u002a': '*', 'u0028': '(',
            'u0029': ')'
        }
        for encoded, decoded in replacements.items():
            s = s.replace(encoded, decoded).replace('\\\\', '')
        logger.debug(f"Cleaned URL: {s}")
        return s
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
        cleaned_url = match.group(1) if match else ''
        if cleaned_url and cleaned_url != url:
            logger.debug(f"Cleaned image URL: {cleaned_url} (from {url})")
        return cleaned_url if cleaned_url else ''
    except Exception as e:
        logger.error(f"Error cleaning image URL '{url}': {e}", exc_info=True)
        return ''

def extract_true_url_from_wrapper(url_string: Optional[str], logger_instance: Optional[logging.Logger] = None) -> Optional[str]:
    """Extracts the true target URL from known wrapper URL patterns."""
    logger_instance = logger_instance or logger
    if not isinstance(url_string, str) or not url_string:
        logger_instance.warning(f"Invalid or empty URL in extract_true_url_from_wrapper: {url_string}")
        return url_string
    try:
        parsed = urllib.parse.urlparse(url_string)
        query_params = urllib.parse.parse_qs(parsed.query)
        if parsed.path.endswith("/_next/image") and 'url' in query_params:
            inner_url = query_params['url'][0]
            logger_instance.debug(f"Extracted from _next/image wrapper: {inner_url} (from {url_string})")
            return urllib.parse.unquote_plus(inner_url)
        is_google_domain = parsed.netloc.startswith(('www.google.', 'images.google.'))
        is_relative_google_path = parsed.netloc == '' and parsed.path in ('/url', '/imgres')
        if is_google_domain or is_relative_google_path:
            target_param_key = next((k for k in ['url', 'q', 'imgurl', 'data-imgurl'] if k in query_params), None)
            if target_param_key and query_params[target_param_key]:
                inner_url = query_params[target_param_key][0]
                logger_instance.debug(f"Extracted from Google wrapper (param '{target_param_key}'): {inner_url} (from {url_string})")
                return urllib.parse.unquote_plus(inner_url)
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
        if confidence < 0.7:
            for enc in [encoding, 'utf-8', 'iso-8859-1', 'windows-1252']:
                try:
                    decoded = html_bytes.decode(enc, errors='replace')
                    logger_instance.debug(f"Decoded HTML with encoding: {enc} (confidence: {confidence:.2f})")
                    return decoded
                except UnicodeDecodeError:
                    continue
        decoded = html_bytes.decode(encoding, errors='replace')
        logger_instance.debug(f"Decoded HTML with encoding: {encoding} (confidence: {confidence:.2f})")
        return decoded
    except Exception as e:
        logger_instance.error(f"Error decoding HTML bytes: {e}", exc_info=True)
        return html_bytes.decode('utf-8', errors='replace')

def get_original_images(html_bytes: bytes, logger_instance: Optional[logging.Logger] = None) -> Tuple[List[str], List[str], List[str], List[str]]:
    """Extract image data from Google image search HTML in grid order, excluding social media icons and small images."""
    logger_instance = logger_instance or logger
    if not html_bytes:
        logger_instance.warning("Empty HTML bytes provided to get_original_images")
        return [], [], [], []
    try:
        html_content = decode_html_bytes(html_bytes, logger_instance)
        soup = BeautifulSoup(html_content, 'html.parser')
        main_image_urls, main_descriptions, main_source_urls, main_thumbs = [], [], [], []

        # Social media icon patterns to exclude
        social_media_patterns = [
            r'facebook\.com.*\.(png|jpg|jpeg|gif|webp|avif|svg)',
            r'twitter\.com.*\.(png|jpg|jpeg|gif|webp|avif|svg)',
            r'instagram\.com.*\.(png|jpg|jpeg|gif|webp|avif|svg)',
            r'linkedin\.com.*\.(png|jpg|jpeg|gif|webp|avif|svg)',
            r'pinterest\.com.*\.(png|jpg|jpeg|gif|webp|avif|svg)',
            r'\.ico$',  # Favicon files
            r'logo\.(png|jpg|jpeg|gif|webp|avif|svg)',  # Generic logo files
            r'share\.(png|jpg|jpeg|gif|webp|avif|svg)',  # Share icons
            r'icon\.(png|jpg|jpeg|gif|webp|avif|svg)',  # Generic icon files
            r'favicon\.(png|jpg|jpeg|gif|webp|avif|svg)'  # Favicon files
        ]

        # Process metadata divs (image result containers)
        meta_divs = soup.select('div[data-ved], div[class*="eA0Zlc"], div[class*="isv-r"], div[class*="rg_bx"]')
        logger_instance.debug(f"Found {len(meta_divs)} potential metadata entries")
        for idx, div in enumerate(meta_divs):
            try:
                # Find image link
                img_link = div.select_one('a[href^="/imgres"], a[href*="imgurl"], a[class*="EZAeBe"], a[class*="rQEFy"]')
                if not img_link:
                    logger_instance.debug(f"No image link in div {idx}")
                    continue
                href = img_link.get('href', '') or img_link.get('data-href', '')
                params = urllib.parse.parse_qs(urllib.parse.urlparse(href).query) if href else {}
                # Image URL
                full_url = params.get('imgurl', [''])[0] or params.get('data-imgurl', [''])[0]
                if full_url:
                    if any(re.search(pattern, full_url, re.IGNORECASE) for pattern in social_media_patterns):
                        logger_instance.debug(f"Skipping social media icon URL: {full_url}")
                        continue
                    full_url = clean_source_url(full_url)
                    full_url = extract_true_url_from_wrapper(full_url, logger_instance)
                    full_url = clean_image_url(full_url)
                if not full_url:
                    logger_instance.debug(f"No valid image URL in div {idx}")
                    continue
                # Source URL
                source_url = params.get('imgrefurl', [''])[0] or img_link.get('href', '')
                if source_url:
                    source_url = clean_source_url(source_url)
                    source_url = extract_true_url_from_wrapper(source_url, logger_instance)
                # Description
                desc_div = div.select_one('div.toI8Rb.OSrXXb, div[class*="VwiC3b"], div[class*="m3LIae"], span[class*="OSrXXb"]')
                desc = desc_div.text.strip() if desc_div else (img_link.select_one('img').get('alt', '') if img_link.select_one('img') else '')
                # Thumbnail
                img_tag = img_link.select_one('img.YQ4gaf, img[class*="rg_i"], img[class*="sFlh5c"], img[class*="WICvnb"], img[class*="isv-r"]')
                thumb_src = img_tag.get('src') or img_tag.get('data-src') or img_tag.get('data-iurl') or img_tag.get('data-lazy-src') or '' if img_tag else ''
                # Check image dimensions to exclude small icons
                width = img_tag.get('width', '0') if img_tag else '0'
                height = img_tag.get('height', '0') if img_tag else '0'
                try:
                    if int(width) <= 50 or int(height) <= 50:
                        logger_instance.debug(f"Skipping small image (width: {width}, height: {height}): {thumb_src}")
                        continue
                except ValueError:
                    pass
                if thumb_src and not thumb_src.startswith('data:image'):
                    if any(re.search(pattern, thumb_src, re.IGNORECASE) for pattern in social_media_patterns):
                        logger_instance.debug(f"Skipping social media icon thumbnail: {thumb_src}")
                        continue
                    thumb_src = clean_source_url(thumb_src)
                    thumb_src = extract_true_url_from_wrapper(thumb_src, logger_instance)
                    thumb_src = clean_image_url(thumb_src) or full_url
                else:
                    thumb_src = full_url
                main_image_urls.append(full_url)
                main_descriptions.append(desc if desc else '')
                main_source_urls.append(source_url if source_url else '')
                main_thumbs.append(thumb_src)
            except Exception as e:
                logger_instance.error(f"Error processing metadata div {idx}: {e}", exc_info=True)
                continue

        # Fallback to img tags if no valid results
        if not main_image_urls:
            logger_instance.info("No images from metadata; falling back to img tags")
            img_tags = soup.select('img.YQ4gaf, img[class*="rg_i"], img[class*="sFlh5c"], img[class*="WICvnb"], img[class*="isv-r"], img[data-src], img[src], img[data-iurl], img[data-lazy-src]')
            for idx, img in enumerate(img_tags):
                try:
                    thumb_src = img.get('src') or img.get('data-src') or img.get('data-iurl') or img.get('data-lazy-src') or ''
                    if thumb_src and not thumb_src.startswith('data:image') and any(thumb_src.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.webp']):
                        if any(re.search(pattern, thumb_src, re.IGNORECASE) for pattern in social_media_patterns):
                            logger_instance.debug(f"Skipping social media icon thumbnail in fallback: {thumb_src}")
                            continue
                        width = img.get('width', '0')
                        height = img.get('height', '0')
                        try:
                            if int(width) <= 50 or int(height) <= 50:
                                logger_instance.debug(f"Skipping small image (width: {width}, height: {height}): {thumb_src}")
                                continue
                        except ValueError:
                            pass
                        thumb_src = clean_source_url(thumb_src)
                        thumb_src = extract_true_url_from_wrapper(thumb_src, logger_instance)
                        thumb_src = clean_image_url(thumb_src)
                        if thumb_src:
                            main_image_urls.append(thumb_src)
                            main_descriptions.append(img.get('alt', '') or '')
                            main_source_urls.append('')
                            main_thumbs.append(thumb_src)
                except Exception as e:
                    logger_instance.error(f"Error processing img tag {idx}: {e}", exc_info=True)
                    continue

        min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
        main_image_urls = main_image_urls[:min_length]
        main_descriptions = main_descriptions[:min_length]
        main_source_urls = main_source_urls[:min_length]
        main_thumbs = main_thumbs[:min_length]
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