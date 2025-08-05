import urllib.parse
import logging # Assuming logging is already configured

# Make sure this logger is accessible or passed appropriately
# For simplicity, defining a logger here, but ideally use the one passed down
module_logger = logging.getLogger(__name__) # Or use the logger passed as param

def extract_true_url_from_wrapper(url_string: str, logger=None) -> str:
    """
    Extracts the true target URL from known wrapper URL patterns
    (e.g., Google redirects, Next.js image optimizer).
    Assumes uXXXX sequences have already been handled by clean_source_url.
    The returned URL will be standard percent-decoded.
    """
    logger = logger or module_logger # Use passed logger or module's default
    if not url_string or not isinstance(url_string, str): # Basic check
        return url_string

    try:
        # Parse the URL. For relative URLs like "/url?q=...", netloc will be empty.
        parsed = urllib.parse.urlparse(url_string)
        # parse_qs automatically unquotes parameter values (handles %20 etc.)
        query_params = urllib.parse.parse_qs(parsed.query)

        # 1. Check for Next.js _next/image wrapper
        #    Example: https://thedropdate.com/_next/image?url=ACTUAL_URL&w=...&q=...
        if parsed.path.endswith("/_next/image") and 'url' in query_params:
            inner_url = query_params['url'][0]
            logger.debug(f"Extracted from _next/image wrapper: {inner_url} (from {url_string})")
            # The inner_url from parse_qs is already unquoted.
            # We can return it, or unquote it again just in case of double encoding (rare).
            return urllib.parse.unquote_plus(inner_url)

        # 2. Check for Google image redirect/wrapper
        #    Examples:
        #    https://www.google.com/url?sa=i&url=ENCODED_TARGET_URL&psig=...
        #    /url?q=ENCODED_TARGET_URL&sa=... (relative path)
        #    https://images.google.com/imgres?imgurl=ACTUAL_URL&imgrefurl=...
        is_google_domain = parsed.netloc.startswith(('www.google.', 'images.google.'))
        is_relative_google_path = parsed.netloc == '' and parsed.path in ('/url', '/imgres')

        if is_google_domain or is_relative_google_path:
            if parsed.path in ('/url', '/imgres') or is_google_domain: # Check path for domain cases too
                target_param_key = None
                if 'url' in query_params: # Common for image results or sa=i links
                    target_param_key = 'url'
                elif 'q' in query_params: # Common for /url?q=
                    target_param_key = 'q'
                elif 'imgurl' in query_params and parsed.path == '/imgres': # Specific to /imgres
                    target_param_key = 'imgurl'

                if target_param_key and query_params[target_param_key]:
                    inner_url = query_params[target_param_key][0]
                    logger.debug(f"Extracted from Google wrapper (param '{target_param_key}'): {inner_url} (from {url_string})")
                    return urllib.parse.unquote_plus(inner_url) # parse_qs already unquotes, this is for safety

        # Add other known wrapper patterns here if necessary (e.g., Bing, Yahoo)
        # Example Bing (hypothetical, check actual Bing patterns):
        # if parsed.netloc.endswith('bing.com') and parsed.path.startswith('/images/async') and 'url' in query_params:
        #    inner_url = query_params['url'][0]
        #    logger.debug(f"Extracted from Bing wrapper: {inner_url} (from {url_string})")
        #    return urllib.parse.unquote_plus(inner_url)

        # If not a known wrapper, return the original URL, but ensure it's standard URL-decoded.
        # This handles cases like "https://example.com/my%20image.jpg"
        final_url = urllib.parse.unquote_plus(url_string)
        if final_url != url_string and not (is_google_domain or is_relative_google_path): # Avoid logging already handled Google URLs as simple unquotes
             logger.debug(f"Unquoted non-wrapper URL: {final_url} (from {url_string})")
        return final_url

    except Exception as e:
        logger.error(f"🔴 Error in extract_true_url_from_wrapper for '{url_string}': {e}", exc_info=False) # exc_info=True for full traceback
        # Fallback: try to unquote the original string, or return as is if unquoting fails
        try:
            return urllib.parse.unquote_plus(url_string)
        except Exception:
            return url_string # Absolute fallback
import re
import chardet
from bs4 import BeautifulSoup
import logging
import pandas as pd
import urllib.parse # Added for the new function

# Assuming LR class is in icon_image_lib.LR; otherwise, include it directly here
from icon_image_lib.LR import LR # Make sure this import works

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# It's good practice to get a logger instance per module
# logger = logging.getLogger(__name__) # This would be used if not passing logger around

def clean_source_url(s): # Your existing function
    """Clean and decode URL string by replacing encoded characters."""
    if not isinstance(s, str): # robustness
        return s
    simplified_str = s.replace('\\\\', '') # Consider if this is always correct or if `codecs.decode(s, 'unicode_escape')` might be better for some inputs
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

def clean_image_url(url): # Your existing function
    """Extract the base image URL without query parameters if it matches image extensions."""
    if not isinstance(url, str): # robustness
        return url
    # This pattern specifically targets URLs ending with common image extensions
    # and strips query parameters from *them*.
    pattern = re.compile(r'(.*\.(?:png|jpg|jpeg|gif|webp|avif|svg))(?:\?.*)?', re.IGNORECASE) # Added more extensions
    match = pattern.match(url)
    return match.group(1) if match else url

def decode_html_bytes(html_bytes, logger): # Your existing function
    """Decode HTML bytes to string with fallback encoding."""
    try:
        detected_encoding = chardet.detect(html_bytes)['encoding'] or 'utf-8'
        logger.debug(f"Detected encoding: {detected_encoding}")
        return html_bytes.decode(detected_encoding, errors='replace')
    except Exception as e:
        logger.error(f"🔴 Decoding error: {e}", exc_info=True)
        return html_bytes.decode('utf-8', errors='replace')

# The extract_true_url_from_wrapper function defined above should be here
def get_original_images(html_bytes, logger=None):
    """Extract image data from Google image search HTML in grid order."""
    logger = logger or logging.getLogger(__name__)
    html_content = decode_html_bytes(html_bytes, logger)
    soup = BeautifulSoup(html_content, 'html.parser')
    
    main_image_urls = []
    main_descriptions = []
    main_source_urls = []
    main_thumbs = []
    
    # Select all metadata divs in order (matches grid sequence)
    meta_divs = soup.select('div.rg_meta.notranslate')
    logger.debug(f"Found {len(meta_divs)} metadata entries.")
    
    for div in meta_divs:
        try:
            # The div text is a JSON string; clean and load it
            meta_json_str = div.text.strip()
            # Fix any escaped unicode or quotes if present (rare, but safe)
            meta_json_str = re.sub(r'\\u([0-9a-fA-F]{4})', lambda m: chr(int(m.group(1), 16)), meta_json_str)
            meta_data = json.loads(meta_json_str)
            
            # Extract fields (keys may vary slightly by Google version; adjust if needed)
            full_url = meta_data.get('ou', '')  # Original/full image
            if full_url:
                full_url = clean_source_url(full_url)
                full_url = extract_true_url_from_wrapper(full_url, logger)
                full_url = clean_image_url(full_url)
                main_image_urls.append(full_url)
            
            desc = meta_data.get('pt', '') or meta_data.get('s', '')  # Title or snippet
            main_descriptions.append(desc)
            
            source_url = meta_data.get('ru', '')  # Referring page
            if source_url:
                source_url = clean_source_url(source_url)
                source_url = extract_true_url_from_wrapper(source_url, logger)
                main_source_urls.append(source_url)
            
            thumb_url = meta_data.get('tu', '') or meta_data.get('st', '')  # Thumbnail (may be base64 or URL)
            if thumb_url and thumb_url.startswith('http'):  # Ignore base64 for now
                thumb_url = clean_source_url(thumb_url)
                thumb_url = extract_true_url_from_wrapper(thumb_url, logger)
                main_thumbs.append(thumb_url)
            else:
                main_thumbs.append('')  # Placeholder if missing
            
        except json.JSONDecodeError as e:
            logger.warning(f"Failed to parse JSON in metadata div: {e}")
            continue  # Skip malformed entries to preserve order
    
    # Fallback: If thumbnails are missing, extract from img tags (in order)
    if not main_thumbs or len(main_thumbs) < len(main_image_urls):
        logger.info("Thumbnails incomplete; falling back to img tags.")
        img_tags = soup.select('img.rg_i.Q4LuWd')  # Thumbnail imgs in order
        main_thumbs = []
        for img in img_tags:
            thumb_src = img.get('src') or img.get('data-src', '')
            if thumb_src:
                thumb_src = clean_source_url(thumb_src)
                thumb_src = extract_true_url_from_wrapper(thumb_src, logger)
                main_thumbs.append(thumb_src)
    
    # Align lists (should be equal now, but truncate if any mismatches)
    min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
    main_image_urls = main_image_urls[:min_length][:5]
    main_descriptions = main_descriptions[:min_length][:5]
    main_source_urls = main_source_urls[:min_length][:5]
    main_thumbs = main_thumbs[:min_length][:5]
    
    logger.debug(f"Extracted (in order): URLs={len(main_image_urls)}, Desc={len(main_descriptions)}, Sources={len(main_source_urls)}, Thumbs={len(main_thumbs)}")
    return main_image_urls, main_descriptions, main_source_urls, main_thumbs

def process_search_result(image_html_bytes, entry_id: int, logger=None) -> pd.DataFrame:
    """Process search result HTML bytes and return a DataFrame with image data."""
    logger = logger or logging.getLogger(__name__) # Ensure logger
    # It seems image_html_bytes is used for both original and results_page_results
    # This implies it contains all necessary data.
    
    # First, try to get images from the script/data block
    final_urls, final_descriptions, final_sources, final_thumbs = get_original_images(image_html_bytes, logger)
    
    # Ensure all lists have the same length before creating DataFrame
    all_lists = [final_urls, final_descriptions, final_sources, final_thumbs]
    if not all_lists: # Should not happen if functions return empty lists
        logger.warning(f"EntryID {entry_id}: No data extracted. Returning empty DataFrame.")
        return pd.DataFrame()
        
    min_len = min(len(lst) for lst in all_lists)
    
    # Check if any list was longer and got truncated, indicating a potential mismatch
    if any(len(lst) > min_len for lst in all_lists):
        logger.warning(
            f"EntryID {entry_id}: Data lists had different lengths "
            f"({[len(lst) for lst in all_lists]}). Truncated to min length: {min_len}."
        )

    df = pd.DataFrame({
        'EntryID': [entry_id] * min_len,
        'ImageUrl': final_urls[:min_len],
        'ImageDesc': final_descriptions[:min_len],
        'ImageSource': final_sources[:min_len],
        'ImageUrlThumbnail': final_thumbs[:min_len]
    })
    
    logger.info(f"Processed EntryID {entry_id} with {len(df)} images (after potential truncation).")
    return df
