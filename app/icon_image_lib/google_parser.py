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
    """Extract image data from Google image search HTML."""
    logger = logger or logging.getLogger(__name__) # Ensure logger is available
    html_content = decode_html_bytes(html_bytes, logger)

    start_tag = 'FINANCE",[22,1]]]]]'
    end_tag = ':[null,null,null,"glbl'
    matched_google_image_data = LR().get(html_content, start_tag, end_tag)
    
    if 'Error' in matched_google_image_data or not matched_google_image_data:
        logger.warning('Main results tags not found or no data extracted (LR step)')
        return [], [], [], []
    
    # This part handles \uXXXX escapes if they are Python string literals
    thumbnails_data_str = str(matched_google_image_data).replace('\u003d', '=').replace('\u0026', '&')
    if '"2003":' not in thumbnails_data_str:
        logger.warning('No "2003" tag found in main thumbnails data')
        return [], [], [], []
    
    # Thumbnails (usually direct, like encrypted-tbn0.gstatic.com)
    raw_thumbs = re.findall(r'\[\"(https:\/\/encrypted-tbn0\.gstatic\.com\/images\?.*?)\",\d+,\d+\]', thumbnails_data_str)
    main_thumbs = [clean_source_url(url) for url in raw_thumbs] # Apply uXXXX decoding

    main_descriptions = re.findall(r'"2003":\[null,"[^"]*","[^"]*","(.*?)"', thumbnails_data_str)
    
    # Source URLs (page where image is found)
    raw_source_urls = re.findall(r'"2003":\[null,"[^"]*","(.*?)"', thumbnails_data_str)
    main_source_urls = []
    for url in raw_source_urls:
        s1 = clean_source_url(url) # Handles uXXXX literals
        s2 = extract_true_url_from_wrapper(s1, logger) # Handles wrappers & %-decode
        main_source_urls.append(s2)
        
    # Main Image URLs (should be direct image links)
    # Regex to find image URLs; removed_thumbs is thumbnails_data_str after removing thumbnail patterns
    removed_thumbs_data = re.sub(r'\[\"(https:\/\/encrypted-tbn0\.gstatic\.com\/images\?.*?)\",\d+,\d+\]', "", thumbnails_data_str)
    raw_main_image_urls = re.findall(r"(?:|,),\[\"(https:|http.*?)\",\d+,\d+\]", removed_thumbs_data)
    main_image_urls = []
    for url in raw_main_image_urls:
        img1 = clean_source_url(url) # Handles uXXXX literals
        img2 = extract_true_url_from_wrapper(img1, logger) # Handles wrappers & %-decode
        img3 = clean_image_url(img2) # Strips query params from actual image URL
        main_image_urls.append(img3)
    
    if not main_image_urls and main_thumbs: # Fallback if direct images not found
        logger.info("No main_image_urls found, using main_thumbs as fallback for image URLs.")
        # Thumbs are already processed by clean_source_url. They are less likely to be wrappers.
        # If they need to be treated as full image URLs, they might need clean_image_url too.
        main_image_urls = [clean_image_url(thumb_url) for thumb_url in main_thumbs]


    min_length = min(len(main_image_urls), len(main_descriptions), len(main_source_urls), len(main_thumbs))
    # Truncate lists to the minimum common length to ensure DataFrame integrity
    main_image_urls = main_image_urls[:min_length][:5]
    main_descriptions = main_descriptions[:min_length][:5]
    main_source_urls = main_source_urls[:min_length][:5]
    main_thumbs = main_thumbs[:min_length][:5]

    logger.debug(f"GetOriginal: URLs={len(main_image_urls)}, Desc={len(main_descriptions)}, Sources={len(main_source_urls)}, Thumbs={len(main_thumbs)}")
    return main_image_urls, main_descriptions, main_source_urls, main_thumbs

import logging
from bs4 import BeautifulSoup
from urllib.parse import unquote, urlparse
import re

def clean_source_url(raw_url: str) -> str:
    """Clean a URL by decoding Unicode escape sequences and removing unnecessary characters."""
    if not raw_url:
        return ""
    cleaned = unquote(raw_url)
    cleaned = re.sub(r"\?.*$|#$", "", cleaned)
    return cleaned.strip()

def extract_true_url_from_wrapper(raw_url: str, logger: logging.Logger) -> str:
    """Extract the true URL from Google's redirect wrapper URLs."""
    if not raw_url:
        return "No URL"
    if raw_url.startswith("/url?"):
        match = re.search(r"q=([^&]+)", raw_url)
        if match:
            return clean_source_url(match.group(1))
    return clean_source_url(raw_url)

def clean_image_url(url: str) -> str:
    """Clean an image URL to ensure it's valid and direct."""
    if not url or "data:image" in url:
        return "No image URL"
    if url.startswith("//"):
        url = f"https:{url}"
    elif url.startswith("/"):
        url = f"https://www.google.com{url}"
    return clean_source_url(url)

def get_results_page_results(html_bytes: bytes, logger: logging.Logger = None) -> tuple[list, list, list, list]:
    """
    Extract organic search results from Google search results HTML.
    
    Args:
        html_bytes: Raw HTML bytes of the Google search results page.
        logger: Optional logger instance for debugging.
    
    Returns:
        Tuple of (urls, titles, sources, thumbnails) as lists.
    """
    logger = logger or logging.getLogger(__name__)
    try:
        html_content = html_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        logger.error(f"Failed to decode HTML bytes: {e}")
        return [], [], [], []

    soup = BeautifulSoup(html_content, "html.parser")
    final_urls, final_titles, final_sources, final_thumbs = [], [], [], []

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
        source = source_tag.get_text(strip=True).split(" › ")[0] if source_tag else urlparse(url).netloc or "No source"
        final_sources.append(clean_source_url(source))

        img_tag = item_container.find("img", class_=lambda c: c and any(cls in c for cls in ["rg_i", "n3VNCb", "YQ4gaf"]))
        thumb_url = img_tag.get("src") or img_tag.get("data-src") if img_tag else None
        thumb = clean_image_url(thumb_url) if thumb_url and "data:image" not in thumb_url else "No thumbnail"
        final_thumbs.append(thumb)

    logger.debug(f"Parsed {len(final_urls)} results: URLs={len(final_urls)}, Titles={len(final_titles)}, Sources={len(final_sources)}, Thumbnails={len(final_thumbs)}")
    return final_urls, final_titles, final_sources, final_thumbs


def process_search_result(image_html_bytes, entry_id: int, logger=None) -> pd.DataFrame:
    """Process search result HTML bytes and return a DataFrame with image data."""
    logger = logger or logging.getLogger(__name__) # Ensure logger
    # It seems image_html_bytes is used for both original and results_page_results
    # This implies it contains all necessary data.
    
    # First, try to get images from the script/data block
    final_urls, final_descriptions, final_sources, final_thumbs = get_original_images(image_html_bytes, logger)
    
    # Then, try to get images from structured HTML elements (if any, or as supplement)
    # Pass the *current* lists to be appended to.
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

import logging
from typing import List, Dict
from bs4 import BeautifulSoup

def process_search_page(html_bytes: bytes, entry_id: int, logger=None) -> List[Dict]:
    """Process Google search page HTML to extract search results.

    Args:
        html_bytes (bytes): The HTML content of the search page in bytes.
        entry_id (int): Identifier for the search entry.
        logger (logging.Logger, optional): Logger instance for debugging and errors.

    Returns:
        List[Dict]: List of dictionaries containing extracted search result details.
    """
    logger = logger or logging.getLogger(__name__)
    
    if not html_bytes:
        logger.debug(f"EntryID {entry_id}: No search page HTML to process.")
        return []
    
    try:
        logger.debug(f"EntryID {entry_id}: Processing search page HTML ({len(html_bytes)} bytes).")
        
        # Decode bytes to string, assuming UTF-8 encoding
        html_content = html_bytes.decode('utf-8', errors='ignore')
        
        # Parse HTML with BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Initialize results list
        results = []
        
        # Find search result containers (based on common Google search result structure)
        # Note: Google’s structure may change; this targets common classes as of 2025
        result_divs = soup.find_all('div', class_='tF2Cxc')  # Common class for main results
        
        for idx, result in enumerate(result_divs):
            try:
                # Extract title
                title_tag = result.find('h3')
                title = title_tag.get_text(strip=True) if title_tag else "N/A"
                
                # Extract URL
                link_tag = result.find('a')
                url = link_tag.get('href') if link_tag else "placeholder://no-url"
                
                # Extract snippet/description
                snippet_tag = result.find('div', class_='VwiC3b')  # Common class for snippets
                snippet = snippet_tag.get_text(strip=True) if snippet_tag else "N/A"
                
                # Extract thumbnail URL if available
                thumbnail_tag = result.find('img')
                thumbnail_url = thumbnail_tag.get('src') if thumbnail_tag else url
                
                # Structure the result
                result_data = {
                    "EntryID": entry_id,
                    "Title": title,
                    "Url": url,
                    "Description": snippet,
                    "ThumbnailUrl": thumbnail_url
                }
                
                results.append(result_data)
                
                logger.debug(f"EntryID {entry_id}: Extracted result {idx + 1} - Title: {title[:50]}...")
                
            except Exception as e:
                logger.warning(f"EntryID {entry_id}: Failed to process result {idx + 1}: {e}")
                continue
        
        logger.info(f"EntryID {entry_id}: Successfully extracted {len(results)} results.")
        return results
    
    except Exception as e:
        logger.error(f"EntryID {entry_id}: Failed to process search page HTML: {e}", exc_info=True)
        return []