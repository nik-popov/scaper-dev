import urllib.parse
import requests
import logging
import pandas as pd
import json

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
    """
    Extracts the true target URL from known wrapper URL patterns.
    """
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

def process_api_image_results(json_data, entry_id: int, logger=None) -> pd.DataFrame:
    """
    Process JSON from Google Custom Search API into a DataFrame.

    Args:
        json_data (dict or str): JSON data from the API or JSON string.
        entry_id (int): Identifier for the search result entry.
        logger (logging.Logger, optional): Logger for debugging.

    Returns:
        pd.DataFrame: DataFrame with columns EntryID, ImageUrl, ImageDesc, ImageSource, ImageUrlThumbnail.
                      Returns empty DataFrame on failure.
    """
    logger = logger or logging.getLogger(__name__)
    if not logger.handlers:
        logger.addHandler(logging.StreamHandler())
        logger.setLevel(logging.INFO)

    # Validate inputs
    if not isinstance(entry_id, int):
        logger.error(f"Invalid entry_id type: {type(entry_id)}. Expected int.")
        return pd.DataFrame()

    try:
        # Parse JSON if string
        if isinstance(json_data, str):
            json_data = json.loads(json_data)

        if not isinstance(json_data, (list, dict)):
            logger.error(f"EntryID {entry_id}: Invalid JSON data type: {type(json_data)}")
            return pd.DataFrame()

        # If JSON is a list (direct results from Worker), use it
        items = json_data if isinstance(json_data, list) else json_data.get("items", [])

        if not items:
            logger.warning(f"EntryID {entry_id}: No items found in JSON data.")
            return pd.DataFrame()

        # Process items
        final_urls = []
        final_descriptions = []
        final_sources = []
        final_thumbs = []

        for item in items[:5]:  # Limit to 5 results, matching your Worker's logic
            image_url = clean_image_url(extract_true_url_from_wrapper(clean_source_url(item.get("ImageUrl", "No image URL")), logger))
            description = item.get("ImageDesc", "No description")
            source = extract_true_url_from_wrapper(clean_source_url(item.get("ImageSource", "No source")), logger)
            thumb = clean_source_url(item.get("ImageUrlThumbnail", "No thumbnail URL"))

            final_urls.append(image_url)
            final_descriptions.append(description)
            final_sources.append(source)
            final_thumbs.append(thumb)

        # Create DataFrame
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
        logger.error(f"EntryID {entry_id}: Unexpected error: {str(e)}")
        return pd.DataFrame()

def fetch_and_process_images(query: str, entry_id: int, search_type: str = "image", worker_url: str = "https://browser-worker.nik-97d.workers.dev") -> pd.DataFrame:
    """
    Fetch image search results from Cloudflare Worker and process into a DataFrame.

    Args:
        query (str): Search query (e.g., "TH03279J078").
        entry_id (int): Identifier for the search result entry.
        search_type (str): Search type ("image" or "web").
        worker_url (str): Worker endpoint URL.

    Returns:
        pd.DataFrame: DataFrame with image results.
    """
    logger = logging.getLogger(__name__)
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
        return process_api_image_results(json_data, entry_id, logger)

    except Exception as e:
        logger.error(f"Error fetching/processing query '{query}': {str(e)}")
        return pd.DataFrame()