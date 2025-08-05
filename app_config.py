import requests
import logging
import json

# Setup logging
logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

def load_config_constants(config_url, fallback_db_password=None):
    try:
        response = requests.get(config_url, timeout=10)
        response.raise_for_status()
        config = response.json()
        if 'database_settings' in config and 'db_password' not in config['database_settings']:
            config['database_settings']['db_password'] = fallback_db_password
            logger.warning("Database password not found in config. Using fallback password.")
        logger.info(f"Successfully loaded configuration from {config_url}")
        return config
    except requests.exceptions.RequestException as e:
        logger.error(f"Failed to fetch configuration from {config_url}: {str(e)}")
        return None
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse JSON configuration from {config_url}: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error while loading configuration from {config_url}: {str(e)}")
        return None

# Fetch configuration
CONFIG_URL = "https://iconluxury.shop/secrets/secrets_scraper_config.json"
app_config = load_config_constants(CONFIG_URL, fallback_db_password="Ftu5675FDG54hjhiuu$")