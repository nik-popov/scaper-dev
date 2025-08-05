import requests
import logging
import json
from sqlalchemy.ext.asyncio import create_async_engine
from sqlalchemy import create_engine

# Setup logging
logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Database connection string
conn_str = "DRIVER={ODBC Driver 17 for SQL Server};Server=35.172.243.170;Database=luxurymarket_p4;Uid=luxurysitescraper;Pwd=Ftu5675FDG54hjhiuu$;"

# Create SQLAlchemy engines
async_engine = create_async_engine("mssql+aioodbc:///?odbc_connect=" + conn_str)
engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={conn_str}",
    echo=False,
    pool_size=5,
    max_overflow=5,
    pool_timeout=60,
    pool_pre_ping=True,
    pool_recycle=3600,
    isolation_level="READ COMMITTED"
)

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

# Define variables from app_config or fallback to defaults
if app_config:
    VERSION = app_config.get('version', '3.8.6')
    SENDER_EMAIL = app_config.get('email_settings', {}).get('sender_email', 'nik@luxurymarket.com')
    SENDER_PASSWORD = app_config.get('email_settings', {}).get('sender_password', 'wvug kynd dfhd xrjh')
    SENDER_NAME = app_config.get('email_settings', {}).get('sender_name', 'superscraper')
    S3_CONFIG = app_config.get('s3_config', {
        "endpoint": "https://s3.us-east-2.amazonaws.com",
        "region": "us-east-2",
        "access_key": "AKIA2CUNLEV6V627SWI7",
        "secret_key": "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",
        "bucket_name": "iconluxurygroup",
        "r2_endpoint": "https://97d91ece470eb7b9aa71ca0c781cfacc.r2.cloudflarestorage.com",
        "r2_access_key": "5547ff7ffb8f3b16a15d6f38322cd8bd",
        "r2_secret_key": "771014b01093eceb212dfea5eec0673842ca4a39456575ca7ff43f768cf42978",
        "r2_account_id": "97d91ece470eb7b9aa71ca0c781cfacc",
        "r2_bucket_name": "iconluxurygroup",
        "r2_custom_domain": "https://iconluxury.shop",
    })
else:
    logger.error("Failed to load configuration. Using fallback values.")
