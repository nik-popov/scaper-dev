import logging
from load_config import load_config_constants  # Adjust import based on your file structure

# Setup logging
logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Load configuration
CONFIG_URL = "https://iconluxury.shop/secrets/secrets_scraper_config.json"
config = load_config_constants(CONFIG_URL, fallback_db_password="Ftu5675FDG54hjhiuu$")
if not config:
    logger.error("Failed to load configuration. Exiting.")
    raise SystemExit(1)

# AES Decryption class
import base64
import hashlib
from Crypto.Cipher import AES

class AESCipher(object):

    def __init__(self, key):
        self.bs = AES.block_size
        self.key = hashlib.sha256(key.encode()).digest()

    def decrypt(self, enc):
        enc = base64.b64decode(enc)
        iv = enc[:AES.block_size]
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        return AESCipher._unpad(cipher.decrypt(enc[AES.block_size:])).decode('utf-8')

    @staticmethod
    def _unpad(s):
        return s[:-ord(s[len(s)-1:])]
    
ENCRYPTION_KEY = "f5319e962b4149fdaac19ecb37ba0d5d"  # Replace with actual key
cipher = AESCipher(ENCRYPTION_KEY)

# Function to safely decrypt config values
def decrypt_if_encrypted(value, is_encrypted=False):
    """
    Decrypt a value if it's encrypted, otherwise return as-is
    """
    if is_encrypted and isinstance(value, str):
        try:
            return cipher.decrypt(value)
        except Exception as e:
            logger.warning(f"Failed to decrypt value: {e}")
            return value
    return value

# Assign config values to variables (with optional decryption)
# IMPORTANT: Only set is_encrypted=True for values that are actually encrypted in your config
# Setting is_encrypted=True for unencrypted values will cause errors

VERSION = config['version']
SENDER_EMAIL = config['email_settings']['sender_email']
SENDER_PASSWORD = decrypt_if_encrypted(config['email_settings']['sender_password'], is_encrypted=False)  # Set to True only if encrypted
SENDER_NAME = config['email_settings']['sender_name']
GOOGLE_API_KEY = decrypt_if_encrypted(config['google_api_settings']['api_key'], is_encrypted=False)  # Set to True only if encrypted
GROK_API_KEY = decrypt_if_encrypted(config['grok_settings']['api_key'], is_encrypted=False)  # Set to True only if encrypted
GROK_ENDPOINT = config['grok_settings']['endpoint']
DATAPROXY_API_KEY = decrypt_if_encrypted(config['dataproxy_settings']['api_key'], is_encrypted=False)  # Set to True only if encrypted
SEARCH_PROXY_API_URL = config['dataproxy_settings']['api_url']
ROAMINGPROXY_API_KEY = decrypt_if_encrypted(config['roamingproxy_settings']['api_key'], is_encrypted=False)  # Set to True only if encrypted
ROAMINGPROXY_API_URL = config['roamingproxy_settings']['api_url']
PROXY_STRATEGY = config.get('proxy_strategy', 'round_robin')
BRAND_RULES_URL = config['brand_settings']['brand_rules_url']
AWS_ACCESS_KEY_ID = decrypt_if_encrypted(config['aws_settings']['access_key_id'], is_encrypted=False)  # Set to True only if encrypted
AWS_SECRET_ACCESS_KEY = decrypt_if_encrypted(config['aws_settings']['secret_access_key'], is_encrypted=False)  # Set to True only if encrypted
REGION = config['aws_settings']['region']
S3_CONFIG = config['s3_config']
DB_PASSWORD = decrypt_if_encrypted(config['database_settings']['db_password'], is_encrypted=False)  # Set to True only if encrypted
BASE_CONFIG_URL = config['base_config']['base_config_url']

# RabbitMQ configuration
RABBITMQ_URL = decrypt_if_encrypted(config['rabbitmq_settings']['url'], is_encrypted=False)  # Set to True only if encrypted
RABBITMQ_USER = config['rabbitmq_settings']['user']
RABBITMQ_PASSWORD = decrypt_if_encrypted(config['rabbitmq_settings']['password'], is_encrypted=False)  # Set to True only if encrypted
RABBITMQ_HOST = config['rabbitmq_settings']['host']
RABBITMQ_PORT = config['rabbitmq_settings']['port']
RABBITMQ_VHOST = config['rabbitmq_settings']['vhost']
# Example: Log loaded config for debugging
logger.info(f"Loaded config: VERSION={VERSION}, BASE_CONFIG_URL={BASE_CONFIG_URL}, RABBITMQ_URL={RABBITMQ_URL}")