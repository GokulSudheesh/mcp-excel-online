import os
import sys
from dotenv import load_dotenv
import logging


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
    stream=sys.stdout,
)

# Load environment variables from .env file
load_dotenv()


class Settings():
    CLIENT_ID = os.getenv('APP_CLIENT_ID')
    TENANT_ID = os.getenv('APP_TENANT_ID')
    CLIENT_SECRET = os.getenv('APP_CLIENT_SECRET')
    DRIVE_ID = os.getenv('DRIVE_ID')
