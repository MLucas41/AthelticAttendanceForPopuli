import requests
import os
import time
import json
from dotenv import load_dotenv
from functools import lru_cache
from pathlib import Path

load_dotenv()  # Load the API key from the .env file into environment variables

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

url = "https://corban.populiweb.com/api2/"  # Corban's Populi API base URL
sandbox_key = os.getenv("SANDBOX_TOKEN")  