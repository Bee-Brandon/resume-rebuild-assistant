"""Central configuration — tune settings here, not in module code."""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env from project root (keeps secrets out of code)
load_dotenv(Path(__file__).parent / ".env")

# Paths
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"
PROMPTS_DIR = BASE_DIR / "prompts"
DB_PATH = BASE_DIR / "resume_rebuild.db"
DEFAULT_TEMPLATE = TEMPLATES_DIR / "default.docx"

# Ensure runtime directories exist
OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Claude API
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL_NAME = "claude-sonnet-4-6"  # supports vision; swap to claude-opus-4-6 for max accuracy
MAX_TOKENS = 4096

# Ingestion
SCAN_DPI = 300  # DPI for converting scanned PDF pages to images

# Logging
API_LOG_FILE = LOGS_DIR / "api_calls.log"
