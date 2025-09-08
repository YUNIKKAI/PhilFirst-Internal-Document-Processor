import os
import tempfile
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class Config:
    # Security
    SECRET_KEY = os.getenv("SECRET_KEY", "fallback-key")

    # Flask / File Handling
    MAX_CONTENT_LENGTH = int(os.getenv("MAX_FILE_SIZE", 16 * 1024 * 1024))  # per file limit
    UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", tempfile.gettempdir())
    SEND_FILE_MAX_AGE_DEFAULT = int(os.getenv("SEND_FILE_MAX_AGE_DEFAULT", 31536000))

    # Custom Upload Limits
    MAX_FILE_SIZE = int(os.getenv("MAX_FILE_SIZE", 16 * 1024 * 1024))        # 16MB
    MAX_TOTAL_SIZE = int(os.getenv("MAX_TOTAL_SIZE", 64 * 1024 * 1024))      # 64MB
    MAX_FILES = int(os.getenv("MAX_FILES", 50))

    # Logging
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
    LOG_MAX_BYTES = int(os.getenv("LOG_MAX_BYTES", 10 * 1024 * 1024))        # 10MB
    LOG_BACKUP_COUNT = int(os.getenv("LOG_BACKUP_COUNT", 10))

    # Server (used by Gunicorn)
    HOST = os.getenv("HOST", "0.0.0.0")
    PORT = int(os.getenv("PORT", "8000"))

class DevelopmentConfig(Config):
    DEBUG = True
    FLASK_ENV = "development"

class ProductionConfig(Config):
    DEBUG = False
    FLASK_ENV = "production"
