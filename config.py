import os
from dotenv import load_dotenv
import tempfile

load_dotenv()
class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "fallback-key")
    MAX_CONTENT_LENGTH = int(os.getenv("MAX_CONTENT_LENGTH", 16 * 1024 * 1024))
    UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", tempfile.gettempdir())
    SEND_FILE_MAX_AGE_DEFAULT = int(os.getenv("SEND_FILE_MAX_AGE_DEFAULT", 31536000))

class DevelopmentConfig(Config):
    DEBUG = True

class ProductionConfig(Config):
    DEBUG = False