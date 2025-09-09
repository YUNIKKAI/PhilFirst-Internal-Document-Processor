from flask import Flask, render_template
from config import DevelopmentConfig, ProductionConfig
from dotenv import load_dotenv
#from renewal.routes import renewal_bp
from soa_direct.routes import soa_bp
import os

def create_app():
    # Load environment variables from .env
    load_dotenv()

    # Initialize Flask app
    app = Flask(__name__, static_folder="static")

    # Choose config based on FLASK_ENV
    if os.getenv("FLASK_ENV") == "production":
        app.config.from_object(ProductionConfig)
    else:
        app.config.from_object(DevelopmentConfig)

    # Register Blueprints
    #app.register_blueprint(renewal_bp, url_prefix="/renewal")
    app.register_blueprint(soa_bp, url_prefix="/soa_direct")

    # Root route
    @app.route("/")
    def home():
        return render_template("home.html")

    return app

# Entry point for direct execution (dev only — Gunicorn handles prod)
if __name__ == "__main__":
    app = create_app()
    app.run(
        host=os.getenv("HOST", "127.0.0.1"),
        port=int(os.getenv("PORT", "8000")),
        debug=os.getenv("FLASK_ENV") != "production"
    )
