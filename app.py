from flask import Flask, render_template
from config import Config
from dotenv import load_dotenv
from renewal.routes import renewal_bp
# from soa_direct.routes import soa_bp  # Uncomment when ready

def create_app():
    # Load environment variables
    load_dotenv()

    # Initialize Flask app
    app = Flask(__name__, static_folder="static")
    app.config.from_object(Config)

    # Register Blueprints
    app.register_blueprint(renewal_bp, url_prefix="/renewal")
    # app.register_blueprint(soa_bp, url_prefix="/soa")  # Future integration

    # Root route
    @app.route("/")
    def home():
        return render_template("home.html")

    return app

# Entry point for direct execution
if __name__ == "__main__":
    app = create_app()
    app.run(debug=True)