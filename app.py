from flask import Flask, request, send_file, render_template_string, jsonify
import os, tempfile
from app import create_app

app = create_app()

if __name__ == "__main__":
    app.run(debug=True)