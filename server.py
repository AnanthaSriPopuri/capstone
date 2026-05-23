"""
server.py  —  Capstone Studio local dev server
================================================
Serves index.html, style.css, and app.js from the same folder.

Usage:
    pip install flask
    python server.py

Then open:  http://localhost:5000
"""

from flask import Flask, send_from_directory
import os

app = Flask(__name__, static_folder=".")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@app.route("/")
def index():
    """Serve the main HTML page."""
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/style.css")
def css():
    """Serve the stylesheet."""
    return send_from_directory(BASE_DIR, "style.css")


@app.route("/app.js")
def js():
    """Serve the JavaScript file."""
    return send_from_directory(BASE_DIR, "app.js")


if __name__ == "__main__":
    print("=" * 50)
    print("  Capstone Studio — Local Dev Server")
    print("  Open:  http://localhost:5000")
    print("  Stop:  Ctrl + C")
    print("=" * 50)
    app.run(debug=True, port=5000)
