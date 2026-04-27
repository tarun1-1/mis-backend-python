from flask import Flask, request, send_file
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)

UPLOAD_FILE = "data.csv"

@app.route("/")
def home():
    return "Backend is running"

# Upload API
@app.route("/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    file.save(UPLOAD_FILE)

    return "File uploaded successfully"

# Get latest data
@app.route("/data", methods=["GET"])
def get_data():
    if not os.path.exists(UPLOAD_FILE):
        return "No data available", 404

    return send_file(UPLOAD_FILE)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)