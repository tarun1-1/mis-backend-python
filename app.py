from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)

UPLOAD_FILE = "data.xlsx"

@app.route("/")
def home():
    return "CarmaOne MIS Backend is running"

@app.route("/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files['file']
    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "Only .xlsx files allowed"}), 400
    file.save(UPLOAD_FILE)
    return jsonify({"success": True, "message": "File uploaded successfully"})

@app.route("/data", methods=["GET"])
def get_data():
    if not os.path.exists(UPLOAD_FILE):
        return jsonify({"error": "No data available"}), 404
    return send_file(UPLOAD_FILE, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)