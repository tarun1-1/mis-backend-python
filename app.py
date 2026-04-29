from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
import os
import datetime

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
    last_modified = datetime.datetime.fromtimestamp(os.path.getmtime(UPLOAD_FILE))
    response = make_response(send_file(
        UPLOAD_FILE,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ))
    response.headers["Cache-Control"] = "no-store"
    response.headers["Last-Updated"] = last_modified.strftime("%d %b %Y, %I:%M %p")
    return response

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
