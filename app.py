from flask import Flask, request, jsonify, session, redirect, url_for
from flask_cors import CORS
import requests
import os
import json
from urllib.parse import urlencode
import torch
import torchaudio

print(torch.__version__)
print(torchaudio.__version__)

app = Flask(__name__)
CORS(app)
app.secret_key = os.urandom(24)  # Secret key for session management

# Microsoft Graph API Credentials
CLIENT_ID = "fa0f6728-1965-42f8-9895-bb8451c4de33"
CLIENT_SECRET = "fe90d0a0-3360-4fe5-83d6-7a110ce0e81b"
TENANT_ID = "c0072956-c395-48bb-9809-bfdea0eb26c8"
REDIRECT_URI = "https://4394-49-150-251-206.ngrok-free.app/callback"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite", "User.Read"]

# Token storage (for testing only, use a database in production)
TOKEN_CACHE = {}


# ✅ User Login (Redirect to Microsoft Authentication)
@app.route("/login")
def login():
    auth_url = f"{AUTHORITY}/oauth2/v2.0/authorize?{urlencode({
        'client_id': CLIENT_ID,
        'response_type': 'code',
        'redirect_uri': REDIRECT_URI,
        'response_mode': 'query',
        'scope': ' '.join(SCOPES),
    })}"
    return redirect(auth_url)


# ✅ OAuth Callback (Get Access Token)
@app.route("/callback")
def callback():
    code = request.args.get("code")
    if not code:
        return jsonify({"error": "Authorization failed"}), 400

    token_url = f"{AUTHORITY}/oauth2/v2.0/token"
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "scope": ' '.join(SCOPES)
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    response = requests.post(token_url, data=token_data, headers=headers)

    if response.status_code == 200:
        token_json = response.json()
        session["token"] = token_json["access_token"]
        print("Access Token:", token_json["access_token"])
        return jsonify({"message": "Login successful", "token": token_json["access_token"]})
    else:
        return jsonify({"error": "Failed to obtain access token", "details": response.text}), 400


# ✅ Fetch Excel File from OneDrive
@app.route("/fetch_excel", methods=["POST"])
def fetch_excel():
    if "token" not in session:
        return jsonify({"error": "User not authenticated"}), 401

    file_id = request.json.get("file_id")
    if not file_id:
        return jsonify({"error": "Missing file ID"}), 400

    graph_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets"
    headers = {"Authorization": f"Bearer {session['token']}"}
    response = requests.get(graph_url, headers=headers)

    if response.status_code != 200:
        return jsonify({"error": "Failed to fetch Excel file", "details": response.json()}), 400

    return jsonify(response.json())


# ✅ Logout Route (Clears Session & Redirects to Microsoft Logout)
@app.route("/logout")
def logout():
    session.clear()  # Clear user session
    logout_url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=http://localhost:5000"
    return redirect(logout_url)


if __name__ == "__main__":
    app.run(debug=True)
