"""
CDM Analysis Web Service
=========================
A small always-on web service that:
  1. Gets called by Zapier whenever a new file lands in your Drive "input" folder
  2. Downloads that file from Google Drive
  3. Runs the full CDM analysis (cdm_analysis.py)
  4. Uploads the charts + Excel report back to your Drive "output" folder
  5. Returns a JSON summary that Zapier can use to send a Slack/email notification

SETUP REQUIRED (see README.md for full step-by-step):
  - A Google Cloud "service account" with access to your Drive folders
  - Its credentials pasted into the GOOGLE_SERVICE_ACCOUNT_JSON environment variable
  - The IDs of your Drive "input" and "output" folders in INPUT_FOLDER_ID / OUTPUT_FOLDER_ID
"""

import os
import io
import json
import tempfile
import traceback

from flask import Flask, jsonify

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

from cdm_analysis import load_data, analyze

app = Flask(__name__)

SCOPES = ["https://www.googleapis.com/auth/drive"]
INPUT_FOLDER_ID = os.environ.get("INPUT_FOLDER_ID", "")
OUTPUT_FOLDER_ID = os.environ.get("OUTPUT_FOLDER_ID", "")


def get_drive_service():
    creds_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    creds_info = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def get_latest_file(service):
    """Find the most recently modified xlsx/csv file in the input folder."""
    query = (
        f"'{INPUT_FOLDER_ID}' in parents and trashed = false and "
        "(mimeType='text/csv' or "
        "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')"
    )
    results = service.files().list(
        q=query, orderBy="modifiedTime desc", pageSize=1,
        fields="files(id, name, modifiedTime)"
    ).execute()
    files = results.get("files", [])
    if not files:
        return None
    return files[0]


def download_file(service, file_id, dest_path):
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def upload_folder_contents(service, local_folder, parent_folder_id, run_name):
    """Create a run_<timestamp> subfolder in Drive output folder and upload all files into it."""
    folder_metadata = {
        "name": run_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }
    folder = service.files().create(body=folder_metadata, fields="id").execute()
    run_folder_id = folder["id"]

    uploaded = []
    for fname in sorted(os.listdir(local_folder)):
        fpath = os.path.join(local_folder, fname)
        if not os.path.isfile(fpath):
            continue
        media = MediaFileUpload(fpath, resumable=True)
        file_metadata = {"name": fname, "parents": [run_folder_id]}
        f = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
        uploaded.append(fname)

    # Make the run folder viewable via link (optional convenience)
    try:
        service.permissions().create(
            fileId=run_folder_id, body={"role": "reader", "type": "anyone"}
        ).execute()
    except Exception:
        pass

    folder_link = f"https://drive.google.com/drive/folders/{run_folder_id}"
    return folder_link, uploaded


@app.route("/run-analysis", methods=["POST", "GET"])
def run_analysis():
    try:
        if not INPUT_FOLDER_ID or not OUTPUT_FOLDER_ID:
            return jsonify({"status": "error", "message": "INPUT_FOLDER_ID / OUTPUT_FOLDER_ID not configured"}), 500

        service = get_drive_service()
        latest = get_latest_file(service)
        if latest is None:
            return jsonify({"status": "error", "message": "No files found in input folder"}), 404

        with tempfile.TemporaryDirectory() as tmpdir:
            local_input_path = os.path.join(tmpdir, latest["name"])
            download_file(service, latest["id"], local_input_path)

            data = load_data(local_input_path)
            outdir = os.path.join(tmpdir, "output")
            summary_text = analyze(data, outdir)

            import pandas as pd
            run_name = f"run_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}"
            folder_link, uploaded_files = upload_folder_contents(service, outdir, OUTPUT_FOLDER_ID, run_name)

        return jsonify({
            "status": "success",
            "source_file": latest["name"],
            "output_folder_link": folder_link,
            "files_generated": uploaded_files,
            "summary": summary_text,
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e),
            "trace": traceback.format_exc(),
        }), 500


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "CDM analysis service is running"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
