import json
import time
import requests
from pathlib import Path

UPLOAD_TIMEOUT = 300  # 5 minutes
POLL_INTERVAL = 10  # 10 seconds
MAX_RETRIES = 5

# Load credentials
with open("X:/adobe/pdfservices-api-credentials.json", "r") as f:
    credentials = json.load(f)

token_cache = None

def get_access_token():
    global token_cache
    if token_cache and token_cache["expires_at"] > time.time():
        return token_cache["access_token"]

    url = "https://ims-na1.adobelogin.com/ims/token/v3"
    data = {
        "grant_type": "client_credentials",
        "client_id": credentials["client_credentials"]["client_id"],
        "client_secret": credentials["client_credentials"]["client_secret"],
        "scope": "openid,AdobeID,read_organizations,exportpdf"
    }
    
    response = requests.post(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=30)
    response.raise_for_status()
    token_data = response.json()
    
    token_cache = {
        "access_token": token_data["access_token"],
        "expires_at": time.time() + token_data["expires_in"] - 30
    }
    return token_cache["access_token"]

def upload_pdf(access_token, file_path):
    url = "https://pdf-services.adobe.io/assets"
    headers = {
        "x-api-key": credentials["client_credentials"]["client_id"],
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, json={"mediaType": "application/pdf"}, headers=headers, timeout=30)
    response.raise_for_status()
    upload_data = response.json()
    
    with open(file_path, "rb") as f:
        requests.put(upload_data["uploadUri"], data=f, headers={"Content-Type": "application/pdf"}, timeout=UPLOAD_TIMEOUT)
    
    return upload_data["assetID"]

def convert_pdf_to_docx(access_token, asset_id):
    url = "https://pdf-services.adobe.io/operation/exportpdf"
    headers = {
        "x-api-key": credentials["client_credentials"]["client_id"],
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, json={"assetID": asset_id, "targetFormat": "docx", "ocrLang": "en-US"}, headers=headers, timeout=30)
    response.raise_for_status()
    
    job_id = response.headers.get("Location", "").split("/")[-2]
    if not job_id:
        raise ValueError("Job ID not found in response")
    return job_id

def poll_and_download_result(access_token, job_id, output_path):
    url = f"https://pdf-services.adobe.io/operation/exportpdf/{job_id}/status"
    headers = {
        "x-api-key": credentials["client_credentials"]["client_id"],
        "Authorization": f"Bearer {access_token}"
    }
    retries = 0
    download_uri = None
    
    while retries < MAX_RETRIES:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        status_data = response.json()
        print(f"Job status: {status_data.get('status')}")
        
        if status_data.get("status") == "done":
            download_uri = status_data.get("downloadUri") or status_data.get("asset", {}).get("downloadUri")
            if not download_uri:
                raise ValueError("Download URI not found")
            break
        
        time.sleep(POLL_INTERVAL)
        retries += 1
    
    if not download_uri:
        raise ValueError("Conversion completed but download URI is missing")
    
    response = requests.get(download_uri, stream=True, timeout=UPLOAD_TIMEOUT)
    response.raise_for_status()
    with open(output_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
    
    print(f"File successfully saved to: {output_path}")

def main():
    try:
        file_path = "X:/adobe/sample3.pdf"
        output_path = str(Path(file_path).with_suffix(".docx"))
        
        print(f"Starting conversion for: {file_path}")
        access_token = get_access_token()
        asset_id = upload_pdf(access_token, file_path)
        job_id = convert_pdf_to_docx(access_token, asset_id)
        poll_and_download_result(access_token, job_id, output_path)
        
        print("Conversion process completed successfully!")
    except Exception as e:
        print(f"Main process error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
