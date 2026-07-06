import subprocess
import json
import urllib.request
import urllib.error
import ssl
import os

proc = subprocess.run(
    ["git", "credential", "fill"],
    input="protocol=https\nhost=github.com\n\n",
    capture_output=True, text=True,
    cwd=r"c:\Users\socia\Documents\GitHub\allone"
)

token = [l for l in proc.stdout.splitlines() if l.startswith("password=")][0].split("=", 1)[1]
ctx = ssl.create_default_context()
repo = "hakanakaslanx-code/allone"

url = f"https://api.github.com/repos/{repo}/releases/tags/v6.5.1"
req = urllib.request.Request(url)
req.add_header("Authorization", f"Bearer {token}")
req.add_header("Accept", "application/vnd.github+json")

try:
    with urllib.request.urlopen(req, context=ctx) as r:
        resp = json.loads(r.read().decode())
        release_id = resp["id"]
        upload_url = resp["upload_url"].split("{")[0]
        html_url = resp["html_url"]
        print("Release found: " + html_url)
except urllib.error.HTTPError as e:
    print("Error getting release: " + str(e.code) + " - " + e.read().decode())
    exit(1)

zip_path = r"dist\AllOneTools.zip"
file_size = os.path.getsize(zip_path)
with open(zip_path, "rb") as f:
    upload_req_url = f"{upload_url}?name=AllOneTools.zip"
    req_up = urllib.request.Request(upload_req_url, data=f, method="POST")
    req_up.add_header("Authorization", f"Bearer {token}")
    req_up.add_header("Accept", "application/vnd.github+json")
    req_up.add_header("Content-Type", "application/zip")
    req_up.add_header("Content-Length", str(file_size))
    
    print("Uploading asset...")
    try:
        with urllib.request.urlopen(req_up, context=ctx) as r_up:
            up_resp = json.loads(r_up.read().decode())
            print("Asset uploaded: " + up_resp["browser_download_url"])
    except urllib.error.HTTPError as e:
        print("Error uploading asset: " + str(e.code) + " - " + e.read().decode())
