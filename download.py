import os
import gdown

FOLDER_ID = "1uWk690Cg2I2wJsSq_e-uA32l3B8Vuko_"
OUTPUT_DIR = "data"

os.makedirs(OUTPUT_DIR, exist_ok=True)
url = f"https://drive.google.com/drive/folders/{FOLDER_ID}"

print(f"Downloading from Google Drive folder: {FOLDER_ID}...")
try:
    gdown.download_folder(url=url, output=OUTPUT_DIR, quiet=False, use_cookies=False)
    print("Download completed successfully.")
except Exception as e:
    print(f"Download warning/error: {e}")
