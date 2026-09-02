"""
azure_sync.py — Azure Blob Storage helpers for the Caviar project.

Handles:
  - Importing (downloading) the shared Excel checklist from Azure
  - Exporting (uploading) the local Excel checklist to Azure
  - Auto-uploading trial files (_A.MP4, _B.MP4, .txt) after a trial completes

Config file: azure_config.json (sits next to this file)
  {
      "connection_string": "DefaultEndpointsProtocol=https;AccountName=...;...",
      "container_name":    "vr-dataset"
  }

Blob paths mirror the local folder structure:
  DataCollection_Checklist.xlsx
  VR_Dataset/P1/Pain/Standard/T1/P1_P_S_T1_A.MP4
  VR_Dataset/P1/Pain/Standard/T1/P1_P_S_T1_B.MP4
  VR_Dataset/P1/Pain/Standard/T1/P1_P_S_T1.txt
"""

import json
import os

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "azure_config.json")
_EXCEL_BLOB_NAME = "DataCollection_Checklist.xlsx"


# ── Config ─────────────────────────────────────────────────────────────────

def _load_config():
    if not os.path.exists(_CONFIG_PATH):
        raise FileNotFoundError(
            f"Azure config not found: {_CONFIG_PATH}\n"
            "Create it with your connection_string and container_name."
        )
    with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cs = cfg.get("connection_string", "")
    if not cs or cs == "YOUR_CONNECTION_STRING_HERE":
        raise ValueError(
            "azure_config.json: connection_string is not set.\n"
            "Paste your Azure Storage connection string there."
        )
    return cfg


def _get_container_client():
    from azure.storage.blob import BlobServiceClient
    cfg = _load_config()
    client = BlobServiceClient.from_connection_string(cfg["connection_string"])
    return client.get_container_client(cfg["container_name"])


# ── Excel Import / Export ───────────────────────────────────────────────────

def import_excel(local_path, log_fn=None):
    """Download the shared Excel checklist from Azure to local_path.

    Raises on failure. log_fn(msg) is called with progress lines.
    """
    container = _get_container_client()
    blob = container.get_blob_client(_EXCEL_BLOB_NAME)

    if log_fn:
        log_fn(f"  Downloading {_EXCEL_BLOB_NAME} from Azure...")

    # Back up current local copy before overwriting (timestamped)
    if os.path.exists(local_path):
        from datetime import datetime
        import shutil
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = local_path + f".bak_{ts}"
        try:
            shutil.copy2(local_path, backup)
            if log_fn:
                log_fn(f"  ✓ Local backup saved as {os.path.basename(backup)}")
        except Exception as e:
            if log_fn:
                log_fn(f"  ⚠ Could not back up local file: {e}")

    os.makedirs(os.path.dirname(os.path.abspath(local_path)), exist_ok=True)
    with open(local_path, "wb") as f:
        stream = blob.download_blob()
        stream.readinto(f)

    size_kb = os.path.getsize(local_path) // 1024
    if log_fn:
        log_fn(f"  ✓ Imported {_EXCEL_BLOB_NAME} ({size_kb} KB)")


def export_excel(local_path, log_fn=None):
    """Upload the local Excel checklist to Azure, overwriting the shared copy.

    Raises on failure. log_fn(msg) is called with progress lines.
    """
    if not os.path.exists(local_path):
        raise FileNotFoundError(f"Local Excel not found: {local_path}")

    container = _get_container_client()
    blob = container.get_blob_client(_EXCEL_BLOB_NAME)

    size_kb = os.path.getsize(local_path) // 1024
    if log_fn:
        log_fn(f"  Uploading {_EXCEL_BLOB_NAME} ({size_kb} KB) to Azure...")

    with open(local_path, "rb") as f:
        blob.upload_blob(f, overwrite=True)

    if log_fn:
        log_fn(f"  ✓ Exported {_EXCEL_BLOB_NAME} to Azure")


# ── Trial File Upload ───────────────────────────────────────────────────────

def _local_to_blob_name(local_path, base_folder):
    """Convert a local absolute path to a blob name.

    E:\\VR_Dataset\\P1\\Pain\\Standard\\T1\\P1_P_S_T1_A.MP4
      → VR_Dataset/P1/Pain/Standard/T1/P1_P_S_T1_A.MP4
    """
    rel = os.path.relpath(local_path, os.path.dirname(base_folder))
    return rel.replace("\\", "/")


def upload_trial_files(trial_code, save_folder, base_folder, log_fn=None):
    """Upload _A.MP4, _B.MP4, and .txt for a trial to Azure Blob Storage.

    Missing files are skipped with a warning. Returns a dict:
      { filename: 'uploaded' | 'missing' | 'error:<msg>' }
    """
    files = {
        f"{trial_code}_A.MP4": os.path.join(save_folder, f"{trial_code}_A.MP4"),
        f"{trial_code}_B.MP4": os.path.join(save_folder, f"{trial_code}_B.MP4"),
        f"{trial_code}.txt":   os.path.join(save_folder, f"{trial_code}.txt"),
    }

    container = _get_container_client()
    report = {}

    for fname, local_path in files.items():
        if not os.path.exists(local_path):
            report[fname] = "missing"
            if log_fn:
                log_fn(f"  ⚠ {fname}: not found locally — skipped")
            continue
        blob_name = _local_to_blob_name(local_path, base_folder)
        size_mb = os.path.getsize(local_path) // (1024 * 1024)
        try:
            if log_fn:
                log_fn(f"  Uploading {fname} ({size_mb} MB)...")
            blob = container.get_blob_client(blob_name)
            with open(local_path, "rb") as f:
                blob.upload_blob(f, overwrite=True)
            report[fname] = "uploaded"
            if log_fn:
                log_fn(f"  ✓ {fname} uploaded")
        except Exception as e:
            report[fname] = f"error:{e}"
            if log_fn:
                log_fn(f"  ✗ {fname} failed: {e}")

    return report
