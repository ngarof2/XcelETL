import os
import hashlib
import mimetypes
from pathlib import Path
from typing import List, Dict

def scan(directory: str) -> List[Path]:
    return [f for f in Path(directory).rglob("*") if f.is_file()]

def hash_file(path: Path, method="md5") -> str:
    hasher = getattr(hashlib, method)()
    with open(path, 'rb') as f:
        while chunk := f.read(8192):
            hasher.update(chunk)
    return hasher.hexdigest()

def get_mime_type(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path)
    return mime or "unknown"

def get_preview(path: Path, max_lines=3) -> str:
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return "\n".join([next(f).strip() for _ in range(max_lines)])
    except Exception:
        return ""

def extract_metadata(files: List[Path]) -> List[Dict]:
    metadata = []
    for f in files:
        stat = f.stat()
        metadata.append({
            "name": f.name,
            "path": str(f),
            "extension": f.suffix.lower(),
            "size_kb": stat.st_size // 1024,
            "modified": stat.st_mtime,
            "created": stat.st_ctime,
            "mime_type": get_mime_type(f),
            "hash": hash_file(f),
            "preview": get_preview(f) if f.suffix.lower() in [".txt", ".csv", ".log"] else ""
        })
    return metadata