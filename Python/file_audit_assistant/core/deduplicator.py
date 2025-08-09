from typing import List, Dict

def find_duplicates(metadata: List[Dict]) -> List[Dict]:
    seen = {}
    duplicates = []

    for file in metadata:
        file_hash = file.get("hash")
        if not file_hash:
            continue
        if file_hash in seen:
            duplicates.append({
                "original": seen[file_hash],
                "duplicate": file
            })
        else:
            seen[file_hash] = file

    return duplicates