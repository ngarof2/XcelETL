import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List

def generate_log(classified: Dict[str, List[Dict]], duplicates: List[Dict], anomalies: str, output_dir="logs"):
    Path(output_dir).mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(output_dir) / f"audit_log_{timestamp}.json"

    log_data = {
        "timestamp": timestamp,
        "classified": classified,
        "duplicates": duplicates,
        "anomalies": anomalies
    }

    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump(log_data, f, indent=2)

    print(f"\n📝 Log saved to: {log_path}")