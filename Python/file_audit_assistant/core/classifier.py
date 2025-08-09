from typing import List, Dict

def classify(metadata: List[Dict], extension_groups: Dict[str, List[str]]) -> Dict[str, List[Dict]]:
    classified = {group: [] for group in extension_groups}
    classified["unclassified"] = []

    for file in metadata:
        matched = False
        for group, extensions in extension_groups.items():
            if file["extension"] in extensions:
                classified[group].append(file)
                matched = True
                break
        if not matched:
            classified["unclassified"].append(file)

    return classified