import yaml
from core import (
    scanner,
    classifier,
    deduplicator,
    summarizer,
    anomaly_detector,
    extension_resolver,
    logger
)

def load_config(path="config/audit_config.yaml"):
    with open(path, 'r') as f:
        return yaml.safe_load(f)

def main():
    config = load_config()

    # Step 1: Resolve extensions via GPT if enabled
    if config.get("use_gpt_extensions"):
        yaml_text = extension_resolver.query_extensions(config["extension_query"])
        extensions = yaml.safe_load(yaml_text)
    else:
        extensions = config.get("extensions", {})

    # Step 2: Scan directory and extract metadata
    files = scanner.scan(config["scan_path"])
    metadata = scanner.extract_metadata(files)

    # Step 3: Classify files
    classified = classifier.classify(metadata, extensions)

    # Step 4: Deduplicate
    duplicates = deduplicator.find_duplicates(metadata)

    # Step 5: Detect anomalies via GPT
    anomalies = anomaly_detector.detect_anomalies(metadata)

    # Step 6: Generate GPT summary
    if config.get("use_gpt_summary"):
        summary = summarizer.summarize_audit({
            "classified": classified,
            "duplicates": duplicates,
            "anomalies": anomalies
        })
        print("\n📋 GPT Summary:\n", summary)

    # Step 7: Log results
    logger.generate_log(classified, duplicates, anomalies)

    # Optional: Print anomalies
    print("\n GPT Anomaly Report:\n", anomalies)

if __name__ == "__main__":
    main()