AI-Powered File Audit Assistant
===============================

A modular Python tool for scanning, classifying, deduplicating, and analyzing files using GPT-4. Designed for maintainability, transparency, and user empowerment.

Features
--------

- Directory Scanning: Recursively scans files and extracts metadata
- Extension Classification: Groups files by type using config or GPT
- Deduplication: Detects duplicates via hash comparison
- Anomaly Detection: GPT flags suspicious files, mismatches, and naming issues
- Audit Summary: GPT generates natural-language reports
- Logging: Saves structured audit logs to disk

Setup
-----

1. Clone the Repo

    git clone https://github.com/your-username/file-audit-assistant.git
    cd file-audit-assistant

2. Install Dependencies

    pip install -r requirements.txt

3. Add Your OpenAI API Key

Create a `.env` file:

    OPENAI_API_KEY=your-key-here

Configuration
-------------

Edit `config/audit_config.yaml`:

    scan_path: "./target_directory"
    extension_query: "Include common extensions for documents, spreadsheets, images, and code files."
    deduplication:
      method: "hash"
      ignore_extensions: [".tmp", ".log"]
    logging:
      level: "INFO"
    use_gpt_summary: true
    use_gpt_extensions: true

Usage
-----

Run the assistant:

    python main.py

Outputs:
- GPT-generated summary
- Anomaly report
- JSON log in `logs/`

Project Structure
-----------------

    file_audit_assistant/
    ├── config/
    ├── core/
    ├── logs/
    ├── main.py
    ├── requirements.txt
    └── README.md

Extensibility
-------------

This project is modular by design. You can easily add:
- MIME-based classification
- CLI flags (--summary, --anomalies)
- Exporters for Excel, CSV, or JSON
- Heuristic-based tagging or remediation

License
-------

MIT License. Use freely and modify as needed.