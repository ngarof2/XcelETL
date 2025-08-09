import openai
from typing import Dict

def summarize_audit(audit_data: Dict) -> str:
    prompt = f"""
You are a file audit assistant. Summarize the following audit results:
- Classified files by type
- Duplicate files
- Anomalies detected

Here is the data:
{audit_data}

Provide a concise summary with bullet points and key insights.
"""

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content