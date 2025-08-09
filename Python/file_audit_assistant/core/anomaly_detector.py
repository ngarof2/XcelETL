import openai

def detect_anomalies(file_metadata: list) -> str:
    prompt = f"""
You are a file audit assistant. Analyze the following list of files with metadata and flag anomalies:
- Unexpected extensions
- Suspicious names (e.g., temp, backup, copy)
- Large files with simple names
- Files that don’t match their extension

Here is the data:
{file_metadata}

Respond with a concise list of flagged files and reasons.
"""

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content