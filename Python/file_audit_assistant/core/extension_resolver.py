import openai

def query_extensions(description: str) -> dict:
    prompt = f"""
You are a file classification assistant. Based on the following description, list common file extensions grouped by category:
"{description}"

Respond in valid YAML format like:
documents: [".docx", ".pdf", ".txt"]
spreadsheets: [".xlsx", ".csv"]
images: [".jpg", ".png"]
code: [".py", ".js", ".html"]
"""

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    yaml_text = response.choices[0].message.content
    return yaml_text