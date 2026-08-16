import json
import urllib.request
import urllib.error

OLLAMA_API_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "qwen3:8b"

SYSTEM_PROMPT = """
You are an intelligent email parser for college placement updates.
Your job is to read an email (Subject and Body) and extract structured information.
You must output a strict JSON object with the following keys exactly:
- "category": Must be one of ["PLACEMENT", "ACADEMIC", "OTHER"]
- "subtype": If PLACEMENT, must be one of ["NEW_COMPANY", "SHORTLIST_TEST", "SELECTION", "UPDATE"]. If ACADEMIC, use "ACADEMIC". Otherwise use "OTHER".
- "extracted_data": A dictionary. If category is PLACEMENT, extract relevant fields such as:
  - "Company Name"
  - "Job Title(s)"
  - "CTC"
  - "Intern Stipend"
  - "Registration Deadline"
  - "Shortlisted Count"
  - "Test Date"
  - "Test/Assessment"
  - "Selected Count"
  Only include fields that are present in the text.
- "digest": A short 1-2 sentence string summarizing the email, designed to be used directly in a Windows toast notification (e.g. "Google - Software Engineer | 20 LPA | Deadline: 12th Aug"). 

Do not include markdown blocks, output raw JSON only.
"""

def analyze_email(subject, body):
    prompt = f"Subject: {subject}\n\nBody:\n{body}\n\nParse this email into the requested JSON format."
    
    data = {
        "model": MODEL_NAME,
        "prompt": prompt,
        "system": SYSTEM_PROMPT,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.0
        }
    }
    
    req = urllib.request.Request(OLLAMA_API_URL, data=json.dumps(data).encode('utf-8'), headers={'Content-Type': 'application/json'})
    
    try:
        with urllib.request.urlopen(req) as response:
            result = json.loads(response.read().decode('utf-8'))
            response_text = result.get('response', '{}')
            try:
                parsed_json = json.loads(response_text)
                return parsed_json
            except json.JSONDecodeError:
                print(f"Failed to parse JSON from LLM: {response_text}")
                return None
    except Exception as e:
        print(f"\n❌ Error connecting to Ollama: {e}")
        print(f"Make sure Ollama is running and '{MODEL_NAME}' is installed.")
        return None
