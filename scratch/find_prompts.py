import json

transcript_path = r'C:\Users\LENOVO\.gemini\antigravity-ide\brain\6a33781f-32df-40dc-8ee9-53ed533f7cd5\.system_generated\logs\transcript.jsonl'
with open(transcript_path, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            d = json.loads(line)
            if d.get('type') == 'USER_INPUT':
                print(f"--- USER STEP {d.get('step_index')} ---")
                print(d.get('content'))
                print()
        except Exception:
            pass
