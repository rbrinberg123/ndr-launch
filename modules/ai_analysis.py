import os, json, base64
import anthropic

SYSTEM = """You are an expert investor relations analyst. Given company documents, recommend conservative
CDF filter values. Return ONLY valid JSON, no markdown, no preamble.
Format: {"industry":[],"style":[],"mcap":[],"geo":[],"reasoning":{"industry":"","style":"","mcap":"","geo":""}}
Geography for US-listed: always include "North America" and "*Global". Never include "*Global (ex US)" or "***Intl ADR".
Market Cap: <$500M->Micro/Small; $500M-$2.5B->Small/Mid; $2.5B-$10B->Mid/Large; $10B-$50B->Large/Mega; >$50B->Mega."""

def analyze_documents(file_data, taxonomy=None):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key: return {'error': 'API key not configured'}
    client = anthropic.Anthropic(api_key=api_key)
    content = []
    for f in file_data:
        mt = f.get('type','')
        if 'pdf' in mt or f['name'].lower().endswith('.pdf'):
            content.append({'type':'document','source':{'type':'base64',
                'media_type':'application/pdf','data':base64.standard_b64encode(f['data']).decode()}})
        else:
            content.append({'type':'text','text':f"File: {f['name']}\n\n{f['data'].decode('utf-8','ignore')[:40000]}"})
    content.append({'type':'text','text':'Analyze and return JSON only.'})
    response = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=2000,
        system=SYSTEM,
        messages=[{'role':'user','content':content}]
    )
    text = response.content[0].text.strip().strip('`')
    if text.startswith('json'): text = text[4:]
    return json.loads(text.strip())
