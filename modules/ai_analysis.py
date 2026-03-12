import os
import json
import base64
import requests


SYSTEM_PROMPT = """You are an expert investor relations analyst. You will be given company background documents
and a CDF (Contact Data Field) taxonomy. Your job is to recommend the most appropriate CDF filter values
for each dimension based on the company described in the documents.

Be CONSERVATIVE — only select values clearly supported by the documents. Do not pad the list.
Return ONLY valid JSON with no preamble, markdown, or explanation.

Output format:
{
  "industry": ["value1", "value2"],
  "style": ["value1", "value2"],
  "mcap": ["value1", "value2"],
  "geo": ["value1", "value2"],
  "reasoning": {
    "industry": "brief explanation",
    "style": "brief explanation",
    "mcap": "brief explanation",
    "geo": "brief explanation"
  }
}

Geography rules for US-listed companies:
- Default to: North America, *Global
- Do NOT include *Global (ex US) — this explicitly excludes the US
- Do NOT include ***Intl ADR — this is for non-US companies listed on US exchanges

Market Cap guidance:
- Under ~$500M → Micro, Small
- ~$500M–$2.5B → Small, Mid
- ~$2.5B–$10B → Mid, Large
- ~$10B–$50B → Large, Mega
- Over ~$50B → Mega

Investment Style guidance:
- Stable, cash-generative, mature → Value, GARP
- High-growth, pre-profit → Growth, Aggressive growth
- Moderate growth, profitable → GARP, Blend
- Dividend/income focus → Yield"""


def analyze_documents(file_data, taxonomy):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return {'error': 'Anthropic API key not configured'}

    taxonomy_text = _format_taxonomy(taxonomy)

    content = []
    for f in file_data:
        if f['name'].lower().endswith('.pdf') or 'pdf' in f.get('type', ''):
            b64 = base64.standard_b64encode(f['data']).decode('utf-8')
            content.append({
                'type': 'document',
                'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64}
            })
        else:
            try:
                text = f['data'].decode('utf-8', errors='ignore')
                content.append({'type': 'text', 'text': f"File: {f['name']}\n\n{text[:50000]}"})
            except Exception:
                pass

    content.append({
        'type': 'text',
        'text': f"""Based on the company documents above, recommend CDF filter values from this taxonomy:

{taxonomy_text}

Return only valid JSON matching the specified output format."""
    })

    response = requests.post(
        'https://api.anthropic.com/v1/messages',
        headers={
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
        },
        json={
            'model': 'claude-sonnet-4-20250514',
            'max_tokens': 2000,
            'system': SYSTEM_PROMPT,
            'messages': [{'role': 'user', 'content': content}],
        },
        timeout=60
    )
    response.raise_for_status()

    text = response.json()['content'][0]['text']
    clean = text.strip()
    if clean.startswith('```'):
        clean = clean.split('```')[1]
        if clean.startswith('json'):
            clean = clean[4:]
    clean = clean.strip().rstrip('`').strip()

    return json.loads(clean)


def _format_taxonomy(taxonomy):
    if not taxonomy:
        return "Taxonomy not available — use standard CDF values."
    lines = []
    for field_type, values in taxonomy.items():
        lines.append(f"\n{field_type}:")
        for v in values:
            lines.append(f"  - {v['value']}")
    return '\n'.join(lines)
