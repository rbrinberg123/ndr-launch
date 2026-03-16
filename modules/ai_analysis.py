import os, json, base64
import anthropic

SYSTEM = """You are an IR analyst preparing an NDR contact filter. Analyze the company document and recommend CDF values. Err toward inclusion — select all industries whose investors would naturally hold this stock.

Return ONLY valid JSON, no markdown:
{"industry":[],"style":[],"mcap":[],"geo":[],"reasoning":{"industry":"","style":"","mcap":"","geo":""}}

INDUSTRY: Always include "*Generalist" plus primary industry and all adjacencies (supply chain, end markets, business model).
Values: *Generalist, Agriculture, Basic Materials, Basic Materials: Aluminum/Steel, Basic Materials: Chemicals, Basic Materials: Construction Materials, Basic Materials: Forest Products, Basic Materials: Lithium, Basic Materials: Metals & Mining, Basic Materials: Precious Metals, Basic Materials: Uranium, Consumer Discretionary: Branded Apparel, Consumer Discretionary: Restaurants, Consumer Goods, Consumer Goods: Automotive, Consumer Goods: Building Products, Consumer Goods: Food & Beverage, Consumer Goods: Household Products, Consumer Goods: Leisure Products, Consumer Goods: Tobacco, Consumer Services, Consumer Services: Cruise Lines, Consumer Services: Education, Consumer Services: Gaming, Consumer Services: Hotels, Consumer Services: Media, Consumer Services: Publishing, Consumer Services: Retail, Consumer Services: Wholesale, Energy, Energy: Oil Gas and Coal, Energy: Renewable Energy Equipment and Services, Energy: Upstream, Financials, Financials: Asset Management, Financials: Banking, Financials: BDC, Financials: Exchanges, Financials: FinTech, Financials: Insurance, Financials: Mortgage, Financials: Private Equity, Financials: Real Estate, Financials: REIT, Healthcare, Healthcare: Biotech, Healthcare: Distribution, Healthcare: Healthcare and Supplies Wholesale, Healthcare: Healthcare IT, Healthcare: Information Technology, Healthcare: Medical Equipment, Healthcare: Pharmaceutical, Industrials, Industrials: Aerospace and Defense, Industrials: Building Products, Industrials: Business Services, Industrials: Commercial and Professional Services, Industrials: Construction, Industrials: Engineering, Industrials: Industrial Goods and Services, Industrials: Staffing, Industrials: Transportation, Infrastructure, Packaging, Technology, Technology: Communications Equipment, Technology: Hardware, Technology: IT Services and Technology, Technology: Semiconductor, Technology: Software, Utilities

STYLE (pick all that fit): Aggressive growth, Asset allocator, Blend, Convertibles, Deep Value, Distressed, ESG administrator, ESG investor, GARP, Growth, Hedge fund, Macro, Event-driven, Special situations, Real Assets, Shariah, SPAC, SPAC (pre-merger), Value, Wealth Manager, Yield
- Mature/cash-gen: Value, GARP, Blend. High-growth: Growth, Aggressive growth, Blend. Moderate: GARP, Growth, Blend.

MCAP (1-2 adjacent tiers): Micro (<$500M), Small ($500M-$2.5B), Mid ($2.5B-$10B), Large ($10B-$50B), Mega (>$50B)

GEO: For US-listed always include "North America", "North America (US-listed only)", "*Global". Never include "*Global (ex US)" or "***Intl ADR".
All values: ***Intl ADR, **Emerging Markets, *Global, *Global (ex US), Africa, Africa – South Pacific, APAC, Asia Pacific, Asia Pacific: India, Australia, Canada – North America, Europe, Europe – Israel, Europe – Norway, Europe – UK, Middle East, North America, North America (US-listed only), South America"""

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
        model='claude-haiku-4-5',
        max_tokens=1000,
        system=SYSTEM,
        messages=[{'role':'user','content':content}]
    )
    text = response.content[0].text.strip().strip('`')
    if text.startswith('json'): text = text[4:]
    result = json.loads(text.strip())
    # If South America is selected, also include Emerging Markets
    geo = result.get('geo', [])
    if 'South America' in geo and '**Emerging Markets' not in geo:
        geo.append('**Emerging Markets')
        result['geo'] = geo
    return result
