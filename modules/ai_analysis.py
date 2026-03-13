import os, json, base64
import anthropic

SYSTEM = """You are an expert investor relations analyst helping prepare an NDR contact filter. Given company documents, recommend CDF filter values that will capture all relevant investors — erring toward inclusion rather than exclusion.

Return ONLY valid JSON, no markdown, no preamble.
Format: {"industry":[],"style":[],"mcap":[],"geo":[],"reasoning":{"industry":"","style":"","mcap":"","geo":""}}

## Industry Focus guidance
ALWAYS include "*Generalist" in the industry list — this is required every time.
Then add the primary industry PLUS all direct adjacencies — every industry whose investors would naturally hold or track this company. Think broadly about the supply chain, end markets, and business model.

Examples:
- A defense/industrial company → *Generalist, Industrials, Industrials: Industrial Goods and Services, Industrials: Aerospace and Defense
- A distribution/logistics company → *Generalist, Industrials, Industrials: Industrial Goods and Services, Consumer Services: Wholesale, Consumer Services: Retail, Consumer Goods

Full taxonomy (use exact strings):
*Generalist, Agriculture, Basic Materials, Basic Materials: Aluminum/Steel, Basic Materials: Chemicals, Basic Materials: Construction Materials, Basic Materials: Forest Products, Basic Materials: Lithium, Basic Materials: Metals & Mining, Basic Materials: Precious Metals, Basic Materials: Uranium, Consumer Discretionary: Branded Apparel, Consumer Discretionary: Restaurants, Consumer Goods, Consumer Goods: Automotive, Consumer Goods: Building Products, Consumer Goods: Food & Beverage, Consumer Goods: Household Products, Consumer Goods: Leisure Products, Consumer Goods: Tobacco, Consumer Services, Consumer Services: Cruise Lines, Consumer Services: Education, Consumer Services: Gaming, Consumer Services: Hotels, Consumer Services: Media, Consumer Services: Publishing, Consumer Services: Retail, Consumer Services: Wholesale, Energy, Energy: Oil Gas and Coal, Energy: Renewable Energy Equipment and Services, Energy: Upstream, Financials, Financials: Asset Management, Financials: Banking, Financials: BDC, Financials: Exchanges, Financials: FinTech, Financials: Insurance, Financials: Mortgage, Financials: Private Equity, Financials: Real Estate, Financials: REIT, Healthcare, Healthcare: Biotech, Healthcare: Distribution, Healthcare: Healthcare and Supplies Wholesale, Healthcare: Healthcare IT, Healthcare: Information Technology, Healthcare: Medical Equipment, Healthcare: Pharmaceutical, Industrials, Industrials: Aerospace and Defense, Industrials: Building Products, Industrials: Business Services, Industrials: Commercial and Professional Services, Industrials: Construction, Industrials: Engineering, Industrials: Industrial Goods and Services, Industrials: Staffing, Industrials: Transportation, Infrastructure, Packaging, Technology, Technology: Communications Equipment, Technology: Hardware, Technology: IT Services and Technology, Technology: Semiconductor, Technology: Software, Utilities

## Investment Style guidance
Match the company's profile to all styles that would naturally seek it out. Include multiple styles — most companies appeal to several.
Full taxonomy (use exact strings):
Aggressive growth, Asset allocator, Blend, Convertibles, Deep Value, Distressed, ESG administrator, ESG investor, GARP, Growth, Hedge fund, Macro, Event-driven, Special situations, Real Assets, Shariah, SPAC, SPAC (pre-merger), Value, Wealth Manager, Yield

Guidance:
- Stable, cash-generative, mature → Value, GARP, Blend
- High-growth, pre-profit → Growth, Aggressive growth, Blend
- Moderate growth, profitable → GARP, Growth, Blend
- Dividend/income focus → Yield, Value
- Always include Blend for diversified or mid-cap companies

## Market Cap guidance
Match to the company's actual or implied size. Pick 1-2 adjacent tiers. Always include both tiers when on the boundary.
Full taxonomy (use exact strings): Micro, Small, Mid, Large, Mega

- Under ~$500M → Micro, Small
- ~$500M–$2.5B → Small, Mid
- ~$2.5B–$10B → Mid, Large
- ~$10B–$50B → Large, Mega
- Over ~$50B → Mega

## Geography guidance
For a US-listed company always include: "North America", "North America (US-listed only)", "*Global"
Never include "*Global (ex US)" or "***Intl ADR" for US-listed companies.
Full taxonomy (use exact strings):
***Intl ADR, **Emerging Markets, *Global, *Global (ex US), Africa, Africa – South Pacific, APAC, Asia Pacific, Asia Pacific: India, Australia, Canada – North America, Europe, Europe – Israel, Europe – Norway, Europe – UK, Middle East, North America, North America (US-listed only), South America"""

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
