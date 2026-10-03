import requests

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
})

def get_company_wikipedia_summary(company_name):
    try:
        # Search for page title
        s_url = f"https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch={requests.utils.quote(company_name)}&format=json"
        res = session.get(s_url, timeout=5).json()
        search_results = res.get('query', {}).get('search', [])
        if not search_results:
            return None
        page_title = search_results[0]['title']
        
        # Get intro extract
        e_url = f"https://en.wikipedia.org/w/api.php?action=query&prop=extracts&exintro=1&explaintext=1&titles={requests.utils.quote(page_title)}&format=json"
        e_res = session.get(e_url, timeout=5).json()
        pages = e_res.get('query', {}).get('pages', {})
        for pid, pdata in pages.items():
            extract = pdata.get('extract', '').strip()
            if extract:
                # Return the first 2-3 sentences or up to ~300 chars suitable for cell B8
                paragraphs = [p for p in extract.split('\n') if p.strip()]
                return paragraphs[0] if paragraphs else extract
    except Exception as e:
        print(f"Wiki error: {e}")
    return None

for name in ['Adani Enterprises', 'Tata Steel', 'ITC Limited', 'Nestle India', 'JSW Steel']:
    summary = get_company_wikipedia_summary(name)
    print(f"\n[{name}]:\n{summary}")
