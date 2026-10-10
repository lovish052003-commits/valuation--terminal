"""
news_service.py
Service for fetching live company market news via Google News RSS
and computing AI market sentiment via Google Gemini.
"""

import os
import re
import json
import urllib.parse
import email.utils
from datetime import datetime

try:
    import feedparser
except ImportError:
    feedparser = None


def fetch_company_news(company_name, max_articles=8):
    """
    Retrieves live news articles for a company from Google News RSS (India edition).
    Returns list of dicts with keys: title, link, published, source.
    """
    if not company_name or feedparser is None:
        return []

    try:
        query = urllib.parse.quote(f"{company_name} stock share news India")
        rss_url = f"https://news.google.com/rss/search?q={query}&hl=en-IN&gl=IN&ceid=IN:en"
        
        feed = feedparser.parse(rss_url)
        items = []

        for entry in feed.entries[:max_articles]:
            raw_title = entry.get('title', '').strip()
            link = entry.get('link', '')
            pub_raw = entry.get('published', '')

            # Extract source
            source = (entry.get('source') or {}).get('title', '')
            title = raw_title
            if ' - ' in raw_title:
                parts = raw_title.rsplit(' - ', 1)
                if not source:
                    source = parts[1].strip()
                title = parts[0].strip()
            if not source:
                source = 'Market News'

            # Format published date cleanly
            pub_formatted = pub_raw
            if pub_raw:
                try:
                    dt = email.utils.parsedate_to_datetime(pub_raw)
                    pub_formatted = dt.strftime('%d %b %Y, %H:%M')
                except Exception:
                    pub_formatted = pub_raw

            items.append({
                'title': title,
                'link': link,
                'published': pub_formatted,
                'source': source
            })

        return items
    except Exception as e:
        print(f"[News Service] Error fetching news for '{company_name}': {e}")
        return []


def analyze_news_sentiment(company_name, news_items, api_key=None, model=None, skip_ai=False):
    """
    Analyzes news headlines using Gemini to produce institutional sentiment scoring.
    Returns:
    {
        'sentiment': 'BULLISH' | 'BEARISH' | 'NEUTRAL',
        'score': int (0-100),
        'summary': '2-sentence institutional summary of the news flow'
    }
    """
    fallback = {
        'sentiment': 'NEUTRAL',
        'score': 50,
        'summary': 'Live market news feed active.'
    }

    if skip_ai or not news_items:
        return fallback

    gemini_key = api_key or os.environ.get('GEMINI_API_KEY')
    if not gemini_key:
        return fallback

    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=gemini_key)
        headlines_text = "\n".join([f"- {it.get('title', '')} ({it.get('source', '')})" for it in news_items])

        system_instruction = (
            "Analyze these stock news headlines. Return a strict JSON response with:\n"
            "{\n"
            '  "sentiment": "BULLISH" | "BEARISH" | "NEUTRAL",\n'
            '  "score": int between 0 and 100,\n'
            '  "summary": "2-sentence institutional summary of the news flow"\n'
            "}"
        )
        prompt = f"Target Company: {company_name}\nRecent Headlines:\n{headlines_text}\n\nStrict JSON:"

        target_model = model or "gemini-2.5-flash"
        models_to_try = [target_model, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
        seen_models = set()
        models = [m for m in models_to_try if not (m in seen_models or seen_models.add(m))]

        for m in models:
            try:
                config = types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    temperature=0.2,
                    response_mime_type="application/json"
                )
                response = client.models.generate_content(
                    model=m,
                    contents=prompt,
                    config=config
                )
                text = (response.text or "").strip()
                if text.startswith("```"):
                    text = re.sub(r"^```(?:json)?\s*", "", text)
                    text = re.sub(r"\s*```$", "", text)
                data = json.loads(text)

                sentiment = str(data.get("sentiment", "NEUTRAL")).upper().strip()
                if sentiment not in ("BULLISH", "BEARISH", "NEUTRAL"):
                    sentiment = "NEUTRAL"

                score = int(data.get("score", 50))
                score = max(0, min(100, score))

                summary = str(data.get("summary", "Live market news feed active.")).strip()

                return {
                    'sentiment': sentiment,
                    'score': score,
                    'summary': summary
                }
            except Exception as e_model:
                continue

        return fallback
    except Exception as e:
        print(f"[News Service] AI Sentiment notice: {e}")
        return fallback
