"""
events_engine.py
Corporate Events & Actions Engine for Valuation Terminal.
Fetches and consolidates corporate actions (Dividends, Stock Splits, Board Meetings, Earnings)
using yfinance and Screener.in corporate disclosures.
"""

import re
from datetime import datetime, date
import requests
from bs4 import BeautifulSoup

try:
    import yfinance as yf
except ImportError:
    yf = None


def fetch_corporate_events(ticker_symbol, max_events=20):
    """
    Consolidates corporate events (Dividends, Splits, Earnings, Board Meetings)
    for a given ticker symbol into a unified, chronologically sorted list.

    Returns:
    [
        {
            'event_type': 'Dividend' | 'Earnings' | 'Board Meeting' | 'Stock Split' | 'Bonus Issue',
            'event_date': '15 Oct 2026',
            'details': '₹5.25 per share',
            'is_upcoming': True | False
        },
        ...
    ]
    """
    if not ticker_symbol:
        return []

    clean_ticker = re.sub(r'[^A-Za-z0-9_&-]', '', str(ticker_symbol)).strip().upper()
    if not clean_ticker:
        return []

    today = date.today()
    events = []
    seen = set()

    # 1. Fetch from yfinance (Actions & Calendar)
    if yf is not None:
        sym = f"{clean_ticker}.NS" if not clean_ticker.endswith(('.NS', '.BO')) else clean_ticker
        try:
            t = yf.Ticker(sym)

            # Calendar: Earnings Date and Ex-Dividend Date
            cal = t.calendar
            if isinstance(cal, dict):
                # Earnings Date
                ed = cal.get('Earnings Date')
                if ed:
                    dates = ed if isinstance(ed, list) else [ed]
                    for d in dates:
                        parsed_date = d.date() if hasattr(d, 'date') else d
                        if isinstance(parsed_date, date):
                            key = ('Earnings', str(parsed_date))
                            if key not in seen:
                                seen.add(key)
                                is_up = parsed_date >= today
                                events.append({
                                    'event_type': 'Earnings',
                                    'raw_date': parsed_date,
                                    'event_date': parsed_date.strftime('%d %b %Y'),
                                    'details': 'Quarterly Financial Results Announcement',
                                    'is_upcoming': is_up
                                })

                # Ex-Dividend Date
                ex_div = cal.get('Ex-Dividend Date')
                if ex_div:
                    parsed_date = ex_div.date() if hasattr(ex_div, 'date') else ex_div
                    if isinstance(parsed_date, date):
                        key = ('Dividend', str(parsed_date), 'ex')
                        if key not in seen:
                            seen.add(key)
                            is_up = parsed_date >= today
                            events.append({
                                'event_type': 'Dividend',
                                'raw_date': parsed_date,
                                'event_date': parsed_date.strftime('%d %b %Y'),
                                'details': 'Ex-Dividend Date',
                                'is_upcoming': is_up
                            })

            # Actions: Historical & Recent Dividends and Splits
            act = t.actions
            if act is not None and not act.empty:
                for dt, row in act.tail(12).iterrows():
                    d = dt.date() if hasattr(dt, 'date') else dt
                    if not isinstance(d, date):
                        continue

                    div = row.get('Dividends', 0)
                    split = row.get('Stock Splits', 0)

                    if div and float(div) > 0:
                        key = ('Dividend', str(d))
                        existing = next((e for e in events if e.get('event_type') == 'Dividend' and str(e.get('raw_date')) == str(d)), None)
                        if existing:
                            existing['details'] = f"₹{float(div):.2f} per share dividend"
                        elif key not in seen:
                            seen.add(key)
                            events.append({
                                'event_type': 'Dividend',
                                'raw_date': d,
                                'event_date': d.strftime('%d %b %Y'),
                                'details': f"₹{float(div):.2f} per share dividend",
                                'is_upcoming': d >= today
                            })

                    if split and float(split) > 0:
                        key = ('Stock Split', str(d))
                        if key not in seen:
                            seen.add(key)
                            events.append({
                                'event_type': 'Stock Split',
                                'raw_date': d,
                                'event_date': d.strftime('%d %b %Y'),
                                'details': f"Stock Split (Ratio {split:.0f}:1)",
                                'is_upcoming': d >= today
                            })

        except Exception as e:
            print(f"[Events Engine] yfinance note for {clean_ticker}: {e}")

    # 2. Scrape Screener.in announcements for Board Meetings & Filings
    try:
        url = f"https://www.screener.in/company/{clean_ticker}/consolidated/"
        resp = requests.get(
            url,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'},
            timeout=4
        )
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            doc_section = soup.find('section', id='documents') or soup.find('div', id='announcements')
            if doc_section:
                items = doc_section.find_all('li')
                for li in items[:15]:
                    text = li.get_text(' ', strip=True)
                    text_lower = text.lower()

                    ev_type = None
                    if 'board meeting' in text_lower:
                        ev_type = 'Board Meeting'
                    elif 'financial results' in text_lower or 'quarterly' in text_lower:
                        ev_type = 'Earnings'
                    elif 'dividend' in text_lower:
                        ev_type = 'Dividend'
                    elif 'bonus' in text_lower:
                        ev_type = 'Bonus Issue'
                    elif 'split' in text_lower:
                        ev_type = 'Stock Split'
                    elif 'general meeting' in text_lower or 'agm' in text_lower:
                        ev_type = 'AGM / EGM'

                    if ev_type:
                        date_tag = li.find(['span', 'div'], class_=lambda c: c and 'date' in c) or li.find('span')
                        date_str = date_tag.get_text(strip=True) if date_tag else ''

                        details = text
                        if date_str and details.endswith(date_str):
                            details = details[:-len(date_str)].strip()
                        details = details.split(' - ')[0].strip()
                        if len(details) > 80:
                            details = details[:77] + '...'

                        key = (ev_type, details[:30])
                        if key not in seen:
                            seen.add(key)
                            events.append({
                                'event_type': ev_type,
                                'raw_date': today,
                                'event_date': date_str or today.strftime('%d %b %Y'),
                                'details': details,
                                'is_upcoming': False
                            })
    except Exception as e:
        print(f"[Events Engine] Screener announcements note for {clean_ticker}: {e}")

    # 3. Sort chronologically (upcoming and newest first)
    def _sort_key(ev):
        d = ev.get('raw_date')
        if isinstance(d, datetime):
            return d.date()
        if isinstance(d, date):
            return d
        return date.min

    events.sort(key=_sort_key, reverse=True)

    # Clean up raw_date before returning
    for ev in events:
        ev.pop('raw_date', None)

    return events[:max_events]
