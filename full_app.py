import pandas as pd
import requests
import datetime
import warnings
import re
import glob
import os
from dotenv import load_dotenv
from openai import OpenAI
from bs4 import BeautifulSoup
load_dotenv()
warnings.simplefilter(action="ignore", category=pd.errors.SettingWithCopyWarning)

# ===============================================
# CONFIGURATION
# ===============================================
API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_MODEL = "deepseek-chat"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

MONTHLY_INVESTMENT_EGP = 10000  # ← Change to your monthly budget
WATCHLIST_SYMBOLS = ["MHOT","ORHD"]    # ← Stocks you don't own but want analyzed/predicted

# Recommendation engine config
RECOMMEND_NEW_STOCKS    = True   # ← Set False to skip the recommendation engine
PREFER_DIVERSIFICATION  = True   # ← True = avoid sectors you already own (allow overlap only for exceptional picks)
NUM_MEDIUM_TERM_PICKS   = 3      # ← 1-3 month horizon
NUM_LONG_TERM_PICKS     = 3      # ← 6m-3y horizon

# ===============================================
# LOAD PORTFOLIO FROM EXCEL
# ===============================================
def load_portfolio(excel_path):
    df = pd.read_excel(excel_path)
    portfolio = {
        "stocks": {},
        "mutual_funds": {},
        "summary": {"total_value": 0, "total_return": 0}
    }
    # Tolerant numeric parser. Handles:
    #   - US format: "1,234.56"
    #   - EU format: "1.234,56"
    #   - Arabic-Indic digits: "٨٩٤٢" or "۸۹۴۲"
    #   - Invisible RTL/LTR marks (Thndr is an Arabic-locale app)
    #   - Currency symbols, NBSP, regular spaces as thousand separators
    #   - NaN, None, "—", "N/A"
    _AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    def _to_float(v, default=0.0):
        if v is None: return default
        if isinstance(v, (int, float)):
            try:
                f = float(v)
                return default if (f != f) else f  # NaN
            except Exception:
                return default
        s = str(v)
        # Strip invisible direction/format marks (RLM/LRM/ZWJ/ZWNJ/BOM/etc.)
        s = re.sub(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]", "", s)
        # Convert Arabic-Indic digits to ASCII
        s = s.translate(_AR_DIGITS)
        s = s.strip()
        if not s or s.lower() in ("nan", "none", "n/a", "-", "—"): return default
        # Normalize NBSP and other spaces to regular space
        s = s.replace("\u00a0", " ").replace("\u2009", " ").replace("\u202f", " ")
        # Strip everything except digits, comma, dot, minus, plus, space
        s = re.sub(r"[^\d,.\-+ ]", "", s)
        # Treat space as thousands separator (e.g. "8 942")
        s = s.replace(" ", "")
        if not s: return default
        last_dot = s.rfind(".")
        last_com = s.rfind(",")
        if last_dot >= 0 and last_com >= 0:
            if last_com > last_dot:   # EU: 1.234,56
                s = s.replace(".", "").replace(",", ".")
            else:                     # US: 1,234.56
                s = s.replace(",", "")
        elif last_com >= 0:
            # Only comma. Decide: decimal (EU) vs thousands (US).
            # Rule: if the comma is followed by exactly 3 digits AND there is no
            # other comma, it is ambiguous — but in financial data comma+3 digits
            # is almost always thousands ("1,234"). Comma+1or2 digits is decimal.
            parts = s.split(",")
            if len(parts) == 2 and 1 <= len(parts[1]) <= 2:
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return default

    for _, row in df.iterrows():
        symbol      = str(row['Asset']).strip()
        asset_class = str(row['Asset Class']).strip()
        units       = _to_float(row['Units Owned'])
        cost        = _to_float(row['Cost Per Unit'])
        current     = _to_float(row['Current Price'])
        mkt_val     = _to_float(row['Market Value'])

        ret_str   = str(row['Unrealized Return'])
        val_match = re.match(r'([+-]?[\d,.]+)', ret_str)
        pct_match = re.search(r'\(([\d.]+)%\)', ret_str)
        ret_egp   = _to_float(val_match.group(1)) if val_match else 0.0
        ret_pct   = _to_float(pct_match.group(1)) if pct_match else 0.0
        if ret_str.strip().startswith('-'):
            ret_egp = -abs(ret_egp)
            ret_pct = -abs(ret_pct)

        asset_data = {
            "symbol":        symbol,
            "class":         asset_class,
            "units":         units,
            "cost_price":    cost,
            "current_price": current,
            "market_value":  mkt_val,
            "return_pct":    ret_pct,
            "return_egp":    ret_egp,
            "sector":        "—",  # filled from TradingView after fetch
        }
        if asset_class.lower() == "stock":
            portfolio["stocks"][symbol] = asset_data
        else:
            portfolio["mutual_funds"][symbol] = asset_data

        portfolio["summary"]["total_value"]  += mkt_val
        portfolio["summary"]["total_return"] += ret_egp

    return portfolio

# ===============================================
# TRADINGVIEW SCANNER API — 30 verified columns
# ===============================================
def fetch_tradingview_data(symbols):
    print("📊 [1/3] Fetching LIVE technical data from TradingView Scanner API...")

    tickers = [f"EGX:{s}" for s in symbols]
    url = "https://scanner.tradingview.com/egypt/scan"
    tv_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Accept-Language': 'en-US,en;q=0.9',
        'Content-Type': 'application/json',
        'Origin': 'https://www.tradingview.com',
        'Referer': 'https://www.tradingview.com/',
    }

    COLUMNS = [
        "Recommend.All",                   # [0]  overall signal
        "Recommend.MA",                    # [1]  MA signal
        "Recommend.Other",                 # [2]  oscillators signal
        "RSI",                             # [3]  RSI(14)
        "RSI[1]",                          # [4]  RSI prev bar
        "close",                           # [5]  last price
        "change",                          # [6]  % change today
        "volume",                          # [7]  today volume
        "price_earnings_ttm",              # [8]  P/E (TTM)
        "earnings_per_share_diluted_ttm",  # [9]  EPS (diluted TTM)
        "High.1M",                         # [10] 1M high
        "Low.1M",                          # [11] 1M low
        "High.3M",                         # [12] 3M high
        "Low.3M",                          # [13] 3M low
        "SMA20",                           # [14] 20d MA
        "SMA50",                           # [15] 50d MA
        "SMA200",                          # [16] 200d MA
        "Mom",                             # [17] Momentum
        "MACD.macd",                       # [18] MACD line
        "MACD.signal",                     # [19] MACD signal
        "BB.upper",                        # [20] BB upper
        "BB.lower",                        # [21] BB lower
        "Stoch.K",                         # [22] Stoch %K
        "Stoch.D",                         # [23] Stoch %D
        "ATR",                             # [24] ATR volatility
        "average_volume_10d_calc",         # [25] 10d avg volume
        "price_book_ratio",                # [26] P/B
        "dividends_yield",                 # [27] Dividend yield
        "debt_to_equity",                  # [28] D/E ratio
        "return_on_equity",                # [29] ROE (TTM)
        "sector",                          # [30] Sector
        "industry",                        # [31] Industry
        "description",                     # [32] Full company name
        # ── FALLBACK COLUMNS (try if primary returned null) ──
        "earnings_per_share_basic_ttm",    # [33] EPS (basic TTM) — fallback for diluted
        "earnings_per_share_fq",           # [34] EPS most recent quarter (annualize)
        "return_on_equity_fq",             # [35] ROE most recent quarter — fallback for TTM
        "return_on_invested_capital",      # [36] ROIC — secondary quality metric
        "dividend_payout_ratio_ttm",       # [37] Payout ratio (sanity check for div yield)
    ]

    payload = {"symbols": {"tickers": tickers, "query": {"types": []}}, "columns": COLUMNS}

    def sig(score):
        if score is None:  return "N/A"
        if score >= 0.5:   return "STRONG BUY"
        if score >= 0.1:   return "BUY"
        if score > -0.1:   return "NEUTRAL"
        if score > -0.5:   return "SELL"
        return "STRONG SELL"

    def fmt(v, n=2): return round(v, n) if v is not None else "N/A"

    results = {}
    try:
        resp = requests.post(url, json=payload, headers=tv_headers, timeout=15)
        if resp.status_code == 200:
            for item in resp.json().get("data", []):
                ticker = item["s"].replace("EGX:", "")
                d = item["d"]
                close, sma50, sma200 = d[5], d[15], d[16]
                high_3m, low_3m     = d[12], d[13]
                macd, macd_s        = d[18], d[19]
                bb_up, bb_lo        = d[20], d[21]
                avg_vol             = d[25]

                r3m   = (high_3m - low_3m) if (high_3m and low_3m and high_3m != low_3m) else None
                rpos  = round((close - low_3m) / r3m * 100, 1) if (r3m and close) else None
                mxing = ("BULLISH" if macd > macd_s else "BEARISH") if (macd is not None and macd_s is not None) else "N/A"
                bbr   = (bb_up - bb_lo) if (bb_up and bb_lo) else None
                bbpos = round((close - bb_lo) / bbr * 100, 1) if (bbr and close) else "N/A"
                volr  = round(d[7] / avg_vol, 2) if (avg_vol and d[7]) else "N/A"

                sector_str   = f"{d[30]} / {d[31]}" if (d[30] and d[31]) else (d[30] or d[31] or "Unknown")
                company_name = d[32] if d[32] else ticker

                # ── FUNDAMENTALS WITH 3-LAYER FALLBACK ──
                # Layer 1: TradingView primary (TTM)
                pe_raw   = d[8]    # price_earnings_ttm
                eps_raw  = d[9]    # earnings_per_share_diluted_ttm
                roe_raw  = d[29]   # return_on_equity (TTM)

                # Layer 2: TradingView fallback columns
                eps_basic_ttm = d[33] if len(d) > 33 else None  # earnings_per_share_basic_ttm
                eps_fq        = d[34] if len(d) > 34 else None  # earnings_per_share_fq
                roe_fq        = d[35] if len(d) > 35 else None  # return_on_equity_fq
                pe_source = "TTM (diluted)"
                eps_source = "TTM (diluted)"
                roe_source = "TTM"

                # EPS fallback: diluted TTM → basic TTM → quarterly × 4
                eps_final = eps_raw
                if eps_final is None and eps_basic_ttm is not None:
                    eps_final = eps_basic_ttm
                    eps_source = "TTM (basic)"
                elif eps_final is None and eps_fq is not None:
                    eps_final = eps_fq * 4  # annualize quarterly
                    eps_source = "FQ × 4 (estimated)"

                # P/E fallback: TTM column → calculate from price ÷ EPS
                pe_final = pe_raw
                if pe_final is None and eps_final is not None and eps_final > 0 and close:
                    pe_final = close / eps_final
                    pe_source = f"calc: price/{eps_source}"
                elif pe_final is None and eps_final is not None and eps_final <= 0:
                    pe_source = "N/A (negative earnings)"

                # ROE fallback: TTM → most recent quarter
                roe_final = roe_raw
                if roe_final is None and roe_fq is not None:
                    roe_final = roe_fq
                    roe_source = "FQ (latest quarter)"

                results[ticker] = {
                    "overall":           sig(d[0]),
                    "moving_averages":   sig(d[1]),
                    "oscillators":       sig(d[2]),
                    "rsi":               fmt(d[3], 1),
                    "rsi_prev":          fmt(d[4], 1),
                    "close":             fmt(close),
                    "change_pct":        fmt(d[6]),
                    "macd_cross":        mxing,
                    "macd_val":          fmt(d[18], 3),
                    "macd_sig":          fmt(d[19], 3),
                    "stoch_k":           fmt(d[22], 1),
                    "stoch_d":           fmt(d[23], 1),
                    "atr":               fmt(d[24]),
                    "bb_position":       bbpos,
                    "vol_today":         int(d[7]) if d[7] else "N/A",
                    "vol_ratio":         volr,
                    "sma20":             fmt(d[14]),
                    "sma50":             fmt(sma50),
                    "sma200":            fmt(sma200),
                    "above_sma20":       (close > d[14]) if (close and d[14]) else None,
                    "above_sma50":       (close > sma50) if (close and sma50) else None,
                    "above_sma200":      (close > sma200) if (close and sma200) else None,
                    "high_1m":           fmt(d[10]),
                    "low_1m":            fmt(d[11]),
                    "high_3m":           fmt(high_3m),
                    "low_3m":            fmt(low_3m),
                    "range_3m_position": rpos,
                    "momentum":          fmt(d[17]),
                    "pe":                fmt(pe_final, 1),
                    "pe_source":         pe_source if pe_final is not None else "N/A",
                    "eps":               fmt(eps_final, 2),
                    "eps_source":        eps_source if eps_final is not None else "N/A",
                    "pb":                fmt(d[26], 1),
                    "div_yield":         fmt(d[27], 2),
                    "debt_equity":       fmt(d[28], 2),
                    "roe":               fmt(roe_final, 1),
                    "roe_source":        roe_source if roe_final is not None else "N/A",
                    "roic":              fmt(d[36] if len(d) > 36 else None, 1),
                    "sector":            sector_str,
                    "company_name":      company_name,
                }
                tv = results[ticker]
                pe_disp  = f"{tv['pe']} ({tv['pe_source']})" if tv['pe'] != "N/A" else "N/A"
                roe_disp = f"{tv['roe']}% ({tv['roe_source']})" if tv['roe'] != "N/A" else "N/A"
                print(f"   ✅ {ticker} [{sector_str}]: {tv['overall']} | RSI={tv['rsi']} | MACD={tv['macd_cross']} | P/E={pe_disp} | ROE={roe_disp}")
        else:
            print(f"   ⚠️ TradingView HTTP {resp.status_code}: {resp.text[:120]}")
    except Exception as e:
        print(f"   ⚠️ TradingView error: {e}")

    empty = {k: "N/A" for k in ["overall","moving_averages","oscillators","rsi","rsi_prev",
             "close","change_pct","macd_cross","macd_val","macd_sig","stoch_k","stoch_d",
             "atr","bb_position","vol_today","vol_ratio","sma20","sma50","sma200",
             "above_sma20","above_sma50","above_sma200","high_1m","low_1m","high_3m","low_3m",
             "range_3m_position","momentum","pe","pe_source","eps","eps_source","pb",
             "div_yield","debt_equity","roe","roe_source","roic",
             "sector","company_name"]}
    for s in symbols:
        if s not in results:
            results[s] = dict(empty)
            print(f"   ⚠️ {s}: No data returned")

    # ── LAYER 3 FALLBACK: Scrape stockanalysis.com for stocks still missing P/E or ROE ──
    missing_fundamentals = [s for s in symbols if results.get(s, {}).get("pe","N/A") == "N/A"
                                                or results.get(s, {}).get("roe","N/A") == "N/A"]
    if missing_fundamentals:
        print(f"   🔍 Trying stockanalysis.com fallback for: {missing_fundamentals}")
        scraped = _scrape_investing_fundamentals(missing_fundamentals)
        for sym, fund in scraped.items():
            if sym in results:
                source = fund.get("_roe_source", "stockanalysis.com")
                if results[sym].get("pe","N/A") == "N/A" and fund.get("pe") is not None:
                    results[sym]["pe"] = round(fund["pe"], 1)
                    results[sym]["pe_source"] = "stockanalysis.com"
                    print(f"      ✅ {sym} P/E filled: {fund['pe']:.1f}")
                if results[sym].get("roe","N/A") == "N/A" and fund.get("roe") is not None:
                    results[sym]["roe"] = round(fund["roe"], 1)
                    results[sym]["roe_source"] = source
                    print(f"      ✅ {sym} ROE filled from {source}: {fund['roe']:.1f}%")
                if results[sym].get("eps","N/A") == "N/A" and fund.get("eps") is not None:
                    results[sym]["eps"] = round(fund["eps"], 2)
                    results[sym]["eps_source"] = "stockanalysis.com"
                if results[sym].get("div_yield","N/A") == "N/A" and fund.get("div_yield") is not None:
                    results[sym]["div_yield"] = round(fund["div_yield"], 2)

    # ── LAYER 4 FALLBACK: Mathematical ROE derivation when all scrapes failed ──
    # Identity: ROE = (1/PE) × PB × 100 — if we have P/E and P/B, we can compute ROE.
    for sym in symbols:
        r = results.get(sym, {})
        if r.get("roe", "N/A") == "N/A":
            try:
                pe = float(r.get("pe", "N/A")) if r.get("pe", "N/A") != "N/A" else None
                pb = float(r.get("pb", "N/A")) if r.get("pb", "N/A") != "N/A" else None
                if pe is not None and pb is not None and pe > 0 and pb > 0:
                    roe_calc = (1.0 / pe) * pb * 100
                    if 0 < roe_calc < 200:  # sanity bound
                        results[sym]["roe"] = round(roe_calc, 1)
                        results[sym]["roe_source"] = "calc: PB/PE"
                        print(f"      ✅ {sym} ROE calculated from P/E & P/B: {roe_calc:.1f}%")
            except Exception:
                pass

    return results


# ===============================================
# FUNDAMENTALS FALLBACK — stockanalysis.com (clean, no Cloudflare)
# ===============================================
def _scrape_investing_fundamentals(symbols):
    """Scrape P/E, EPS, dividend yield from stockanalysis.com for EGX symbols.
    URL pattern: https://stockanalysis.com/quote/egx/{SYMBOL}/
    Note: stockanalysis.com displays EPS, P/E, dividend, market cap on the main page.
    For ROE we calculate from Net Income / Equity if needed (rare — most stocks have it on TradingView).
    Returns dict {symbol: {pe, eps, roe, div_yield}}."""
    results = {}
    sa_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Upgrade-Insecure-Requests': '1',
    }

    def parse_num(text):
        """Parse '4.3', '32.1%', '17.14', '6.55' → float. Returns None if can't parse."""
        if text is None: return None
        try:
            s = str(text).strip().replace(',', '').replace('%', '').replace('$', '')
            s = re.sub(r'[^\d\.\-]', '', s)
            if not s or s in ('-', '.', '-.'): return None
            n = float(s)
            return n if -100000 < n < 100000 else None
        except Exception:
            return None

    def find_value_after_label(text, label_patterns, max_chars=80):
        """Find a numeric value appearing right after one of the label patterns.
        Tries each pattern, returns first match."""
        for pat in label_patterns:
            m = re.search(pat + r'[\s:]*([\-]?\d+(?:,\d{3})*(?:\.\d+)?)\s*%?', text, re.IGNORECASE)
            if m:
                val = parse_num(m.group(1))
                if val is not None:
                    return val
        return None

    def scrape_main_page(symbol):
        """Scrape https://stockanalysis.com/quote/egx/{SYMBOL}/ for P/E, EPS, dividend yield.
        Uses direct HTML table parsing — far more reliable than regex on text."""
        url = f"https://stockanalysis.com/quote/egx/{symbol.upper()}/"
        try:
            r = requests.get(url, headers=sa_headers, timeout=15)
            if r.status_code != 200:
                print(f"      ⚠️ {symbol} stockanalysis.com HTTP {r.status_code}")
                return {}

            soup = BeautifulSoup(r.content, 'html.parser')
            data = {"pe": None, "eps": None, "roe": None, "div_yield": None}

            # ── Strategy 1: Walk every <tr> and pair label cell → value cell ──
            # stockanalysis.com renders the stats box as a 2-column table:
            # <tr><td>PE Ratio</td><td>13.29</td></tr>
            # <tr><td>EPS</td><td>6.55</td></tr>
            # <tr><td>Dividend</td><td>6.00 (6.90%)</td></tr>
            for row in soup.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                if len(cells) < 2:
                    continue
                label = cells[0].get_text(" ", strip=True).lower()
                value = cells[1].get_text(" ", strip=True)
                if not label or not value:
                    continue

                # PE Ratio (skip Forward PE)
                if data["pe"] is None and label in ("pe ratio", "p/e ratio", "pe", "p/e"):
                    data["pe"] = parse_num(value)

                # EPS — match "EPS", "EPS (TTM)", "Diluted EPS"
                elif data["eps"] is None and label in ("eps", "eps (ttm)", "diluted eps", "basic eps", "earnings per share"):
                    data["eps"] = parse_num(value)

                # Dividend — value comes as "6.00 (6.90%)" → extract yield from parens
                elif data["div_yield"] is None and label == "dividend":
                    m = re.search(r'\(\s*([\-]?\d+\.?\d*)\s*%\s*\)', value)
                    if m: data["div_yield"] = parse_num(m.group(1))

                # ROE — usually only on /statistics/ page, but check just in case
                elif data["roe"] is None and label in ("return on equity", "return on equity (roe)", "roe"):
                    data["roe"] = parse_num(value)

            # ── Strategy 2: Same logic for non-table layouts (some pages use divs) ──
            # Some EGX pages render as <div><span>Label</span><span>Value</span></div>
            if data["pe"] is None or data["eps"] is None or data["div_yield"] is None:
                for parent in soup.find_all(['div', 'li']):
                    children = [c for c in parent.find_all(['span', 'div', 'p'], recursive=False)]
                    if len(children) != 2:
                        continue
                    label = children[0].get_text(" ", strip=True).lower()
                    value = children[1].get_text(" ", strip=True)
                    if not label or not value:
                        continue
                    if data["pe"] is None and label in ("pe ratio", "p/e ratio"):
                        data["pe"] = parse_num(value)
                    elif data["eps"] is None and label in ("eps", "eps (ttm)"):
                        data["eps"] = parse_num(value)
                    elif data["div_yield"] is None and label == "dividend":
                        m = re.search(r'\(\s*([\-]?\d+\.?\d*)\s*%\s*\)', value)
                        if m: data["div_yield"] = parse_num(m.group(1))

            # ── Strategy 3: Whole-page regex fallback (last resort) ──
            if data["pe"] is None or data["eps"] is None or data["div_yield"] is None:
                full_text = soup.get_text(" | ", strip=True)
                # Match patterns like "PE Ratio | 13.29" or "PE Ratio 13.29"
                if data["pe"] is None:
                    m = re.search(r'PE\s*Ratio\s*\|?\s*([\-]?\d+\.?\d*)', full_text, re.IGNORECASE)
                    if m: data["pe"] = parse_num(m.group(1))
                if data["eps"] is None:
                    # Match EPS surrounded by pipes (table-like structure)
                    m = re.search(r'\|\s*EPS\s*(?:\(ttm\))?\s*\|\s*([\-]?\d+\.?\d*)', full_text, re.IGNORECASE)
                    if m: data["eps"] = parse_num(m.group(1))
                if data["div_yield"] is None:
                    m = re.search(r'Dividend\s*\|?\s*[\d\.]+\s*\(\s*([\-]?\d+\.?\d*)\s*%\s*\)', full_text, re.IGNORECASE)
                    if m: data["div_yield"] = parse_num(m.group(1))

            return data
        except Exception as e:
            print(f"      ⚠️ {symbol} stockanalysis.com error: {e}")
            return {}

    def scrape_financials_for_roe(symbol):
        """Try multiple stockanalysis.com sub-pages for ROE."""
        urls = [
            f"https://stockanalysis.com/quote/egx/{symbol.upper()}/statistics/",
            f"https://stockanalysis.com/quote/egx/{symbol.upper()}/financials/ratios/",
            f"https://stockanalysis.com/quote/egx/{symbol.upper()}/financials/",
        ]
        for url in urls:
            try:
                r = requests.get(url, headers=sa_headers, timeout=15)
                if r.status_code != 200:
                    continue
                soup = BeautifulSoup(r.content, 'html.parser')
                text = soup.get_text(" ", strip=True)
                # ROE row: "Return on Equity (ROE) X.XX%" or "ROE X.XX%"
                m = re.search(r'Return\s+on\s+Equity[^\d\-]{0,40}([\-]?\d+\.?\d*)\s*%', text, re.IGNORECASE)
                if m:
                    val = parse_num(m.group(1))
                    if val is not None: return val
                m = re.search(r'\bROE\b[^\d\-]{0,20}([\-]?\d+\.?\d*)\s*%', text, re.IGNORECASE)
                if m:
                    val = parse_num(m.group(1))
                    if val is not None: return val
            except Exception:
                continue
        return None

    def scrape_mubasher(symbol):
        """Mubasher is the Egyptian financial portal — no Cloudflare, has ROE/PE for many EGX stocks."""
        url = f"https://english.mubasher.info/markets/EGX/stocks/{symbol.upper()}/financials"
        try:
            r = requests.get(url, headers=sa_headers, timeout=15)
            if r.status_code != 200:
                return {}
            soup = BeautifulSoup(r.content, 'html.parser')
            text = soup.get_text(" ", strip=True)
            data = {"pe": None, "eps": None, "roe": None, "div_yield": None}
            # Mubasher uses labels like "Return on Equity" and "P/E Ratio"
            m = re.search(r'Return\s+on\s+Equity[^\d\-]{0,40}([\-]?\d+\.?\d*)\s*%?', text, re.IGNORECASE)
            if m: data["roe"] = parse_num(m.group(1))
            m = re.search(r'P[/\s]*E\s*Ratio[^\d\-]{0,30}([\-]?\d+\.?\d*)', text, re.IGNORECASE)
            if m: data["pe"] = parse_num(m.group(1))
            m = re.search(r'EPS[^\d\-]{0,30}([\-]?\d+\.?\d*)', text, re.IGNORECASE)
            if m: data["eps"] = parse_num(m.group(1))
            return data
        except Exception:
            return {}

    def derive_roe_from_pe_pb(pe, pb):
        """Mathematical identity: ROE = (1/PE) × PB × 100.
        This is exact accounting math — works whenever both PE and PB are available.
        Returns ROE as a percentage."""
        try:
            if pe is None or pb is None: return None
            pe_f = float(pe); pb_f = float(pb)
            if pe_f <= 0 or pb_f <= 0: return None
            return (1.0 / pe_f) * pb_f * 100
        except Exception:
            return None

    for sym in symbols:
        print(f"      🔎 Scraping stockanalysis.com for {sym}...", end=" ", flush=True)
        data = scrape_main_page(sym)

        # ROE: try stockanalysis.com sub-pages
        if data.get("roe") is None:
            roe = scrape_financials_for_roe(sym)
            if roe is not None:
                data["roe"] = roe

        # ROE: try Mubasher (Egyptian portal)
        if data.get("roe") is None or data.get("pe") is None:
            mub = scrape_mubasher(sym)
            for k in ("roe", "pe", "eps", "div_yield"):
                if data.get(k) is None and mub.get(k) is not None:
                    data[k] = mub[k]
                    if k == "roe":
                        data["_roe_source"] = "mubasher.info"

        results[sym] = data
        # Print what we got — visible to user so they know it worked
        got_pe   = f"P/E={data.get('pe')}"   if data.get("pe")   is not None else "P/E=—"
        got_eps  = f"EPS={data.get('eps')}"  if data.get("eps")  is not None else "EPS=—"
        got_roe  = f"ROE={data.get('roe')}%" if data.get("roe")  is not None else "ROE=—"
        got_div  = f"Div={data.get('div_yield')}%" if data.get("div_yield") is not None else "Div=—"
        print(f"{got_pe} | {got_eps} | {got_roe} | {got_div}")

    return results


# ===============================================
# INVESTING.COM — Market context
# ===============================================
def fetch_market_context():
    print("📡 [2/3] Fetching market context from TradingView Scanner...")
    data = {
        "egx30":         {"value": "N/A", "change_pct": "N/A"},
        "egx70":         {"value": "N/A", "change_pct": "N/A"},
        "egx100":        {"value": "N/A", "change_pct": "N/A"},
        "oil":           {"price": "N/A", "change_pct": "N/A"},
        "gold":          {"price": "N/A", "change_pct": "N/A"},
        "usd_egp":       {"rate":  "N/A", "change_pct": "N/A"},
        "market_status": "Unknown",
        "last_updated":  datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    tv_headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json',
        'Accept-Language': 'en-US,en;q=0.9',
        'Content-Type': 'application/json',
        'Origin': 'https://www.tradingview.com',
        'Referer': 'https://www.tradingview.com/',
    }

    def tv_scan(url, payload):
        try:
            r = requests.post(url, json=payload, headers=tv_headers, timeout=15)
            if r.status_code == 200:
                return r.json().get("data", [])
        except Exception as e:
            print(f"   ⚠️ scan error {url}: {e}")
        return []

    # ── 1. EGX Indices via egypt/scan with explicit symbol list ──────
    egx_payload = {
        "symbols": {
            "tickers": [
                "EGX:EGX30", "EGX:EGX30TR",
                "EGX:EGX70EWI", "EGX:EGX70",
                "EGX:EGX100EWI", "EGX:EGX100",
            ],
            "query": {"types": []}
        },
        "columns": ["close", "change"]
    }
    found_indices = {}
    for item in tv_scan("https://scanner.tradingview.com/egypt/scan", egx_payload):
        sym   = item["s"]
        close = item["d"][0]
        chg   = item["d"][1]
        if close is not None:
            found_indices[sym] = (close, chg)

    for key, candidates in [
        ("egx30",  ["EGX:EGX30",     "EGX:EGX30TR"]),
        ("egx70",  ["EGX:EGX70EWI",  "EGX:EGX70"]),
        ("egx100", ["EGX:EGX100EWI", "EGX:EGX100"]),
    ]:
        for c in candidates:
            if c in found_indices:
                close, chg = found_indices[c]
                data[key]["value"]      = f"{close:,.2f}"
                data[key]["change_pct"] = f"{chg:+.2f}%" if chg is not None else "N/A"
                break

    # ── 2. Oil & Gold via cfd/scan (TradingView CFD scanner) ─────────
    cfd_payload = {
        "symbols": {
            "tickers": ["OANDA:XAUUSD", "OANDA:USOIL", "TVC:GOLD", "TVC:USOIL",
                        "BLACKBULL:GOLD", "BLACKBULL:WTI"],
            "query": {"types": []}
        },
        "columns": ["close", "change"]
    }
    for item in tv_scan("https://scanner.tradingview.com/cfd/scan", cfd_payload):
        sym   = item["s"]
        close = item["d"][0]
        chg   = item["d"][1]
        if close is None:
            continue
        sym_upper = sym.upper()
        if "GOLD" in sym_upper or "XAU" in sym_upper:
            if data["gold"]["price"] == "N/A":
                data["gold"]["price"]      = f"{close:,.2f}"
                data["gold"]["change_pct"] = f"{chg:+.2f}%" if chg is not None else "N/A"
        elif "OIL" in sym_upper or "WTI" in sym_upper or "USOIL" in sym_upper:
            if data["oil"]["price"] == "N/A":
                data["oil"]["price"]      = f"{close:,.2f}"
                data["oil"]["change_pct"] = f"{chg:+.2f}%" if chg is not None else "N/A"

    # fallback: try futures/scan for oil & gold
    if data["oil"]["price"] == "N/A" or data["gold"]["price"] == "N/A":
        futures_payload = {
            "symbols": {
                "tickers": ["NYMEX:CL1!", "COMEX:GC1!", "MCX:CRUDEOIL1!", "MCX:GOLD1!"],
                "query": {"types": []}
            },
            "columns": ["close", "change"]
        }
        for item in tv_scan("https://scanner.tradingview.com/futures/scan", futures_payload):
            sym   = item["s"]
            close = item["d"][0]
            chg   = item["d"][1]
            if close is None:
                continue
            if ("CL" in sym or "CRUDE" in sym) and data["oil"]["price"] == "N/A":
                data["oil"]["price"]      = f"{close:,.2f}"
                data["oil"]["change_pct"] = f"{chg:+.2f}%" if chg is not None else "N/A"
            elif ("GC" in sym or "GOLD" in sym) and data["gold"]["price"] == "N/A":
                data["gold"]["price"]      = f"{close:,.2f}"
                data["gold"]["change_pct"] = f"{chg:+.2f}%" if chg is not None else "N/A"

    # ── 3. USD/EGP via forex/scan ────────────────────────────────────
    forex_payload = {
        "symbols": {
            "tickers": [
                "FX:USDEGP", "FX_IDC:USDEGP", "OANDA:USDEGP",
                "FOREXCOM:USDEGP", "FXOPEN:USDEGP"
            ],
            "query": {"types": []}
        },
        "columns": ["close", "change"]
    }
    for item in tv_scan("https://scanner.tradingview.com/forex/scan", forex_payload):
        close = item["d"][0]
        chg   = item["d"][1]
        if close is not None and close > 1:   # sanity check: EGP rate > 1
            data["usd_egp"]["rate"]       = f"{close:.4f}"
            data["usd_egp"]["change_pct"] = f"{chg:+.4f}%" if chg is not None else "N/A"
            break

    # fallback: try global/scan for forex
    if data["usd_egp"]["rate"] == "N/A":
        global_forex_payload = {
            "symbols": {
                "tickers": ["FX:USDEGP", "FX_IDC:USDEGP", "OANDA:USDEGP"],
                "query": {"types": []}
            },
            "columns": ["close", "change"]
        }
        for item in tv_scan("https://scanner.tradingview.com/global/scan", global_forex_payload):
            close = item["d"][0]
            chg   = item["d"][1]
            if close is not None and close > 1:
                data["usd_egp"]["rate"]       = f"{close:.4f}"
                data["usd_egp"]["change_pct"] = f"{chg:+.4f}%" if chg is not None else "N/A"
                break

    data["market_status"] = (
        "Weekend — Market Closed"
        if datetime.date.today().weekday() in [4, 5]
        else "Trading Day"
    )

    print(
        f"   ✅ EGX30={data['egx30']['value']} ({data['egx30']['change_pct']}) | "
        f"EGX70={data['egx70']['value']} ({data['egx70']['change_pct']}) | "
        f"EGX100={data['egx100']['value']} ({data['egx100']['change_pct']})"
    )
    print(
        f"   ✅ Oil=${data['oil']['price']} ({data['oil']['change_pct']}) | "
        f"Gold=${data['gold']['price']} ({data['gold']['change_pct']}) | "
        f"USD/EGP={data['usd_egp']['rate']} ({data['usd_egp']['change_pct']})"
    )
    return data
 
 
 

# ===============================================
# BUSINESS NEWS
# ===============================================
def fetch_news():
    print("📰 [3/3] Fetching latest Egyptian business news...")
    news = []
    try:
        r = requests.get("https://www.dailynewsegypt.com/category/business/", headers=HEADERS, timeout=10)
        if r.status_code == 200:
            soup = BeautifulSoup(r.content, 'html.parser')
            for h in soup.find_all('h3', class_='entry-title')[:5]:
                a = h.find('a')
                if a: news.append(f"[DNE] {a.text.strip()}")
    except: pass
    try:
        r = requests.get("https://enterprise.press/", headers=HEADERS, timeout=10)
        if r.status_code == 200:
            soup = BeautifulSoup(r.content, 'html.parser')
            for h in soup.find_all(['h2','h3'], class_=re.compile(r'title|headline'))[:4]:
                t = h.text.strip()
                if len(t) > 15: news.append(f"[Enterprise] {t}")
    except: pass
    if not news: news = ["No news scraped — check internet connection"]
    for i, n in enumerate(news[:5], 1): print(f"   {i}. {n[:80]}")
    return news[:6]

# ===============================================
# DEEPSEEK ANALYSIS — Full data, detailed prompt
# ===============================================
def deepseek_analysis(portfolio, market, tv_data, news, monthly_egp):
    print("\n🤖 Running DeepSeek analysis (full data mode)...")

    today     = datetime.date.today()
    total_val = portfolio["summary"]["total_value"]
    total_ret = portfolio["summary"]["total_return"]
    sv        = sum(s["market_value"] for s in portfolio["stocks"].values())
    mv        = sum(m["market_value"] for m in portfolio["mutual_funds"].values())

    # Full detailed stock lines
    stock_lines = []
    for sym, d in portfolio["stocks"].items():
        tv = tv_data.get(sym, {})
        a20  = "↑SMA20"  if tv.get("above_sma20")  else "↓SMA20"
        a50  = "↑SMA50"  if tv.get("above_sma50")  else "↓SMA50"
        a200 = "↑SMA200" if tv.get("above_sma200") else "↓SMA200"
        stock_lines.append(
            f"\n{sym} | Sector: {tv.get('sector', d.get('sector','—'))} | Company: {tv.get('company_name', sym)} | {d['units']:.0f} units\n"
            f"  Position: cost={d['cost_price']:.2f} → price={d['current_price']:.2f} EGP | "
            f"return={d['return_pct']:+.1f}% ({d['return_egp']:+.0f} EGP) | mkt_val={d['market_value']:,.0f} EGP\n"
            f"  Signal: {tv.get('overall','N/A')} | MA_signal={tv.get('moving_averages','N/A')} | Osc_signal={tv.get('oscillators','N/A')}\n"
            f"  Momentum: RSI={tv.get('rsi','N/A')} (prev={tv.get('rsi_prev','N/A')}) | "
            f"MACD={tv.get('macd_cross','N/A')} (macd={tv.get('macd_val','N/A')} sig={tv.get('macd_sig','N/A')}) | "
            f"Stoch_K={tv.get('stoch_k','N/A')} Stoch_D={tv.get('stoch_d','N/A')} | Mom={tv.get('momentum','N/A')}\n"
            f"  Trend: {a20} {a50} {a200} | SMA20={tv.get('sma20','N/A')} SMA50={tv.get('sma50','N/A')} SMA200={tv.get('sma200','N/A')}\n"
            f"  Range: 3M_low={tv.get('low_3m','N/A')} → 3M_high={tv.get('high_3m','N/A')} | "
            f"position_in_3M_range={tv.get('range_3m_position','N/A')}% | "
            f"1M_low={tv.get('low_1m','N/A')} 1M_high={tv.get('high_1m','N/A')}\n"
            f"  Volatility: ATR={tv.get('atr','N/A')} | BB_position={tv.get('bb_position','N/A')}% | "
            f"vol_today={tv.get('vol_today','N/A')} vol_vs_10d_avg=x{tv.get('vol_ratio','N/A')}\n"
            f"  Fundamentals: PE={tv.get('pe','N/A')} | PB={tv.get('pb','N/A')} | "
            f"EPS={tv.get('eps','N/A')} | ROE={tv.get('roe','N/A')}% | "
            f"Div_yield={tv.get('div_yield','N/A')}% | D/E={tv.get('debt_equity','N/A')}"
        )

    mf_lines = []
    for sym, d in portfolio["mutual_funds"].items():
        mf_lines.append(
            f"{sym} [{d['sector']}]: val={d['market_value']:,.0f}EGP | "
            f"cost/unit={d['cost_price']:.2f} → {d['current_price']:.2f} | ret={d['return_pct']:+.1f}%"
        )

    # Sector concentration
    from collections import defaultdict
    sector_vals = defaultdict(float)
    for s, d in portfolio["stocks"].items():
        sector_vals[d["sector"]] += d["market_value"]
    sector_str = " | ".join(f"{k}: {v:,.0f}EGP ({v/sv*100:.0f}%)" for k,v in sorted(sector_vals.items(), key=lambda x: -x[1]))

    prompt = f"""You are a senior Egyptian equity analyst providing a detailed portfolio review.
Date: {today} | Monthly DCA budget: {monthly_egp:,} EGP | Investment horizon: 6-12 months

═══════════════════════════════
MARKET CONDITIONS
═══════════════════════════════
EGX30: {market['egx30']['value']} ({market['egx30']['change_pct']})
EGX70: {market['egx70']['value']} ({market['egx70']['change_pct']})
EGX100: {market['egx100']['value']} ({market['egx100']['change_pct']})
WTI Oil: ${market['oil']['price']} ({market['oil']['change_pct']})
Gold: ${market['gold']['price']} ({market['gold']['change_pct']})
USD/EGP: {market['usd_egp']['rate']} ({market['usd_egp']['change_pct']})
Market: {market['market_status']}

Latest News:
{chr(10).join(f"- {n}" for n in news[:5])}

═══════════════════════════════
PORTFOLIO OVERVIEW
═══════════════════════════════
Total Value: {total_val:,.0f} EGP
Total Return: {total_ret:+,.0f} EGP ({total_ret/total_val*100:+.2f}%)
Stocks: {sv:,.0f} EGP ({sv/total_val*100:.0f}%)
Mutual Funds: {mv:,.0f} EGP ({mv/total_val*100:.0f}%)

SECTOR CONCENTRATION (Stocks only):
{sector_str}

═══════════════════════════════
STOCK POSITIONS — FULL DATA
═══════════════════════════════
{"".join(stock_lines)}

═══════════════════════════════
MUTUAL FUNDS
═══════════════════════════════
{chr(10).join(mf_lines)}

═══════════════════════════════
ANALYSIS REQUIRED
═══════════════════════════════

## 1. MARKET & MACRO IMPACT
Analyze how each macro factor (EGX30 level, oil price, gold, USD/EGP rate) specifically
impacts each sector in this portfolio. Be sector-specific: how does oil price affect ABUK vs AMOC?
How does USD/EGP affect pharma (ISPH, PHAR) vs banking (COMI)? 3-5 sentences.

## 2. SECTOR CONCENTRATION RISK
Comment on the sector weights above. Is the portfolio over-exposed to any sector?
What does this mean for the next 6 months given current macro? What sectors should
be increased or reduced?

## 3. STOCK-BY-STOCK DEEP ANALYSIS
For EACH stock, provide a structured analysis covering ALL of the following:
a) Technical picture: RSI trend (rising/falling vs prev bar), MACD crossover direction,
   Stochastic position, where price sits in the 3M range, above/below key MAs
b) Volatility context: ATR relative to price, BB position (squeezing or expanding?)
c) Volume analysis: Is today's volume above or below 10d average? What does that signal?
d) Fundamental view: Is PE cheap/expensive for this sector? ROE quality? Dividend support?
   Debt risk (D/E)?
e) 6-12 month price target: Give a realistic target range based on fundamentals + technicals
f) Verdict with conviction level: STRONG ACCUMULATE / ACCUMULATE / HOLD / REDUCE / EXIT
   and WHY in one sentence

## 4. MUTUAL FUND ASSESSMENT
Are the current funds appropriate for a 6-12M DCA strategy in the current EGP rate environment?
Should any fund allocation shift between money market, fixed income, or equity given current conditions?

## 5. DCA PLAN — {monthly_egp:,} EGP THIS MONTH
This is the most actionable section. Provide:
- Exact allocation: e.g. "1,200 EGP → COMI, 800 EGP → ORAS"
- For each pick: WHY this month specifically (technical entry point? fundamental catalyst?)
- Entry price target: what price to buy at or below
- Stop-loss level: at what price you were wrong
- 6M price target: where you expect it to be
- What NOT to buy this month and why

## 6. PORTFOLIO RISKS — NEXT 6 MONTHS
List the top 5 risks with probability estimate (Low/Medium/High) and which holdings are most exposed.

## 7. NEXT TRADING SESSION ACTION PLAN
5 specific, numbered, executable actions in priority order.
Include: what to monitor before opening, what orders to place, what triggers to watch.

Use exact numbers throughout. Reference specific data points from the analysis above.
Write as if advising a real client with real money."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "You are a senior Egyptian equity analyst with 15 years experience on the EGX. You provide detailed, specific, data-driven analysis referencing exact numbers. You know EGX-listed companies, their sectors, and Egyptian macroeconomic context deeply.\n\nSTUDY THESE REAL EGX ANALYSIS EXAMPLES — match this depth and format exactly:\n\n══════════════════════════════════════════════════\nEXAMPLE 1 — BULL MARKET / REFORM RALLY (Feb 2026)\nContext: EGX30 at all-time high ~52,800. Egypt closed the $35bn Ras El Hekma deal with ADQ. EGP stabilized post-devaluation. Tourism hit record 19M visitors. EGX was best-performing MENA market up 70% YTD at peak.\n══════════════════════════════════════════════════\n[COMI — Finance / Regional Banks]\nTechnical: RSI 61 rising from 54. MACD bullish crossover confirmed 5 sessions ago. Price at 78% of 3M range (104-145 EGP). Above SMA20 (124), SMA50 (129), SMA200 (103) — full bull alignment. BB position 82% — approaching upper band, watch for consolidation. Volume x1.4 vs 10d avg = institutional accumulation.\nFundamental: P/E 7.2x vs Egyptian banking sector avg 10-11x = 30% discount despite superior quality. ROE 32% = highest in sector. EPS beat last quarter by 17% (actual 5.50 vs estimate 4.70 EGP). Analysts raised 12M target to EGP 160 (Mar 2026). Dividend yield 5.8%. D/E 0.17 = fortress balance sheet. Revenue grew 12% YoY to 128.54bn EGP in 2025.\nMacro: High CBE rates (27.25%) directly boost CIB net interest margin. USD/EGP stability removes FX provisioning pressure. Suez Canal recovery adds sovereign tailwind.\n6M target: 148-160 EGP (+14% to +24%). Stop loss: 118 EGP (below SMA200).\nVerdict: STRONG ACCUMULATE — cheapest high-quality bank on EGX. Earnings momentum + rate tailwind + 30% valuation discount = rare combination.\n\n[TMGH — Real Estate / Development]\nTechnical: RSI 58, MACD bullish, price at 65% of 3M range. Above all MAs. Volume x1.2 avg.\nFundamental: EFG Hermes top pick. Contracted sales backlog EGP 364bn (5-year revenue visibility). Revenue to rise 56% YoY in 2026 to EGP 41.2bn. Expanding to Iraq and Oman. Land bank 125.9M sqm. P/E 18x — premium justified by backlog quality.\n6M target: 95-110 EGP. Stop loss: 70 EGP.\nVerdict: ACCUMULATE — strong operational performance not yet priced in. Regional expansion catalysts ahead.\n\n══════════════════════════════════════════════════\nEXAMPLE 2 — GEOPOLITICAL SHOCK (Oct 7-8, 2023)\nContext: Hamas attack on Israel Oct 7 2023. EGX30 fell 2% on Oct 8. Gulf markets fell 1.5-2%. Brent oil volatile at $84. Egypt under pressure: tourism threat, FX crisis ongoing, USD/EGP black market premium widening. Sentiment: extreme fear.\n══════════════════════════════════════════════════\n[COMI — Finance / Banking]\nTechnical: RSI dropped from 55 to 44 in one session. MACD bearish crossover triggered. Price broke below SMA20. BB position collapsed to 18% — oversold. Volume x2.8 avg = panic selling, not fundamentals.\nFundamental: CIB core business entirely unaffected by conflict. NIM protected by high rates. FX risk: if Egypt tourism/Suez revenue falls, CBE may devalue — temporarily compresses bank P/E multiples. P/E 6.1x (extreme value). ROE 28%. No direct Gaza exposure.\nRisk assessment: Tourism headwind (Sharm El Sheikh bookings down). Suez Canal risk if conflict spreads (Houthis not yet active in Oct 2023). This is sentiment risk, not fundamental impairment.\n6M target: Hold through volatility. 12M target: 95-110 EGP as geopolitical premium fades.\nVerdict: HOLD — do not panic sell a bank at P/E 6x and ROE 28% into a geopolitical shock. Add in tranches if RSI falls below 35. Stop loss: 52 EGP.\n\nDCA strategy during geopolitical shock: Do NOT deploy full monthly budget in week 1. Split into 3 tranches over 3 weeks. Prioritize defensive positions (banking, money market funds) over cyclicals. Keep 30% cash as dry powder for a potential second leg down. Geopolitical discounts in Egypt typically recover within 8-12 weeks if conflict does not directly involve Egypt.\n\n══════════════════════════════════════════════════\nEXAMPLE 3 — EGP DEVALUATION / HIGH INFLATION (Mar 2024)\nContext: CBE devalued EGP 38% on Mar 6 2024 (from 30 to 49/USD). IMF deal expanded to $8bn. Rates hiked 600bps to 27.25% in one session. EGX30 surged 6% on devaluation day then volatile. Inflation at 35%+. Suez Canal revenues collapsing from Houthi attacks.\n══════════════════════════════════════════════════\n[COMI — Finance / Banking]\nTechnical: Post-devaluation RSI spiked to 72 (overbought day 1), settled to 58 over 2 weeks. MACD strongly bullish. Price surged from 45 to 68 EGP (+51%) in one month.\nFundamental: Devaluation is POSITIVE for CIB: FX revaluation gains on USD assets, NIM expansion from rate hike, EGP loan book repriced higher. P/E re-rated from 5x to 7x as EPS surged. Analysts revised 12M target from 55 to 90 EGP within 60 days of devaluation.\nDCA tip: Rushed buyers in week 1 pay premium. Better entry in weeks 2-4 after initial euphoria fades and RSI normalizes to 55-60.\nVerdict: STRONG ACCUMULATE — banking stocks are the best inflation hedge in Egypt after devaluation. Real return vs T-bills at 27% nominal (negative real) makes equities the only viable asset class.\n\n[ISPH — Healthcare / Pharmaceuticals]\nTechnical: RSI 38 post-devaluation. Below SMA50. BB position 22% — underperformer.\nFundamental: DEVALUATION NEGATIVE for pharma — raw material imports repriced +38% overnight. EGP revenue, USD cost = severe margin compression. Government price controls prevent passing costs to consumers. P/E 15x = expensive post-devaluation. ROE falling.\nVerdict: REDUCE — pharma underperforms in devaluation environments until price adjustments are approved. Hold minimum position only until margin clarity returns (typically 2-3 quarters post-devaluation).\n\nDCA in devaluation environment: Rotate toward banking (COMI) and hard-asset real estate (TMGH, ORAS). Reduce pharma and consumer staples with heavy import exposure. In mutual funds: shift from fixed income (eroded by inflation) toward equity funds.\n══════════════════════════════════════════════════\n\nNow apply this same depth and contextual awareness to the portfolio below."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.25,
            max_tokens=6000,
        )
        return {
            "analysis":    response.choices[0].message.content,
            "tokens_used": response.usage.total_tokens,
            "cost_usd":    (response.usage.total_tokens / 1_000_000) * 0.14,
        }
    except Exception as e:
        return {"error": str(e)}

# ===============================================
# DISPLAY HELPERS
# ===============================================
def _bar(value, low, high, width=10):
    try:
        pct    = (float(value) - float(low)) / (float(high) - float(low))
        pct    = max(0.0, min(1.0, pct))
        filled = int(pct * width)
        return "[" + "█"*filled + "░"*(width-filled) + f"] {pct*100:.0f}%"
    except:
        return "[??????????] N/A"

def _sig_icon(s):
    return {"STRONG BUY":"🟢🟢","BUY":"🟢  ","NEUTRAL":"🟡  ","SELL":"🔴  ","STRONG SELL":"🔴🔴"}.get(s,"⚪  ")

def _ri(pct):
    try:    return "🟢" if float(pct) >= 0 else "🔴"
    except: return "⚪"

def _sep(w=90): return "─"*w

# ===============================================
# DISPLAY REPORT
# ===============================================
def display_report(portfolio, market, tv_data, news, ai_result, monthly_egp):
    W  = 90
    LS = []
    def p(s=""): LS.append(str(s))

    total   = portfolio["summary"]["total_value"]
    ret_egp = portfolio["summary"]["total_return"]
    sv      = sum(d["market_value"] for d in portfolio["stocks"].values())
    mv      = sum(d["market_value"] for d in portfolio["mutual_funds"].values())
    ret_pct = ret_egp / total * 100 if total else 0

    # HEADER
    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  📊  THNDR PORTFOLIO REPORT   " + str(datetime.date.today()) +
       f"   DCA: {monthly_egp:,} EGP/month").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")

    # MARKET
    p(); p("┌─ 🌍  LIVE MARKET CONTEXT " + "─"*(W-27) + "┐")
    p(f"│  {'Indicator':<12} {'Value':>14}  {'Change':>9}   Context")
    p(f"│  {_sep(W-4)}")
    p(f"│  {'EGX30':<12} {market['egx30']['value']:>14}  {market['egx30']['change_pct']:>9}   Broad market (top 30)")
    p(f"│  {'EGX70':<12} {market['egx70']['value']:>14}  {market['egx70']['change_pct']:>9}   Mid/small caps")
    p(f"│  {'EGX100':<12} {market['egx100']['value']:>14}  {market['egx100']['change_pct']:>9}   Broad 100 index")
    p(f"│  {'Oil (WTI)':<12} {'$'+market['oil']['price']:>14}  {market['oil']['change_pct']:>9}   → ABUK, AMOC")
    p(f"│  {'Gold':<12} {'$'+market['gold']['price']:>14}  {market['gold']['change_pct']:>9}   Inflation / sentiment")
    p(f"│  {'USD/EGP':<12} {market['usd_egp']['rate']:>14}  {market['usd_egp']['change_pct']:>9}   → Pharma, Import-heavy")
    p(f"│  Status: {market['market_status']}   |   Updated: {market['last_updated']}")
    
    p("└" + "─"*(W-2) + "┘")

    # PORTFOLIO SNAPSHOT
    p(); p("┌─ 💰  PORTFOLIO SNAPSHOT " + "─"*(W-26) + "┐")
    p(f"│  Total Value      {total:>12,.0f} EGP")
    p(f"│  Total Return     {ret_egp:>+12,.0f} EGP   ({ret_pct:+.2f}%)  {_ri(ret_pct)}")
    p(f"│  Stocks           {sv:>12,.0f} EGP   ({sv/total*100:.1f}%)")
    p(f"│  Mutual Funds     {mv:>12,.0f} EGP   ({mv/total*100:.1f}%)")
    p("└" + "─"*(W-2) + "┘")

    # STOCKS — POSITION TABLE
    p(); p("┌─ 📈  STOCK POSITIONS " + "─"*(W-23) + "┐")
    p(f"│  {'SYM':<6} {'SECTOR':<28} {'UNITS':>5} {'COST':>7} {'PRICE':>7} {'RET%':>8}  {'SIGNAL':<13} {'RSI':>5} {'MACD':<9} {'P/E':>5} {'ROE%':>6} {'DIV%':>5}")
    p(f"│  {_sep(W-4)}")
    for sym, d in portfolio["stocks"].items():
        tv  = tv_data.get(sym, {})
        sig = tv.get("overall", "N/A")
        sec = d.get("sector", "Unknown")[:27]
        p(f"│  {sym:<6} {sec:<28} {d['units']:>5.0f} {d['cost_price']:>7.2f} {d['current_price']:>7.2f} "
          f"{d['return_pct']:>+7.1f}%{_ri(d['return_pct'])} "
          f"{_sig_icon(sig)}{sig:<11} "
          f"{str(tv.get('rsi','N/A')):>5} {str(tv.get('macd_cross','N/A')):<9} "
          f"{str(tv.get('pe','N/A')):>5} {str(tv.get('roe','N/A')):>6} {str(tv.get('div_yield','N/A')):>5}")

    # STOCKS — TREND TABLE
    p(f"│")
    p(f"│  TREND & VOLATILITY:")
    p(f"│  {'SYM':<6} {'SMA20':>7} {'SMA50':>7} {'SMA200':>8}   {'LOW':>7}  {'─ 3M RANGE ─':^14}  {'HIGH':<7}  {'BB%':>5} {'Vol/Avg':>7} {'ATR':>6}")
    p(f"│  {_sep(W-4)}")
    for sym, d in portfolio["stocks"].items():
        tv   = tv_data.get(sym, {})
        a20  = "↑" if tv.get("above_sma20")  else "↓"
        a50  = "↑" if tv.get("above_sma50")  else "↓"
        a200 = "↑" if tv.get("above_sma200") else "↓"
        bar  = _bar(tv.get("close","N/A"), tv.get("low_3m","N/A"), tv.get("high_3m","N/A"))
        bb   = f"{tv.get('bb_position','N/A')}%" if tv.get("bb_position") != "N/A" else "N/A"
        vr   = f"x{tv.get('vol_ratio','N/A')}"
        p(f"│  {sym:<6} {a20}{str(tv.get('sma20','?')):<6} {a50}{str(tv.get('sma50','?')):<6} {a200}{str(tv.get('sma200','?')):<7}   "
          f"{str(tv.get('low_3m','?')):>7}  {bar}  {str(tv.get('high_3m','?')):<7}  "
          f"{bb:>5} {vr:>7} {str(tv.get('atr','N/A')):>6}")
    p("└" + "─"*(W-2) + "┘")

    # MUTUAL FUNDS
    p(); p("┌─ 🏦  MUTUAL FUNDS " + "─"*(W-20) + "┐")
    p(f"│  {'SYM':<6} {'TYPE':<28} {'VALUE (EGP)':>14} {'RETURN':>9}  {'COST':>8} {'PRICE':>8}")
    p(f"│  {_sep(W-4)}")
    for sym, d in portfolio["mutual_funds"].items():
        sec = d.get("sector","Unknown")[len("Mutual Fund / "):][:27]
        p(f"│  {sym:<6} {sec:<28} {d['market_value']:>14,.0f} {d['return_pct']:>+8.1f}%{_ri(d['return_pct'])}  "
          f"{d['cost_price']:>8.2f} {d['current_price']:>8.2f}")
    p("└" + "─"*(W-2) + "┘")

    # NEWS
    p(); p("┌─ 📰  LATEST BUSINESS NEWS " + "─"*(W-28) + "┐")
    for i, n in enumerate(news[:5], 1):
        p(f"│  {i}. {n[:W-6]}")
    p("└" + "─"*(W-2) + "┘")

    # AI ANALYSIS
    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  🤖  DEEPSEEK ANALYSIS  —  6-12 Month Medium-Term + Monthly DCA Plan").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")

    if "error" in ai_result:
        p(f"  ❌  {ai_result['error']}")
    else:
        for line in ai_result["analysis"].split("\n"):
            stripped = line.strip()
            if stripped.startswith("## "):
                p(); p("  ┌─ " + stripped[3:].upper() + " " + "─"*max(0, W-8-len(stripped)) + "┐")
            elif stripped.startswith("- ") or stripped.startswith("* "):
                p("  │   • " + stripped[2:])
            elif stripped.startswith("**") and stripped.endswith("**"):
                p("  │  ► " + stripped.strip("*"))
            elif stripped == "":
                p("  │")
            else:
                words = stripped.split()
                cur   = "  │  "
                for w in words:
                    if len(cur) + len(w) + 1 > W:
                        p(cur); cur = "  │   " + w
                    else:
                        cur += (" " if cur != "  │  " else "") + w
                p(cur)
        p(); p(f"  {_sep()}")
        p(f"  Tokens: {ai_result['tokens_used']:,}   |   Cost: ${ai_result['cost_usd']:.5f} USD")

    p(); p(_sep()); p("  ⚠️  Informational only. Not financial advice."); p(_sep()); p()

    output = "\n".join(LS)
    print(output)
    return output


# ===============================================
# PRICE PREDICTIONS — separate DeepSeek call
# ===============================================
def deepseek_price_predictions(portfolio, market, tv_data, monthly_egp):
    print("\n🔮 Running price predictions (separate call)...")

    today = datetime.date.today()

    stock_lines = []
    for sym, d in portfolio["stocks"].items():
        tv = tv_data.get(sym, {})
        stock_lines.append(
            f"{sym} ({tv.get('company_name', sym)}) | Sector: {tv.get('sector','?')} / {tv.get('industry','?')}"
            f" | price={d['current_price']:.2f} EGP | cost={d['cost_price']:.2f} | ret={d['return_pct']:+.1f}%"
            f" | PE={tv.get('pe','N/A')} PB={tv.get('pb','N/A')} EPS={tv.get('eps','N/A')} ROE={tv.get('roe','N/A')}%"
            f" | RSI={tv.get('rsi','N/A')} MACD={tv.get('macd_cross','N/A')} Signal={tv.get('overall','N/A')}"
            f" | SMA50={tv.get('sma50','N/A')} SMA200={tv.get('sma200','N/A')}"
            f" | 3M_range={tv.get('low_3m','N/A')}-{tv.get('high_3m','N/A')} pos={tv.get('range_3m_position','N/A')}%"
            f" | ATR={tv.get('atr','N/A')} BB_pos={tv.get('bb_position','N/A')}%"
        )

    prompt = f"""You are a quantitative analyst specializing in EGX (Egyptian Exchange) stocks.
Date: {today}
Market: EGX30={market['egx30']['value']} ({market['egx30']['change_pct']}) | USD/EGP={market['usd_egp']['rate']} | Oil=${market['oil']['price']} | Gold=${market['gold']['price']}

STOCKS WITH FULL TECHNICAL & FUNDAMENTAL DATA:
{chr(10).join(stock_lines)}

For EACH stock above, provide price predictions across 5 timeframes.
Base predictions on: current technicals (RSI, MACD, BB, Stoch), trend (SMA50/200), 3M price range position, fundamentals (PE, ROE, EPS), ATR for volatility sizing, and sector macro context.

Format your response EXACTLY like this for every stock (no deviations):

### {"{SYM}"}
- Current Price: X.XX EGP
- Next Week (7 days): LOW X.XX — HIGH X.XX EGP | Bias: BULLISH/BEARISH/NEUTRAL | Key driver: [one sentence]
- Next Month (30 days): LOW X.XX — HIGH X.XX EGP | Bias: BULLISH/BEARISH/NEUTRAL | Key driver: [one sentence]
- 3 Months: LOW X.XX — HIGH X.XX EGP | Bias: BULLISH/BEARISH/NEUTRAL | Key driver: [one sentence]
- 6 Months: LOW X.XX — HIGH X.XX EGP | Target: X.XX EGP | Upside/Downside: X% | Key driver: [one sentence]
- 12 Months: LOW X.XX — HIGH X.XX EGP | Target: X.XX EGP | Upside/Downside: X% | Key driver: [one sentence]
- Key Risk: [one sentence on what could invalidate the bullish/bearish case]
- Key Catalyst: [one sentence on what could accelerate the move]

Be realistic. Use ATR to size weekly/monthly ranges. Use fundamentals for 6-12M targets.
Acknowledge uncertainty. Give ranges not just single numbers.
Reference the actual data provided — do not make up numbers."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "You are a quantitative EGX analyst. You provide precise, data-driven price predictions with realistic ranges. You always reference the specific technical and fundamental data provided."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=8000,
        )
        content = response.choices[0].message.content
        finish = getattr(response.choices[0], "finish_reason", None)
        if finish == "length":
            print(f"   ⚠️  Price predictions were truncated by token limit (got {response.usage.completion_tokens} output tokens).")
            content += "\r\n\r\n⚠️  [OUTPUT TRUNCATED — DeepSeek hit max_tokens. Reduce WATCHLIST_SYMBOLS or owned stocks to fit.]"
        return {
            "predictions":  content,
            "tokens_used":  response.usage.total_tokens,
            "cost_usd":     (response.usage.total_tokens / 1_000_000) * 0.14,
            "truncated":    finish == "length",
        }
    except Exception as e:
        return {"error": str(e)}


# ===============================================
# ARABIC REPORT — separate DeepSeek call
# ===============================================
def deepseek_arabic_report(portfolio, market, tv_data, news, ai_result, predictions, monthly_egp):
    print("\n🇪🇬 Generating Arabic report...")

    today     = datetime.date.today()
    total_val = portfolio["summary"]["total_value"]
    total_ret = portfolio["summary"]["total_return"]

    # Compact stock summary for Arabic prompt
    stock_lines = []
    for sym, d in portfolio["stocks"].items():
        tv = tv_data.get(sym, {})
        stock_lines.append(
            f"{sym} ({tv.get('company_name', sym)}) | قطاع: {tv.get('sector','?')} | "
            f"السعر: {d['current_price']:.2f} جنيه | العائد: {d['return_pct']:+.1f}% | "
            f"إشارة: {tv.get('overall','N/A')} | RSI={tv.get('rsi','N/A')} | "
            f"P/E={tv.get('pe','N/A')} | ROE={tv.get('roe','N/A')}%"
        )

    english_analysis = ai_result.get("analysis","") if "analysis" in ai_result else ""
    price_preds      = predictions.get("predictions","") if "predictions" in predictions else ""

    prompt = f"""أنت مترجم ومبسط للتقارير المالية. مهمتك الوحيدة هي إنك تترجم وتبسط التحليل الإنجليزي الموجود تحت بالظبط — من غير ما تغير أي توصية أو رأي.

⚠️ قاعدة أساسية: لو التحليل الإنجليزي قال "ACCUMULATE" قول "اشتري أكتر". لو قال "HOLD" قول "احتفظ". لو قال "REDUCE" قول "ابيع جزء". لا تغير الرأي أبداً.

التاريخ: {today}
ميزانية الاستثمار الشهري: {monthly_egp:,} جنيه مصري

السوق دلوقتي:
- مؤشر EGX30: {market['egx30']['value']} ({market['egx30']['change_pct']})
- مؤشر EGX70: {market['egx70']['value']} ({market['egx70']['change_pct']})
- مؤشر EGX100: {market['egx100']['value']} ({market['egx100']['change_pct']})
- سعر البترول: ${market['oil']['price']} ({market['oil']['change_pct']})
- سعر الذهب: ${market['gold']['price']} ({market['gold']['change_pct']})
- سعر الدولار: {market['usd_egp']['rate']} جنيه

المحفظة:
القيمة الإجمالية: {total_val:,.0f} جنيه
الربح/الخسارة: {total_ret:+,.0f} جنيه ({total_ret/total_val*100:+.1f}%)

الأسهم:
{chr(10).join(stock_lines)}

══════════════════════════════════
التحليل الإنجليزي الكامل — ترجمه وبسطه بالعربي المصري:
══════════════════════════════════
{english_analysis}

══════════════════════════════════
توقعات الأسعار — استخدمها في القسم التالت:
══════════════════════════════════
{price_preds[:3000]}

اكتب التقرير بالعربي المصري العامي على النقاط دي — مع الالتزام الكامل بتوصيات التحليل الإنجليزي:

## ١. إيه اللي بيحصل في السوق دلوقتي؟
اشرح وضع السوق بكلام بسيط. البترول والذهب والدولار بيأثروا إزاي على أسهمنا بالظبط؟

## ٢. محفظتك إزاي؟
اشرح الوضع العام. ربحانين ولا خسرانين؟ فين الأقوى وفين الأضعف؟

## ٣. تحليل كل سهم (بكلام بسيط)
لكل سهم — ترجم بالظبط من التحليل الإنجليزي:
- وضع السهم دلوقتي بكلام بسيط (مثلاً: "السهم قوي وفوق المتوسطات")
- السبب ببساطة من غير مصطلحات
- التوقعات من توقعات الأسعار (الأسبوع الجاي / الشهر الجاي / 6 شهور)
- الرأي بالظبط من التحليل الإنجليزي: اشتري أكتر / احتفظ / ابيع جزء

## ٤. فين تحط فلوسك الشهر ده؟
ترجم خطة الـ DCA من التحليل الإنجليزي بالظبط:
- نفس الأسهم المقترحة ونفس التوزيع بالجنيه
- سعر الدخول ووقف الخسارة والهدف السعري

## ٥. المخاطر اللي لازم تعرفها
ترجم نفس المخاطر من التحليل الإنجليزي بكلام بسيط.

## ٦. إيه اللي هتعمله أول ما البورصة تفتح؟
ترجم نفس خطة العمل من التحليل الإنجليزي.

الأسلوب: عربي مصري عامي — سهل وواضح — الأرقام بالجنيه — لا تغير أي توصية."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "أنت مترجم ومبسط للتقارير المالية. مهمتك ترجمة التحليل الإنجليزي للعربي المصري العامي بدون تغيير أي توصية أو رأي. لو التحليل قال اشتري، قول اشتري. لو قال احتفظ، قول احتفظ. الأرقام والتوصيات لازم تكون متطابقة."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            max_tokens=4000,
        )
        return {
            "arabic":      response.choices[0].message.content,
            "tokens_used": response.usage.total_tokens,
            "cost_usd":    (response.usage.total_tokens / 1_000_000) * 0.14,
        }
    except Exception as e:
        return {"error": str(e)}


# ===============================================
# RECOMMENDATION ENGINE — Step 1: Pick symbols
# ===============================================
def deepseek_pick_recommendations(portfolio, market, num_medium, num_long, prefer_diversification):
    """First call: DeepSeek picks EGX symbols based on macro context + diversification preference.
    Returns ONLY a JSON list of symbols + brief seed-thesis. Live data is fetched after."""
    print("\n🎯 [Recommendation Engine 1/2] Picking candidate stocks...")

    today = datetime.date.today()
    owned = list(portfolio["stocks"].keys())
    owned_sectors = sorted({d.get("sector","Unknown") for d in portfolio["stocks"].values()})

    diversify_rule = (
        "STRONGLY PREFER stocks in SECTORS NOT ALREADY OWNED. Only suggest a stock in an owned sector if it's an exceptional opportunity (top-3 best EGX setup right now). When suggesting in an owned sector, justify why it's exceptional."
        if prefer_diversification else
        "Sector overlap with owned stocks is fine. Pick the best opportunities regardless of sector."
    )

    prompt = f"""You are a senior EGX (Egyptian Exchange) analyst. Pick stock recommendations for a client.

Date: {today}
Market context: EGX30={market['egx30']['value']} ({market['egx30']['change_pct']}) | USD/EGP={market['usd_egp']['rate']} | Oil=${market['oil']['price']} | Gold=${market['gold']['price']}

Client's CURRENT HOLDINGS (do NOT recommend these — they already own them):
Symbols: {', '.join(owned)}
Sectors held: {', '.join(owned_sectors)}

DIVERSIFICATION RULE:
{diversify_rule}

YOUR TASK:
Pick EXACTLY {num_medium} MEDIUM-TERM (1-3 month) candidates and {num_long} LONG-TERM (6 months to 3 years) candidates from the EGX.

MEDIUM-TERM criteria (1-3 months): Pick stocks with strong technical setups (breakout patterns, momentum, sector rotation), positive news flow, or earnings catalysts within 90 days. Liquid stocks preferred.

LONG-TERM criteria (6m-3y): Pick fundamentally strong compounders — durable competitive advantage, growing earnings, reasonable valuation, dividend yield a plus. Quality over hype.

Choose from REAL EGX-listed companies. Use their actual EGX ticker symbols (e.g. HRHO, EAST, MFPC, EFIH, JUFO, ETEL, SWDY, ADIB, CIEB, EKHO, OCDI, PHDC, MNHD, AUTO, RAYA, IRON, EFID, SKPC, DOMT, MTIE, etc).

Return your answer in this EXACT JSON format and nothing else (no markdown fences, no commentary):

{{
  "medium_term": [
    {{"symbol": "TICKER1", "company": "Full company name", "sector": "Sector name", "thesis": "One-sentence reason picking this for 1-3 months."}},
    {{"symbol": "TICKER2", "company": "...", "sector": "...", "thesis": "..."}},
    {{"symbol": "TICKER3", "company": "...", "sector": "...", "thesis": "..."}}
  ],
  "long_term": [
    {{"symbol": "TICKER4", "company": "...", "sector": "...", "thesis": "One-sentence reason picking this for 6m-3y."}},
    {{"symbol": "TICKER5", "company": "...", "sector": "...", "thesis": "..."}},
    {{"symbol": "TICKER6", "company": "...", "sector": "...", "thesis": "..."}}
  ]
}}

IMPORTANT: Return ONLY the JSON object. No prefix, no suffix, no markdown."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "You are an EGX analyst who returns valid JSON only. Never wrap output in markdown fences. Never add commentary outside the JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            max_tokens=1500,
        )
        raw = response.choices[0].message.content.strip()
        # Strip markdown fences if model added them despite instructions
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        raw = raw.strip()

        import json
        try:
            picks = json.loads(raw)
        except Exception as je:
            # Try to extract first JSON object from response
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            if m:
                picks = json.loads(m.group(0))
            else:
                raise je

        med = picks.get("medium_term", [])
        lng = picks.get("long_term", [])
        print(f"   ✅ Medium-term picks ({len(med)}): {[p['symbol'] for p in med]}")
        print(f"   ✅ Long-term picks  ({len(lng)}): {[p['symbol'] for p in lng]}")
        return {
            "medium_term": med,
            "long_term":   lng,
            "tokens_used": response.usage.total_tokens,
            "cost_usd":    (response.usage.total_tokens / 1_000_000) * 0.14,
        }
    except Exception as e:
        print(f"   ❌ Pick error: {e}")
        return {"error": str(e), "medium_term": [], "long_term": [], "tokens_used": 0, "cost_usd": 0}


# ===============================================
# RECOMMENDATION ENGINE — Step 2: Refine with live data
# ===============================================
def deepseek_refine_recommendations(picks, tv_data, market, monthly_egp):
    """Second call: send live TradingView data for picks, get detailed analysis with real numbers."""
    print("\n🔬 [Recommendation Engine 2/2] Refining with live data...")

    if "error" in picks or (not picks.get("medium_term") and not picks.get("long_term")):
        return {"error": picks.get("error", "No picks to refine"), "tokens_used": 0, "cost_usd": 0}

    def fmt_pick_block(p):
        sym = p["symbol"]
        tv  = tv_data.get(sym, {})
        if tv.get("close", "N/A") == "N/A":
            return (
                f"{sym} ({p.get('company','?')}) | Sector: {p.get('sector','?')}\n"
                f"  ⚠️ NO LIVE DATA from TradingView (symbol may be wrong/illiquid/delisted)\n"
                f"  Original thesis: {p.get('thesis','?')}\n"
            )
        return (
            f"{sym} ({tv.get('company_name', p.get('company','?'))}) | Sector: {tv.get('sector', p.get('sector','?'))}\n"
            f"  PRICE: {tv.get('close','?')} EGP ({tv.get('change_pct','?')}% today)\n"
            f"  TECHNICAL: Signal={tv.get('overall','?')} | RSI={tv.get('rsi','?')} (prev={tv.get('rsi_prev','?')}) | "
            f"MACD={tv.get('macd_cross','?')} | Stoch_K={tv.get('stoch_k','?')}\n"
            f"  TREND: SMA20={tv.get('sma20','?')} SMA50={tv.get('sma50','?')} SMA200={tv.get('sma200','?')} | "
            f"3M range={tv.get('low_3m','?')}-{tv.get('high_3m','?')} (pos {tv.get('range_3m_position','?')}%)\n"
            f"  VOLATILITY: ATR={tv.get('atr','?')} | BB_pos={tv.get('bb_position','?')}% | Vol/avg=x{tv.get('vol_ratio','?')}\n"
            f"  FUNDAMENTALS: PE={tv.get('pe','?')} | PB={tv.get('pb','?')} | EPS={tv.get('eps','?')} | "
            f"ROE={tv.get('roe','?')}% | Div={tv.get('div_yield','?')}% | D/E={tv.get('debt_equity','?')}\n"
            f"  Original thesis: {p.get('thesis','?')}\n"
        )

    med_lines = [fmt_pick_block(p) for p in picks.get("medium_term", [])]
    lng_lines = [fmt_pick_block(p) for p in picks.get("long_term", [])]

    prompt = f"""You are a senior EGX analyst. You previously picked these stocks for a client. Now refine each recommendation using the LIVE technical and fundamental data below. Use REAL numbers — do not invent prices.

Date: {datetime.date.today()}
Market: EGX30={market['egx30']['value']} ({market['egx30']['change_pct']}) | USD/EGP={market['usd_egp']['rate']} | Oil=${market['oil']['price']} | Gold=${market['gold']['price']}
Client's monthly DCA budget: {monthly_egp:,} EGP

═══════════════════════════════════
MEDIUM-TERM PICKS (1-3 months) — LIVE DATA
═══════════════════════════════════
{chr(10).join(med_lines)}

═══════════════════════════════════
LONG-TERM PICKS (6 months to 3 years) — LIVE DATA
═══════════════════════════════════
{chr(10).join(lng_lines)}

For EACH stock above, provide a recommendation block in EXACTLY this format:

### {{SYM}} — {{Company}} [{{Sector}}]
- Horizon: MEDIUM-TERM (1-3 months) or LONG-TERM (6m-3y)
- Current Price: X.XX EGP
- Entry Zone: Buy at or below X.XX EGP (sweet spot: X.XX-X.XX EGP)
- Stop Loss: X.XX EGP (you were wrong below this — lose ~X% from entry)
- Targets: T1=X.XX (timeline) | T2=X.XX (timeline) | T3=X.XX (timeline)
- Upside: +X% to T2 (over Y months)
- Technical Read: [2-3 sentences on RSI, MACD, trend, range position — use actual numbers from data above]
- Fundamental Read: [2-3 sentences on PE, ROE, EPS, D/E — compare to sector norm — use actual numbers]
- Macro Fit: [1-2 sentences on how current market conditions support or threaten this thesis]
- Why This Pick: [The one specific reason this stock right now]
- Key Risk: [What could break the thesis]
- Conviction: HIGH / MEDIUM / LOW
- Suggested Allocation: X% of monthly DCA budget = X,XXX EGP (only suggest if conviction is MEDIUM or higher)

If a stock's live data shows "NO LIVE DATA", say so and provide thesis-only analysis with caveat that prices need manual verification before trading.

After all stock blocks, add ONE section:

### 💡 PORTFOLIO FIT SUMMARY
- How these picks complement (or risk concentrating) the client's existing holdings
- Recommended order of priority (which to buy first, which to wait on)
- Total budget allocation across all 6 picks vs. existing-holdings DCA

Be precise. Use real numbers throughout. No filler."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "You are a senior EGX analyst providing actionable, data-driven stock recommendations with realistic price levels based on live technical and fundamental data. Always reference the specific numbers provided. Use ATR for stop-loss sizing and 3M range for entry zones."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.25,
            max_tokens=4500,
        )
        return {
            "recommendations": response.choices[0].message.content,
            "tokens_used":     response.usage.total_tokens,
            "cost_usd":        (response.usage.total_tokens / 1_000_000) * 0.14,
            "raw_picks":       picks,  # keep for display
        }
    except Exception as e:
        print(f"   ❌ Refine error: {e}")
        return {"error": str(e), "tokens_used": 0, "cost_usd": 0}


# ===============================================
# DISPLAY RECOMMENDATIONS
# ===============================================
def display_recommendations(rec_result, picks):
    W  = 90
    LS = []
    def p(s=""): LS.append(str(s))

    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  💡  NEW STOCK IDEAS  —  Stocks you DON'T currently own").ljust(W-1) + "║")
    p(("║      Medium-Term (1-3 months) + Long-Term (6m to 3y)").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")

    # Quick pick summary table at top
    p()
    p("┌─ 🎯 PICKS AT A GLANCE " + "─"*(W-25) + "┐")
    p(f"│  {'HORIZON':<14} {'SYM':<6} {'COMPANY':<35} {'SECTOR':<28}")
    p(f"│  {_sep(W-4)}")
    for pk in picks.get("medium_term", []):
        p(f"│  {'MEDIUM (1-3M)':<14} {pk.get('symbol','?'):<6} {pk.get('company','?')[:34]:<35} {pk.get('sector','?')[:27]:<28}")
    for pk in picks.get("long_term", []):
        p(f"│  {'LONG (6m-3y)':<14} {pk.get('symbol','?'):<6} {pk.get('company','?')[:34]:<35} {pk.get('sector','?')[:27]:<28}")
    p("└" + "─"*(W-2) + "┘")
    p()

    if "error" in rec_result:
        p(f"  ❌  {rec_result['error']}")
    else:
        for line in rec_result["recommendations"].split("\n"):
            stripped = line.strip()
            if stripped.startswith("### "):
                p(); p("  ┌─ 📌 " + stripped[4:] + " " + "─"*max(0, W-12-len(stripped)) + "┐")
            elif stripped.startswith("- Horizon"):
                p("  │  ⏱️  " + stripped[2:])
            elif stripped.startswith("- Current Price"):
                p("  │  💰 " + stripped[2:])
            elif stripped.startswith("- Entry Zone"):
                p("  │  🎯 " + stripped[2:])
            elif stripped.startswith("- Stop Loss"):
                p("  │  🛑 " + stripped[2:])
            elif stripped.startswith("- Targets"):
                p("  │  📊 " + stripped[2:])
            elif stripped.startswith("- Upside"):
                p("  │  📈 " + stripped[2:])
            elif stripped.startswith("- Technical Read"):
                p("  │  📉 " + stripped[2:])
            elif stripped.startswith("- Fundamental Read"):
                p("  │  💼 " + stripped[2:])
            elif stripped.startswith("- Macro Fit"):
                p("  │  🌍 " + stripped[2:])
            elif stripped.startswith("- Why This Pick"):
                p("  │  ✅ " + stripped[2:])
            elif stripped.startswith("- Key Risk"):
                p("  │  ⚠️  " + stripped[2:])
            elif stripped.startswith("- Conviction"):
                p("  │  🔥 " + stripped[2:])
            elif stripped.startswith("- Suggested Allocation"):
                p("  │  💵 " + stripped[2:])
            elif stripped.startswith("- "):
                p("  │   • " + stripped[2:])
            elif stripped == "":
                p("  │")
            else:
                p("  │  " + stripped)
        p(); p("─"*W)
        p(f"  Tokens: {rec_result['tokens_used']:,}   |   Cost: ${rec_result['cost_usd']:.5f} USD")

    p(); p("─"*W)
    p("  ⚠️  Recommendations are informational. Verify prices live before trading.")
    p("─"*W); p()

    output = "\n".join(LS)
    print(output)
    return output


# ===============================================
# WATCHLIST TRADE PLAN — Detailed entry/stop/target for watchlist stocks
# ===============================================
def deepseek_watchlist_trade_plan(portfolio, tv_data, market, monthly_egp):
    """Generate detailed entry zones, stop losses, and targets for watchlist stocks (units=0)."""
    print("\n📋 Generating watchlist trade plans...")

    # Find watchlist symbols (units=0 with watchlist flag)
    watchlist = [(sym, d) for sym, d in portfolio["stocks"].items() if d.get("watchlist")]
    if not watchlist:
        print("   ℹ️  No watchlist stocks to analyze")
        return {"trade_plans": "", "tokens_used": 0, "cost_usd": 0, "skipped": True}

    print(f"   📊 Building plans for: {[s for s,_ in watchlist]}")

    stock_blocks = []
    for sym, d in watchlist:
        tv = tv_data.get(sym, {})
        if tv.get("close", "N/A") == "N/A":
            stock_blocks.append(
                f"{sym} ({tv.get('company_name', sym)}) | Sector: {tv.get('sector','?')}\n"
                f"  ⚠️ NO LIVE DATA from TradingView — provide thesis-only analysis with caveat\n"
            )
            continue
        stock_blocks.append(
            f"{sym} ({tv.get('company_name', sym)}) | Sector: {tv.get('sector','?')}\n"
            f"  PRICE: {tv.get('close','?')} EGP ({tv.get('change_pct','?')}% today)\n"
            f"  TECHNICAL: Signal={tv.get('overall','?')} | RSI={tv.get('rsi','?')} (prev={tv.get('rsi_prev','?')}) | "
            f"MACD={tv.get('macd_cross','?')} | Stoch_K={tv.get('stoch_k','?')} | Stoch_D={tv.get('stoch_d','?')}\n"
            f"  TREND: SMA20={tv.get('sma20','?')} SMA50={tv.get('sma50','?')} SMA200={tv.get('sma200','?')} | "
            f"3M range={tv.get('low_3m','?')}-{tv.get('high_3m','?')} (pos {tv.get('range_3m_position','?')}%) | "
            f"1M={tv.get('low_1m','?')}-{tv.get('high_1m','?')}\n"
            f"  VOLATILITY: ATR={tv.get('atr','?')} | BB_pos={tv.get('bb_position','?')}% | Vol/avg=x{tv.get('vol_ratio','?')}\n"
            f"  FUNDAMENTALS: PE={tv.get('pe','?')} | PB={tv.get('pb','?')} | EPS={tv.get('eps','?')} | "
            f"ROE={tv.get('roe','?')}% | Div={tv.get('div_yield','?')}% | D/E={tv.get('debt_equity','?')}\n"
        )

    prompt = f"""You are a senior EGX analyst. The client is WATCHING (not yet owning) the stocks below and wants a complete trading playbook for each — exact entry zones, stop-losses, and targets based on LIVE technical and fundamental data.

Date: {datetime.date.today()}
Market: EGX30={market['egx30']['value']} ({market['egx30']['change_pct']}) | USD/EGP={market['usd_egp']['rate']} | Oil=${market['oil']['price']} | Gold=${market['gold']['price']}
Client's monthly DCA budget: {monthly_egp:,} EGP

═══════════════════════════════════
WATCHLIST STOCKS — LIVE DATA
═══════════════════════════════════
{chr(10).join(stock_blocks)}

For EACH stock above, provide a trade plan in EXACTLY this format:

### {{SYM}} — {{Company}} [{{Sector}}]
- Current Price: X.XX EGP
- Setup Quality: EXCELLENT / GOOD / FAIR / POOR — [one phrase explaining why]
- Entry Zone: Buy at or below X.XX EGP (sweet spot: X.XX-X.XX EGP)
- Aggressive Entry: X.XX EGP (buy now if you're bullish)
- Conservative Entry: X.XX EGP (wait for pullback to this level)
- Stop Loss: X.XX EGP (you were wrong below this — risk ~X% from entry)
- Target 1 (Short-term, 1-2 months): X.XX EGP (+X%)
- Target 2 (Medium-term, 3-6 months): X.XX EGP (+X%)
- Target 3 (Long-term, 12 months): X.XX EGP (+X%)
- Position Sizing: Suggested allocation X% of monthly DCA = X,XXX EGP
- Technical Read: [2-3 sentences using actual RSI/MACD/SMA numbers from data above]
- Fundamental Read: [2-3 sentences on PE/ROE/EPS — compare to sector if possible]
- Why Buy: [The specific catalyst or setup right now]
- Why NOT Buy Yet: [Any reason to wait — overbought? earnings ahead? macro risk?]
- Risk/Reward: X.X:X (e.g. 3:1 means risking 1 EGP to make 3)
- Verdict: BUY NOW / WAIT FOR ENTRY / AVOID

Use ATR to size stop-loss distance (typically 1.5-2x ATR below entry).
Use 3M range position to assess if the stock is at a good entry point.
If "NO LIVE DATA" was shown, say so and provide thesis-only analysis with caveat.
Use REAL numbers from the data above — do not invent prices."""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": "You are a senior EGX analyst providing precise trade plans with realistic entry zones, stop-losses sized using ATR, and price targets based on technical and fundamental data. Always reference the specific numbers provided. Be a careful trader, not a hype-man."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.25,
            max_tokens=3500,
        )
        return {
            "trade_plans": response.choices[0].message.content,
            "tokens_used": response.usage.total_tokens,
            "cost_usd":    (response.usage.total_tokens / 1_000_000) * 0.14,
            "skipped":     False,
        }
    except Exception as e:
        print(f"   ❌ Watchlist plan error: {e}")
        return {"error": str(e), "tokens_used": 0, "cost_usd": 0, "skipped": False}


# ===============================================
# DISPLAY WATCHLIST TRADE PLANS
# ===============================================
def display_watchlist_trade_plans(plan_result):
    if plan_result.get("skipped"):
        return ""

    W  = 90
    LS = []
    def p(s=""): LS.append(str(s))

    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  👁️   WATCHLIST TRADE PLANS  —  Ready-to-execute entries for tracked stocks").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")

    if "error" in plan_result:
        p(f"  ❌  {plan_result['error']}")
    else:
        for line in plan_result["trade_plans"].split("\n"):
            stripped = line.strip()
            if stripped.startswith("### "):
                p(); p("  ┌─ 👁️  " + stripped[4:] + " " + "─"*max(0, W-12-len(stripped)) + "┐")
            elif stripped.startswith("- Current Price"):
                p("  │  💰 " + stripped[2:])
            elif stripped.startswith("- Setup Quality"):
                p("  │  ⭐ " + stripped[2:])
            elif stripped.startswith("- Entry Zone"):
                p("  │  🎯 " + stripped[2:])
            elif stripped.startswith("- Aggressive Entry"):
                p("  │  🟢 " + stripped[2:])
            elif stripped.startswith("- Conservative Entry"):
                p("  │  🟡 " + stripped[2:])
            elif stripped.startswith("- Stop Loss"):
                p("  │  🛑 " + stripped[2:])
            elif stripped.startswith("- Target 1"):
                p("  │  📊 " + stripped[2:])
            elif stripped.startswith("- Target 2"):
                p("  │  📈 " + stripped[2:])
            elif stripped.startswith("- Target 3"):
                p("  │  🎯 " + stripped[2:])
            elif stripped.startswith("- Position Sizing"):
                p("  │  💵 " + stripped[2:])
            elif stripped.startswith("- Technical Read"):
                p("  │  📉 " + stripped[2:])
            elif stripped.startswith("- Fundamental Read"):
                p("  │  💼 " + stripped[2:])
            elif stripped.startswith("- Why Buy"):
                p("  │  ✅ " + stripped[2:])
            elif stripped.startswith("- Why NOT"):
                p("  │  ⏸️  " + stripped[2:])
            elif stripped.startswith("- Risk/Reward"):
                p("  │  ⚖️  " + stripped[2:])
            elif stripped.startswith("- Verdict"):
                p("  │  🔥 " + stripped[2:])
            elif stripped.startswith("- "):
                p("  │   • " + stripped[2:])
            elif stripped == "":
                p("  │")
            else:
                p("  │  " + stripped)
        p(); p("─"*W)
        p(f"  Tokens: {plan_result['tokens_used']:,}   |   Cost: ${plan_result['cost_usd']:.5f} USD")

    p(); p("─"*W)
    p("  ⚠️  Trade plans are informational. Verify prices live before placing orders.")
    p("─"*W); p()

    output = "\n".join(LS)
    print(output)
    return output


# ===============================================
# DISPLAY PRICE PREDICTIONS
# ===============================================
def display_predictions(predictions):
    W  = 90
    LS = []
    def p(s=""): LS.append(str(s))

    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  🔮  PRICE PREDICTIONS  —  1 Week / 1 Month / 3 Months / 6 Months / 12 Months").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")
    p("║  ⚠️  Predictions are probabilistic estimates, NOT guarantees. Always use stop-losses.")
    p("─"*W)

    if "error" in predictions:
        p(f"  ❌  {predictions['error']}")
    else:
        for line in predictions["predictions"].split("\n"):
            stripped = line.strip()
            if stripped.startswith("### "):
                p(); p("  ┌─ 📌 " + stripped[4:] + " " + "─"*max(0, W-10-len(stripped)) + "┐")
            elif stripped.startswith("- Next Week"):
                p("  │  📅 " + stripped[2:])
            elif stripped.startswith("- Next Month"):
                p("  │  📆 " + stripped[2:])
            elif stripped.startswith("- 3 Month"):
                p("  │  🗓️  " + stripped[2:])
            elif stripped.startswith("- 6 Month"):
                p("  │  📊 " + stripped[2:])
            elif stripped.startswith("- 12 Month"):
                p("  │  🎯 " + stripped[2:])
            elif stripped.startswith("- Current"):
                p("  │  💰 " + stripped[2:])
            elif stripped.startswith("- Key Risk"):
                p("  │  ⚠️  " + stripped[2:])
            elif stripped.startswith("- Key Catalyst"):
                p("  │  🚀 " + stripped[2:])
            elif stripped == "":
                p("  │")
            else:
                p("  │  " + stripped)
        p(); p("─"*W)
        p(f"  Tokens: {predictions['tokens_used']:,}   |   Cost: ${predictions['cost_usd']:.5f} USD")

    p(); p("─"*W); p()
    output = "\n".join(LS)
    print(output)
    return output


# ===============================================
# DISPLAY ARABIC REPORT
# ===============================================
def display_arabic_report(arabic_result):
    W  = 90
    LS = []
    def p(s=""): LS.append(str(s))

    p(); p("╔" + "═"*(W-2) + "╗")
    p(("║  🇪🇬  التقرير بالعربي  —  للمستثمر المصري العادي").ljust(W-1) + "║")
    p("╚" + "═"*(W-2) + "╝")

    if "error" in arabic_result:
        p(f"  ❌  {arabic_result['error']}")
    else:
        for line in arabic_result["arabic"].split("\n"):
            stripped = line.strip()
            if stripped.startswith("## "):
                p(); p("  ┌─ " + stripped[3:] + " " + "─"*max(0, W-8-len(stripped)) + "┐")
            elif stripped.startswith("- ") or stripped.startswith("• "):
                p("  │   • " + stripped[2:])
            elif stripped == "":
                p("  │")
            else:
                p("  │  " + stripped)
        p(); p("─"*W)
        p(f"  Tokens: {arabic_result['tokens_used']:,}   |   Cost: ${arabic_result['cost_usd']:.5f} USD")

    p(); p("─"*W)
    p("  ⚠️  التقرير ده للمعلومات بس ومش نصيحة استثمارية رسمية")
    p("─"*W); p()

    output = "\n".join(LS)
    print(output)
    return output

# Run this standalone to test PDF generation without burning tokens
# python test_pdf.py

import re, datetime

# ===============================================
# SAVE REPORT
# ===============================================
def save_report(report_text, predictions_text, arabic_text, recommendations_text="", watchlist_text="", folder=None):
    if folder is None:
        folder = os.getcwd()
    ts   = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    path = f"{folder}/portfolio_report_{ts}.txt"
    with open(path, "w", encoding="utf-8") as f:
        f.write(report_text)
        f.write("\n\n")
        f.write(predictions_text)
        f.write("\n\n")
        if watchlist_text:
            f.write(watchlist_text)
            f.write("\n\n")
        if recommendations_text:
            f.write(recommendations_text)
            f.write("\n\n")
        f.write(arabic_text)
    print(f"\n💾 Full report saved: {path}")

# ===============================================
# MAIN
# ===============================================
def main():
    print("\n" + "="*90)
    print("🚀 THNDR SMART PORTFOLIO ANALYZER  —  Medium-Term DCA Edition")
    print("="*90 + "\n")

    # Look for the investments file in the current working directory.
    # The .bat sets cwd to the Desktop output folder where the scraper just
    # saved investments_*.xlsx.
    SCRIPT_DIR = os.getcwd()
    excel_files = glob.glob(os.path.join(SCRIPT_DIR, "investments_*.xlsx"))
    if not excel_files:
        print(f"❌ No Excel file found in {SCRIPT_DIR}")
        return

    latest_excel = max(excel_files, key=os.path.getmtime)
    print(f"📂 Loading: {latest_excel}")

    try:
        portfolio = load_portfolio(latest_excel)
        print(f"✅ Loaded: {len(portfolio['stocks'])} stocks, {len(portfolio['mutual_funds'])} mutual funds")
        print(f"   Total Value: {portfolio['summary']['total_value']:,.0f} EGP\n")
    except Exception as e:
        print(f"❌ Error loading Excel: {e}"); return

    # ── Merge owned stocks + watchlist symbols for TradingView fetch ──
    owned_syms     = list(portfolio["stocks"].keys())
    watchlist_syms = [s for s in WATCHLIST_SYMBOLS if s not in owned_syms]
    all_syms       = owned_syms + watchlist_syms
    tv_data        = fetch_tradingview_data(all_syms)

    # Add watchlist stocks as "virtual" portfolio entries (0 units, just for analysis)
    for sym in watchlist_syms:
        tv = tv_data.get(sym, {})
        current_price = tv.get("close", 0) if tv.get("close") != "N/A" else 0
        portfolio["stocks"][sym] = {
            "symbol":        sym,
            "class":         "Stock (Watchlist)",
            "units":         0,
            "cost_price":    current_price,
            "current_price": current_price,
            "market_value":  0,
            "return_pct":    0.0,
            "return_egp":    0.0,
            "sector":        tv.get("sector", "—"),
            "company_name":  tv.get("company_name", sym),
            "watchlist":     True,
        }
        print(f"   👁️  Added watchlist: {sym} @ {current_price} EGP")

    # Update sector/company info for owned stocks from TradingView
    for sym in owned_syms:
        if sym in tv_data and tv_data[sym].get("sector","N/A") != "N/A":
            portfolio["stocks"][sym]["sector"]       = tv_data[sym]["sector"]
            portfolio["stocks"][sym]["company_name"] = tv_data[sym].get("company_name", sym)

    print()
    market     = fetch_market_context(); print()
    news       = fetch_news(); print()

    ai_result     = deepseek_analysis(portfolio, market, tv_data, news, MONTHLY_INVESTMENT_EGP)
    report_output = display_report(portfolio, market, tv_data, news, ai_result, MONTHLY_INVESTMENT_EGP)

    predictions     = deepseek_price_predictions(portfolio, market, tv_data, MONTHLY_INVESTMENT_EGP)
    pred_output     = display_predictions(predictions)

    # ── Watchlist Trade Plans (entry/stop/target for stocks you're tracking) ──
    watchlist_result = deepseek_watchlist_trade_plan(portfolio, tv_data, market, MONTHLY_INVESTMENT_EGP)
    watchlist_output = display_watchlist_trade_plans(watchlist_result)

    # ── Recommendation Engine (Option B: pick → fetch live → refine) ──
    rec_output = ""
    rec_result = {"tokens_used": 0, "cost_usd": 0}
    pick_result = {"tokens_used": 0, "cost_usd": 0}
    if RECOMMEND_NEW_STOCKS:
        pick_result = deepseek_pick_recommendations(
            portfolio, market, NUM_MEDIUM_TERM_PICKS, NUM_LONG_TERM_PICKS, PREFER_DIVERSIFICATION
        )
        if "error" not in pick_result:
            # Fetch live TradingView data for the new picks
            new_syms = [p["symbol"] for p in pick_result.get("medium_term", [])] + \
                       [p["symbol"] for p in pick_result.get("long_term", [])]
            new_syms = [s for s in new_syms if s and s not in tv_data]
            if new_syms:
                print(f"\n📊 Fetching live data for {len(new_syms)} recommended symbols: {new_syms}")
                new_tv = fetch_tradingview_data(new_syms)
                tv_data.update(new_tv)
                print()

            rec_result = deepseek_refine_recommendations(
                pick_result, tv_data, market, MONTHLY_INVESTMENT_EGP
            )
            rec_output = display_recommendations(rec_result, pick_result)
        else:
            print(f"   ⚠️ Skipping recommendations — pick step failed: {pick_result['error']}")

    arabic_result   = deepseek_arabic_report(portfolio, market, tv_data, news, ai_result, predictions, MONTHLY_INVESTMENT_EGP)
    arabic_output   = display_arabic_report(arabic_result)

    save_report(report_output, pred_output, arabic_output,
                recommendations_text=rec_output, watchlist_text=watchlist_output)

    # Print total cost summary
    t1 = ai_result.get('tokens_used', 0)
    t2 = predictions.get('tokens_used', 0)
    t3 = arabic_result.get('tokens_used', 0)
    t4 = pick_result.get('tokens_used', 0)
    t5 = rec_result.get('tokens_used', 0)
    t6 = watchlist_result.get('tokens_used', 0)
    c1 = ai_result.get('cost_usd', 0)
    c2 = predictions.get('cost_usd', 0)
    c3 = arabic_result.get('cost_usd', 0)
    c4 = pick_result.get('cost_usd', 0)
    c5 = rec_result.get('cost_usd', 0)
    c6 = watchlist_result.get('cost_usd', 0)
    total_tokens = t1 + t2 + t3 + t4 + t5 + t6
    total_cost   = c1 + c2 + c3 + c4 + c5 + c6
    total_egp    = total_cost * float(market.get('usd_egp', {}).get('rate', '50').replace(',','')) \
                   if str(market.get('usd_egp', {}).get('rate', 'N/A')).replace('.','').replace(',','').isdigit() \
                   else total_cost * 50  # fallback EGP rate

    # Build a prominent banner
    W = 70
    cost_lines = []
    cost_lines.append("")
    cost_lines.append("╔" + "═"*(W-2) + "╗")
    cost_lines.append(("║" + "  💰  TOTAL TOKEN & COST REPORT".ljust(W-2) + "║"))
    cost_lines.append("╠" + "═"*(W-2) + "╣")
    cost_lines.append(f"║  {'Section':<22} {'Tokens':>10}     {'Cost (USD)':>12}   ║".ljust(W) + "║"[len("║"):] if False else f"║  {'Section':<22} {'Tokens':>10}     {'Cost (USD)':>12}   ║")
    cost_lines.append("║" + "─"*(W-2) + "║")
    cost_lines.append(f"║  {'English Analysis':<22} {t1:>10,}     ${c1:>11.5f}   ║")
    cost_lines.append(f"║  {'Price Predictions':<22} {t2:>10,}     ${c2:>11.5f}   ║")
    if not watchlist_result.get('skipped'):
        cost_lines.append(f"║  {'Watchlist Plans':<22} {t6:>10,}     ${c6:>11.5f}   ║")
    if RECOMMEND_NEW_STOCKS and t4 > 0:
        cost_lines.append(f"║  {'Recommend. Picker':<22} {t4:>10,}     ${c4:>11.5f}   ║")
        cost_lines.append(f"║  {'Recommend. Refiner':<22} {t5:>10,}     ${c5:>11.5f}   ║")
    cost_lines.append(f"║  {'Arabic Report':<22} {t3:>10,}     ${c3:>11.5f}   ║")
    cost_lines.append("║" + "─"*(W-2) + "║")
    cost_lines.append(f"║  {'TOTAL':<22} {total_tokens:>10,}     ${total_cost:>11.5f}   ║")
    cost_lines.append(f"║  {'≈ in EGP':<22} {' ':>10}     {total_egp:>10.2f} EGP  ║")
    cost_lines.append("╚" + "═"*(W-2) + "╝")
    cost_lines.append("")

    cost_summary = "\n".join(cost_lines)
    print(cost_summary)

    # Append to saved report file
    SCRIPT_DIR = os.getcwd()
    saved_files = sorted(glob.glob(os.path.join(SCRIPT_DIR, "portfolio_report_*.txt")), key=os.path.getmtime)
    if saved_files:
        with open(saved_files[-1], "a", encoding="utf-8") as f:
            f.write("\n\n")
            f.write(cost_summary)

if __name__ == "__main__":
    main()