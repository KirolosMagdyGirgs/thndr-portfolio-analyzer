# Thndr Portfolio Analyzer 📊

An automated tool that scrapes your Thndr investment portfolio, enriches it 
with live TradingView technical data, and generates a DeepSeek AI-powered 
analysis report in English and Arabic.

## What it does
- Scrapes your live portfolio from web.thndr.app (Scraper_app.py)
- Fetches live RSI, MACD, P/E, ROE and 30+ technical indicators from TradingView
- Generates a detailed report with DCA recommendations, price predictions, and 
  Arabic translation via DeepSeek AI

## Setup
1. pip install -r requirements.txt
2. Copy .env.example to .env and add your DEEPSEEK_API_KEY and GROQ_API_KEY
3. Run Scraper_app.py first, then full_app.py
   Or use run_thnder_analysis.bat (Windows)

## Notes
- Requires a Thndr account and DeepSeek API key
- Not financial advice
