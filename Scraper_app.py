import asyncio
import os
from dotenv import load_dotenv
from config import BASE_URL
from utils.scraper_utils import fetch_investments, save_to_excel

load_dotenv()

# Save to the current working directory.
# The .bat sets cwd to the Desktop output folder before launching us.
SAVE_FOLDER = os.getcwd()


async def main():
    investments = await fetch_investments(BASE_URL)

    if investments:
        print("\n📊 Your Investments:")
        print(f"{'Asset':<10} {'Class':<15} {'Units':>8} {'Cost/Unit':>12} {'Cur Price':>12} {'Mkt Value':>12} {'Daily Chg':>22} {'Unreal. Ret':>22}")
        print("-" * 120)
        for inv in investments:
            print(f"{inv['Asset']:<10} {inv['Asset Class']:<15} {inv['Units Owned']:>8} {inv['Cost Per Unit']:>12} {inv['Current Price']:>12} {inv['Market Value']:>12} {inv['Daily Change']:>22} {inv['Unrealized Return']:>22}")

        save_to_excel(investments, SAVE_FOLDER)
    else:
        print("No investments found.")


if __name__ == "__main__":
    asyncio.run(main())