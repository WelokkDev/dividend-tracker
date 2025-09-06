import yfinance as yf
import json
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv
import os
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Create a new workbook
wb = Workbook()
ws = wb.active
ws.title = "Dividends"

load_dotenv()

class App():
    def __init__(self, root):
        self.root = root
        self.root.title("Dividend Tracker")
        self.root.geometry("600x600")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.dataManager = DividendDataManager()
        self.setup_ui()

        

    def setup_ui(self):
        # Frame for ticker list
        list_frame = ttk.LabelFrame(self.root, text="Current Stock Tickers")
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.ticker_listbox = tk.Listbox(list_frame, height=10, font=("Segoe UI", 12))
        self.ticker_listbox.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.ticker_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.ticker_listbox.config(yscrollcommand=scrollbar.set)
        # Populate listbox
        for ticker in self.dataManager.tickers:
           self.ticker_listbox.insert(tk.END, ticker.symbol)

        # Frame for adding/removing tickers
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=10)

        self.ticker_entry = ttk.Entry(control_frame, font=("Segoe UI", 12))
        self.ticker_entry.pack(side="left", fill="x", expand=True, padx=(0,10))

        add_button = ttk.Button(control_frame, text="Add Ticker", command=self.add_ticker)
        add_button.pack(side="left", padx=(0,10))

        remove_button = ttk.Button(control_frame, text="Remove Selected", command=self.update_ticker_list)
        remove_button.pack(side="left")

        # Export/update buttons at bottom
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=10, pady=10)

        update_button = ttk.Button(bottom_frame, text="Rebuild Excel", command=self.build_excel)
        update_button.pack(side="left", padx=(0,10))

        export_button = ttk.Button(bottom_frame, text="Open Excel", command=self.open_excel)
        export_button.pack(side="left")
        

    
    def update_ticker_list(self):
        selected = self.ticker_listbox.curselection()
        if selected:
            symbol = self.ticker_listbox.get(selected[0])
            self.ticker_listbox.delete(selected[0])
            self.dataManager.remove_ticker(symbol)

    def add_ticker(self):
        new_ticker = self.ticker_entry.get().strip().upper()
        if new_ticker and new_ticker not in self.dataManager.tickers:
            self.dataManager.add_ticker(new_ticker)
            self.ticker_listbox.insert(tk.END, new_ticker)
        self.ticker_entry.delete(0, tk.END)

    def build_excel(self):
        try:
            path = Path("dividend_data.json")
            if not path.is_file():
                raise FileNotFoundError("dividend_data.json not found")

            # Load JSON data
            with open(path, "r", encoding="utf-8") as f:
                dividend_data = json.load(f)

            # Create workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Dividends"

            # Headers - expanded to include all Alpha Vantage fields
            headers = ["Ex-Date", "Declaration Date", "Record Date", "Payment Date", "Ticker", "Currency", "Dividend"]
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            alignment = Alignment(horizontal="center", vertical="center")

            for col_num, title in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num, value=title)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = alignment

            # Column widths
            ws.column_dimensions["A"].width = 15  # Ex-Date
            ws.column_dimensions["B"].width = 15  # Declaration Date
            ws.column_dimensions["C"].width = 15  # Record Date
            ws.column_dimensions["D"].width = 15  # Payment Date
            ws.column_dimensions["E"].width = 12  # Ticker
            ws.column_dimensions["F"].width = 10  # Currency
            ws.column_dimensions["G"].width = 12  # Dividend

            # Fill rows
            row_num = 2
            for entry in dividend_data:
                ticker = entry.get("ticker")
                currency = entry.get("currency", "USD")
                for div in entry.get("dividends", []):
                    # Handle both TSX (ex_date) and US (ex_dividend_date) formats
                    ex_date = div.get("ex_date") or div.get("ex_dividend_date")
                    declaration_date = div.get("declaration_date")
                    record_date = div.get("record_date")
                    payment_date = div.get("payment_date")
                    amount = div.get("amount")
                    
                    # Convert amount to float for proper Excel number formatting
                    try:
                        amount = float(amount) if amount else 0
                    except (ValueError, TypeError):
                        amount = 0
                    
                    # Convert "None" strings to empty cells for better Excel display
                    if declaration_date == "None":
                        declaration_date = ""
                    if record_date == "None":
                        record_date = ""
                    if payment_date == "None":
                        payment_date = ""
                    
                    ws.cell(row=row_num, column=1, value=ex_date)
                    ws.cell(row=row_num, column=2, value=declaration_date)
                    ws.cell(row=row_num, column=3, value=record_date)
                    ws.cell(row=row_num, column=4, value=payment_date)
                    ws.cell(row=row_num, column=5, value=ticker)
                    ws.cell(row=row_num, column=6, value=currency)
                    ws.cell(row=row_num, column=7, value=amount)
                    row_num += 1

            # Save workbook
            output_path = "dividends-sheet.xlsx"
            wb.save(output_path)
            print(f"Excel file saved to {output_path}")
            print("Excel export completed successfully!")
            return output_path
            
        except Exception as e:
            print(f"Error building Excel file: {e}")
            import traceback
            traceback.print_exc()  # This will show the full error details
            raise
        

    def open_excel(self):
        """Open the Excel file with the system's default program"""
        import subprocess
        import sys
        
        excel_path = "dividends-sheet.xlsx"
        
        # Check if file exists
        if not os.path.exists(excel_path):
            print("Excel file not found. Please build the Excel file first.")
            return
        
        try:
            # Open file with system default program
            if sys.platform == "win32":
                os.startfile(excel_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", excel_path])
            else:  # Linux and others
                subprocess.run(["xdg-open", excel_path])
            print(f"Opening {excel_path}...")
        except Exception as e:
            print(f"Error opening Excel file: {e}")


class DividendDataManager:
    # Gather all ticker data necessary for excel and json...
    def __init__(self):
        
        self.tickers = []
        path = Path("dividend_data.json")
        if path.is_file():
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data:
                    self.tickers.append(StockTicker(item.get("ticker")))




    def add_ticker(self, symbol):
        ticker = StockTicker(symbol, True)
        self.tickers.append(ticker)
        self.add_to_json(ticker)

    def add_to_json(self, ticker):
        path = Path("dividend_data.json")
        if path.is_file():
            with open(path, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    data = []
        else:
            data = []

        data.append(ticker.data)
                
        with open(path, "w") as f:
            json.dump(data, f, indent=2)

    def remove_ticker(self, symbol):
        self.tickers = [t for t in self.tickers if t.symbol != symbol]
        print(symbol)
        print(self.tickers)
        self.remove_from_json(symbol)

    def remove_from_json(self, symbol):
        path = Path("dividend_data.json")
        
        if path.is_file():
            with open(path, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    data = []

        # filter out the ticker
        data = [entry for entry in data if entry.get("ticker") != symbol]

        # write back updated data
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        



class StockTicker:
    def __init__(self, symbol, new=False):
        self.symbol = symbol.strip().upper()
        self.data = {
            "ticker": self.symbol,
            "currency": "USD",
            "dividends": []
        }
        if new == True:
            if self.symbol.endswith((".TO", ".V", ".CN", ":CA")):
                self.fetch_tsx_dividends()
            else:
                self.fetch_dividends()

    def fetch_tsx_dividends(self):
        self.ticker = yf.Ticker(self.symbol)
        self.data["currency"] = "CAD"
        # Fetch the last 3 months of dividend data
        div_series = self.ticker.get_dividends(period="3mo")

        # Store as list of dicts
        self.data["dividends"] = [
            {"ex_dividend_date": date.strftime("%Y-%m-%d"), "amount": float(amount)}
            for date, amount in div_series.items()
        ]
        print("Dividends for TSX stocks are being fetched!")
        print(self.data)

    def fetch_dividends(self):
        print("Dividends for US stocks are being fetched!")
        
        # Get Alpha Vantage API key from environment
        api_key = os.getenv("ALPHA_VANTAGE_API_KEY")
        if not api_key:
            print("Warning: ALPHA_VANTAGE_API_KEY not found in environment variables")
            return
        
        # Alpha Vantage API URL for dividend data
        url = f'https://www.alphavantage.co/query?function=DIVIDENDS&symbol={self.symbol}&apikey={api_key}'
        
        try:
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()
            
            # Check if we got valid data
            if "data" not in data:
                print(f"Error: No dividend data found for {self.symbol}")
                print(f"API response: {data}")
                return
            
            # Filter to last 6 months (approximately 6 months)
            from datetime import datetime, timedelta
            cutoff_date = datetime.now() - timedelta(days=180)  # 6 months ago
            
            filtered_dividends = []
            for dividend in data["data"]:
                ex_date_str = dividend.get("ex_dividend_date")
                if ex_date_str and ex_date_str != "None":
                    try:
                        ex_date = datetime.strptime(ex_date_str, "%Y-%m-%d")
                        if ex_date >= cutoff_date:
                            # Store the complete dividend object as-is
                            filtered_dividends.append(dividend)
                    except ValueError:
                        # Skip invalid dates
                        continue
            
            # Sort by date (newest first)
            filtered_dividends.sort(key=lambda x: x.get("ex_dividend_date", ""), reverse=True)
            
            self.data["dividends"] = filtered_dividends
            print(f"Fetched {len(filtered_dividends)} dividend records for {self.symbol}")
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching dividend data for {self.symbol}: {e}")
        except Exception as e:
            print(f"Unexpected error for {self.symbol}: {e}")
    



if __name__ == "__main__":
    import tkinter as tk
    from tkinter import ttk
    import ctypes
    

   
    # Optional: make app DPI aware on Windows
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

  

    root = tk.Tk()  # Create the main Tkinter window
    app = App(root)  # Create your app instance



    root.mainloop()  # Start Tkinter's main loop
