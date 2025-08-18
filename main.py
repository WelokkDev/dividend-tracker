import yfinance as yf
import json
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv
import os
from polygon import RESTClient
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
        self.root.geometry("600x500")

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

        export_button = ttk.Button(bottom_frame, text="Open Excel", command=lambda: print("Exporting..."))
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

        # Headers
        headers = ["Ex-Date", "Ticker", "Currency", "Dividend"]
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        alignment = Alignment(horizontal="center", vertical="center")

        for col_num, title in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=title)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = alignment

        # Optional: column widths
        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 12

        # Fill rows
        row_num = 2
        for entry in dividend_data:
            ticker = entry.get("ticker")
            currency = entry.get("currency", "USD")
            for div in entry.get("dividends", []):
                ex_date = div.get("ex_date") or div.get("ex_dividend_date")
                amount = div.get("amount")
                ws.cell(row=row_num, column=1, value=ex_date)
                ws.cell(row=row_num, column=2, value=ticker)
                ws.cell(row=row_num, column=3, value=currency)
                ws.cell(row=row_num, column=4, value=amount)
                row_num += 1

        # Save workbook
        output_path = "dividends-sheet.xlsx"
        wb.save(output_path)
        print(f"Excel file saved to {output_path}")
        return output_path


class DividendDataManager:
    # Gather all ticker data necessary for excel and json...
    def __init__(self):
        self.client = RESTClient(os.getenv("POLYGON_API_KEY"))
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
