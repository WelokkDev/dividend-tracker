import yfinance as yf
import json
import pandas as pd
from pathlib import Path

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

        update_button = ttk.Button(bottom_frame, text="Rebuild Excel", command=lambda: print("Fetching dividends..."))
        update_button.pack(side="left", padx=(0,10))

        export_button = ttk.Button(bottom_frame, text="Open Excel", command=lambda: print("Exporting..."))
        export_button.pack(side="left")
        

    
    def update_ticker_list(self):
        selected = self.ticker_listbox.curselection()
        if selected:
            self.ticker_listbox.delete(selected[0])

    def add_ticker(self):
        new_ticker = self.ticker_entry.get().strip().upper()
        if new_ticker and new_ticker not in self.dataManager.tickers:
            self.dataManager.addTicker(new_ticker)
            self.ticker_listbox.insert(tk.END, new_ticker)
        self.ticker_entry.delete(0, tk.END)

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




    def addTicker(self, symbol):
        ticker = StockTicker(symbol)
        ticker.fetch_dividends()
        self.tickers.append(ticker)


class StockTicker:
    def __init__(self, symbol):
        self.symbol = symbol

    def fetch_dividends(self):
        print("Testing, dividends fetched for", self.symbol)

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
