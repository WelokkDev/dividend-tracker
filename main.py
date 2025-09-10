import yfinance as yf
import json
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv
import os
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import tkinter as tk
from tkinter import ttk, messagebox
import re
import time
from datetime import datetime

# Create a new workbook
wb = Workbook()
ws = wb.active
ws.title = "Dividends"

load_dotenv()

class ValidationUtils:
    """Utility class for input validation and error handling"""
    
    @staticmethod
    def validate_ticker_symbol(symbol):
        """Validate ticker symbol format"""
        if not symbol or not isinstance(symbol, str):
            return False, "Ticker symbol cannot be empty"
        
        symbol = symbol.strip().upper()
        
        # Basic format validation
        if len(symbol) < 1 or len(symbol) > 10:
            return False, "Ticker symbol must be 1-10 characters long"
        
        # Check for valid characters (letters, numbers, dots, colons)
        if not re.match(r'^[A-Z0-9\.:]+$', symbol):
            return False, "Ticker symbol can only contain letters, numbers, dots, and colons"
        
        # Check for common invalid patterns
        if symbol.startswith('.') or symbol.endswith('.'):
            return False, "Ticker symbol cannot start or end with a dot"
        
        if '..' in symbol:
            return False, "Ticker symbol cannot contain consecutive dots"
        
        return True, symbol
    
    @staticmethod
    def is_duplicate_ticker(symbol, existing_tickers):
        """Check if ticker already exists in portfolio"""
        symbol = symbol.strip().upper()
        for ticker in existing_tickers:
            if hasattr(ticker, 'symbol') and ticker.symbol == symbol:
                return True
        return False

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

        update_button = ttk.Button(bottom_frame, text="Build/Rebuild Excel", command=self.build_excel)
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
        """Add a new ticker with validation and error handling"""
        try:
            new_ticker = self.ticker_entry.get().strip()
            
            # Validate ticker symbol
            is_valid, result = ValidationUtils.validate_ticker_symbol(new_ticker)
            if not is_valid:
                messagebox.showerror("Invalid Ticker", f"Error: {result}")
                return
            
            new_ticker = result  # Use the validated/cleaned symbol
            
            # Check for duplicates
            if ValidationUtils.is_duplicate_ticker(new_ticker, self.dataManager.tickers):
                messagebox.showwarning("Duplicate Ticker", f"'{new_ticker}' is already in your portfolio")
                return
            
            # Show loading state
            self.show_loading_state()
            
            # Add ticker with error handling
            success = self.dataManager.add_ticker(new_ticker)
            
            if success:
                self.ticker_listbox.insert(tk.END, new_ticker)
                messagebox.showinfo("Success", f"Successfully added {new_ticker} to your portfolio")
            else:
                messagebox.showerror("Error", f"Failed to add {new_ticker}. Please check the ticker symbol and try again.")
                
        except Exception as e:
            messagebox.showerror("Unexpected Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.ticker_entry.delete(0, tk.END)
            self.hide_loading_state()
    
    def show_loading_state(self):
        """Show loading indicator during API calls"""
        self.ticker_entry.config(state='disabled')
        # Add a progress bar / loading label here
    
    def hide_loading_state(self):
        """Hide loading indicator after API calls"""
        self.ticker_entry.config(state='normal')

    def build_excel(self):
        """Build Excel file with comprehensive error handling"""
        try:
            path = Path("dividend_data.json")
            if not path.is_file():
                messagebox.showerror("File Not Found", "dividend_data.json not found. Please add some tickers first.")
                return

            # Load JSON data
            try:
                with open(path, "r", encoding="utf-8") as f:
                    dividend_data = json.load(f)
            except json.JSONDecodeError as e:
                messagebox.showerror("Data Error", f"Error reading portfolio data: {str(e)}")
                return
            except Exception as e:
                messagebox.showerror("File Error", f"Error opening portfolio file: {str(e)}")
                return

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
            try:
                wb.save(output_path)
                messagebox.showinfo("Success", f"Excel file saved successfully!\nLocation: {output_path}")
                return output_path
            except PermissionError:
                messagebox.showerror("Permission Error", 
                    "Cannot save Excel file. Please close the file if it's open in Excel and try again.")
                return None
            except Exception as e:
                messagebox.showerror("Save Error", f"Error saving Excel file: {str(e)}")
                return None
            
        except Exception as e:
            messagebox.showerror("Excel Export Error", f"An error occurred while building the Excel file:\n{str(e)}")
            return None
        

    def open_excel(self):
        """Open the Excel file with the system's default program"""
        import subprocess
        import sys
        
        excel_path = "dividends-sheet.xlsx"
        
        # Check if file exists
        if not os.path.exists(excel_path):
            messagebox.showwarning("File Not Found", 
                "Excel file not found. Please click 'Rebuild Excel' first to generate the file.")
            return
        
        try:
            # Open file with system default program
            if sys.platform == "win32":
                os.startfile(excel_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", excel_path])
            else:  # Linux and others
                subprocess.run(["xdg-open", excel_path])
            messagebox.showinfo("Success", f"Opening {excel_path}...")
        except FileNotFoundError:
            messagebox.showerror("Application Not Found", 
                "No application found to open Excel files. Please install Excel or a compatible spreadsheet application.")
        except Exception as e:
            messagebox.showerror("Error", f"Error opening Excel file: {str(e)}")


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
        """Add a new ticker"""
        try:
            ticker = StockTicker(symbol, True)
            if ticker.data.get("dividends") is not None:  # Check if data was fetched successfully
                self.tickers.append(ticker)
                self.add_to_json(ticker)
                return True
            else:
                return False
        except Exception as e:
            print(f"Error adding ticker {symbol}: {e}")
            return False

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
        """Fetch US stock dividends"""
        from datetime import timedelta
        print("Dividends for US stocks are being fetched!")
        
        # Get Alpha Vantage API key from environment
        api_key = os.getenv("ALPHA_VANTAGE_API_KEY")
        if not api_key:
            print("Warning: ALPHA_VANTAGE_API_KEY not found in environment variables")
            self.data["dividends"] = []
            return
        
        # Alpha Vantage API URL for dividend data
        url = f'https://www.alphavantage.co/query?function=DIVIDENDS&symbol={self.symbol}&apikey={api_key}'
        
        # Retry logic for network failures
        max_retries = 3
        retry_delay = 1  # seconds
        
        for attempt in range(max_retries):
            try:
                print(f"Fetching data for {self.symbol} (attempt {attempt + 1}/{max_retries})")
                
                # Make API request with timeout
                response = requests.get(url, timeout=30)
                response.raise_for_status()
                data = response.json()
                
                # Check for API errors
                if "Error Message" in data:
                    print(f"API Error for {self.symbol}: {data['Error Message']}")
                    self.data["dividends"] = []
                    return
                
                if "Note" in data:
                    print(f"API Rate Limit for {self.symbol}: {data['Note']}")
                    if attempt < max_retries - 1:
                        print(f"Waiting {retry_delay * 2} seconds before retry...")
                        time.sleep(retry_delay * 2)
                        retry_delay *= 2  # Exponential backoff
                        continue
                    else:
                        print("Max retries reached. Using empty data.")
                        self.data["dividends"] = []
                        return
                
                # Check if we got valid data
                if "data" not in data or not data["data"]:
                    print(f"No dividend data found for {self.symbol}")
                    self.data["dividends"] = []
                    return
                
                # Filter to last 6 months
                cutoff_date = datetime.now() - timedelta(days=180)
                
                filtered_dividends = []
                for dividend in data["data"]:
                    ex_date_str = dividend.get("ex_dividend_date")
                    if ex_date_str and ex_date_str != "None":
                        try:
                            ex_date = datetime.strptime(ex_date_str, "%Y-%m-%d")
                            if ex_date >= cutoff_date:
                                filtered_dividends.append(dividend)
                        except ValueError:
                            continue
                
                # Sort by date (newest first)
                filtered_dividends.sort(key=lambda x: x.get("ex_dividend_date", ""), reverse=True)
                
                self.data["dividends"] = filtered_dividends
                print(f"Successfully fetched {len(filtered_dividends)} dividend records for {self.symbol}")
                return  # Success, exit retry loop
                
            except requests.exceptions.Timeout:
                print(f"Timeout error for {self.symbol} (attempt {attempt + 1})")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    print("Max retries reached due to timeout.")
                    self.data["dividends"] = []
                    
            except requests.exceptions.ConnectionError:
                print(f"Connection error for {self.symbol} (attempt {attempt + 1})")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    print("Max retries reached due to connection error.")
                    self.data["dividends"] = []
                    
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 429:  # Rate limit
                    print(f"Rate limit exceeded for {self.symbol}")
                    if attempt < max_retries - 1:
                        wait_time = retry_delay * 5  # Longer wait for rate limits
                        print(f"Waiting {wait_time} seconds before retry...")
                        time.sleep(wait_time)
                        retry_delay *= 2
                    else:
                        print("Max retries reached due to rate limiting.")
                        self.data["dividends"] = []
                else:
                    print(f"HTTP error for {self.symbol}: {e}")
                    self.data["dividends"] = []
                    return
                    
            except Exception as e:
                print(f"Unexpected error for {self.symbol}: {e}")
                self.data["dividends"] = []
                return
    



if __name__ == "__main__":
    import ctypes
    

   
    # Optional: make app DPI aware on Windows
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

  

    root = tk.Tk()  # Create the main Tkinter window
    app = App(root)  # Create your app instance



    root.mainloop()  # Start Tkinter's main loop
