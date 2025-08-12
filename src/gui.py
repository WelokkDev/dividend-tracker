import tkinter as tk
from tkinter import ttk
import ctypes
 
ctypes.windll.shcore.SetProcessDpiAwareness(1)


def update_ticker_list():
    selected = ticker_listbox.curselection()
    if selected:
        ticker_listbox.delete(selected[0])

def add_ticker():
    new_ticker = ticker_entry.get().strip().upper()
    if new_ticker and new_ticker not in tickers:
        tickers.append(new_ticker)
        ticker_listbox.insert(tk.END, new_ticker)
    ticker_entry.delete(0, tk.END)

# Sample initial tickers
tickers = ["AAPL", "MSFT", "GOOG"]

root = tk.Tk()
root.title("Dividend Tracker")
root.geometry("600x500")

style = ttk.Style()
style.theme_use("clam") 

# Frame for ticker list
list_frame = ttk.LabelFrame(root, text="Current Stock Tickers")
list_frame.pack(fill="both", expand=True, padx=10, pady=10)

ticker_listbox = tk.Listbox(list_frame, height=10, font=("Segoe UI", 12))
ticker_listbox.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)

# Add scrollbar
scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=ticker_listbox.yview)
scrollbar.pack(side="right", fill="y")
ticker_listbox.config(yscrollcommand=scrollbar.set)

# Populate listbox
for ticker in tickers:
    ticker_listbox.insert(tk.END, ticker)

# Frame for adding/removing tickers
control_frame = ttk.Frame(root)
control_frame.pack(fill="x", padx=10, pady=10)

ticker_entry = ttk.Entry(control_frame, font=("Segoe UI", 12))
ticker_entry.pack(side="left", fill="x", expand=True, padx=(0,10))

add_button = ttk.Button(control_frame, text="Add Ticker", command=add_ticker)
add_button.pack(side="left", padx=(0,10))

remove_button = ttk.Button(control_frame, text="Remove Selected", command=update_ticker_list)
remove_button.pack(side="left")

# Export/update buttons at bottom
bottom_frame = ttk.Frame(root)
bottom_frame.pack(fill="x", padx=10, pady=10)

update_button = ttk.Button(bottom_frame, text="Rebuild Excel", command=lambda: print("Fetching dividends..."))
update_button.pack(side="left", padx=(0,10))

export_button = ttk.Button(bottom_frame, text="Open Excel", command=lambda: print("Exporting..."))
export_button.pack(side="left")

root.mainloop()
