from gui import run_app   # Tkinter GUI
from tracker import Tracker  # Your merged Excel+logic handler

def main():
    tracker = Tracker("dividends.xlsx", "backup.json")
    run_app(tracker)

if __name__ == "__main__":
    main()