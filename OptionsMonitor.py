import tkinter as tk
from tkinter import ttk, messagebox
import yfinance as yf
import csv
import os
from datetime import datetime, time
import pytz
import math
import time as time_module
import threading

class OptionsMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Options Monitor - Ver 1.8")
        self.data_file = r"C:\ProgramData\ShadowWhisperer\OptionsMonitor\data.csv"
        self.data = self.load_data()
        self.price_cache = {}
        self.sort_reverse = {}
        self.last_market_status = None
        self.refresh_interval = None
        self.last_interval = "15 Mins"
        self.last_updated_label = None
        self.DevMode = 0
        self.current_sort_col = None
        self.current_sort_reverse = False
        self.setup_gui()
        self.populate_treeview()
        threading.Thread(target=self.fetch_initial_prices, daemon=True).start()

    def load_data(self):
        os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
        try:
            with open(self.data_file, 'r', newline='') as f:
                reader = csv.reader(f)
                data = []
                for row in reader:
                    if not row or row[0] == "Ticker":  # skip header if present
                        continue
                    try:
                        if len(row) == 6:
                            data.append([row[0], row[1], row[2], int(row[3]), float(row[4]), float(row[5])])
                        elif len(row) == 5:
                            data.append([row[0], row[1], row[2], int(row[3]), 0.0, float(row[4])])
                    except ValueError:
                        continue
                return data
        except FileNotFoundError:
            return []

    def save_data(self):
        # Write header then data rows (header will be ignored on load)
        os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
        with open(self.data_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Ticker", "Ends", "Option", "Contracts", "Premium", "Strike"])
            for row in self.data:
                writer.writerow([
                    row[0],
                    row[1],
                    row[2],
                    str(row[3]),
                    str(row[4]),
                    str(row[5])
                ])

    def setup_gui(self):
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        cols = ("Ticker", "Ends", "Option", "Contracts", "Premium", "Strike", "Current", "Diff", "Outcome", "Value")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings")
        for col in cols:
            width = 20
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c), anchor="w")
            self.tree.column(col, width=width, anchor="w")
            self.sort_reverse[col] = False

        self.tree.tag_configure('oddrow', background='#f0f0f0')
        self.tree.tag_configure('evenrow', background='#ffffff')
        self.tree.tag_configure('redrow', background='#ffcccc')
        self.tree.tag_configure('green_diff', foreground='green')
        self.tree.tag_configure('red_diff', foreground='red')
        self.tree.grid(row=0, column=0, pady=2, sticky="nsew")
        self.tree.bind("<Double-1>", self.on_double_click)

        button_frame = tk.Frame(self.root, height=30)
        button_frame.grid(row=1, column=0, pady=2, sticky="ew")
        self.refresh_button = tk.Button(button_frame, text="Refresh", command=self.refresh_data, width=10)
        self.refresh_button.pack(side="left", padx=(15, 5))
        tk.Button(button_frame, text="Add", command=self.add_option, width=10).pack(side="left", padx=(10, 5))
        tk.Button(button_frame, text="Remove", command=self.remove_selected, width=10).pack(side="left", padx=(10, 5))
        tk.Button(button_frame, text="Remove All", command=self.confirm_remove_all, width=10).pack(side="left", padx=(10, 5))

        self.last_updated_label = tk.Label(button_frame, text="")
        self.last_updated_label.pack(side="right", padx=5)

        self.interval_var = tk.StringVar(value="5 Mins")
        self.interval_combo = ttk.Combobox(
            button_frame,
            textvariable=self.interval_var,
            values=["Don't Update", "5 Mins", "10 Mins", "15 Mins", "30 Mins", "1 Hour", "2 Hours"],
            width=16, state="readonly"
        )
        self.interval_combo.pack(side="right", padx=(0, 0))
        self.interval_combo.bind("<<ComboboxSelected>>", self.schedule_refresh)

        style = ttk.Style()
        style.configure("Red.TCombobox", foreground="red")
        style.map("Red.TCombobox", foreground=[("disabled", "red")])

        self.status_frame = tk.Frame(self.root)
        self.status_frame.grid(row=2, column=0, sticky="ew")

        self.update_market_status()
        self.schedule_refresh()

    def is_market_open(self):
        if self.DevMode == 1:
            return True
        now = datetime.now(pytz.timezone('US/Eastern'))
        return now.weekday() < 5 and time(9, 30) <= now.time() <= time(16, 0)

    def update_market_status(self):
        now = datetime.now(pytz.timezone('US/Eastern'))
        is_open = self.is_market_open()
        market_open_time = datetime.combine(now.date(), time(9, 30)).replace(tzinfo=pytz.timezone('US/Eastern'))
        time_to_open = (market_open_time - now).total_seconds()
        is_near_open = 0 <= time_to_open <= 300

        if self.last_market_status != is_open:
            self.refresh_button.config(state="normal" if is_open else "disabled")
            if is_open:
                self.interval_combo.config(values=["Don't Update", "5 Mins", "10 Mins", "15 Mins", "30 Mins", "1 Hour", "2 Hours"], state="readonly", style="TCombobox")
                self.interval_var.set(self.last_interval)
                self.schedule_refresh()
            else:
                self.interval_combo.config(values=["Markets Closed"], state="disabled", style="Red.TCombobox")
                self.interval_var.set("Markets Closed")
                if self.refresh_interval is not None:
                    self.root.after_cancel(self.refresh_interval)
                    self.refresh_interval = None
            if (is_open and self.last_market_status is False) or (not is_open and self.last_market_status is True):
                self.refresh_data()
            self.last_market_status = is_open

        self.root.after(30000 if is_near_open else 10000, self.update_market_status)

    def schedule_refresh(self, event=None):
        if self.refresh_interval is not None:
            self.root.after_cancel(self.refresh_interval)
            self.refresh_interval = None

        interval = self.interval_var.get()
        if interval == "Don't Update" or interval == "Markets Closed":
            return

        self.last_interval = interval
        try:
            if "Hour" in interval:
                ms = int(interval.replace(" Hour", "").replace("s", "")) * 60 * 60 * 1000
            else:
                ms = int(interval.replace(" Mins", "")) * 60 * 1000
        except ValueError:
            return

        def refresh_if_open():
            if self.is_market_open():
                self.refresh_data()
            self.refresh_interval = self.root.after(ms, refresh_if_open)

        self.refresh_interval = self.root.after(ms, refresh_if_open)

    def add_option(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Add")
        add_window.resizable(False, False)

        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        window_width = 260
        window_height = 280
        x = root_x + (root_width - window_width) // 2
        y = root_y + (root_height - window_height) // 2
        add_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        form_frame = tk.Frame(add_window)
        form_frame.pack(pady=6, padx=6)

        tk.Label(form_frame, text="Ticker").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ticker_entry = tk.Entry(form_frame)
        ticker_entry.grid(row=0, column=1, pady=4)

        tk.Label(form_frame, text="Option").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        call_put_var = tk.StringVar(value="Call")
        option_frame = tk.Frame(form_frame)
        tk.Radiobutton(option_frame, text="Call", variable=call_put_var, value="Call").pack(side="left")
        tk.Radiobutton(option_frame, text="Put", variable=call_put_var, value="Put").pack(side="left")
        option_frame.grid(row=1, column=1, pady=4)

        tk.Label(form_frame, text="Contracts").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        contracts_entry = tk.Entry(form_frame)
        contracts_entry.grid(row=2, column=1, pady=4)

        tk.Label(form_frame, text="Premium").grid(row=3, column=0, sticky="e", padx=4, pady=4)
        premium_entry = tk.Entry(form_frame)
        premium_entry.grid(row=3, column=1, pady=4)

        tk.Label(form_frame, text="Strike").grid(row=4, column=0, sticky="e", padx=4, pady=4)
        strike_entry = tk.Entry(form_frame)
        strike_entry.grid(row=4, column=1, pady=4)

        tk.Label(form_frame, text="Close (M/D)").grid(row=5, column=0, sticky="e", padx=4, pady=4)
        close_date_entry = tk.Entry(form_frame)
        close_date_entry.grid(row=5, column=1, pady=4)

        tk.Button(add_window, text="Add Option", command=lambda: self._submit_option(
            ticker_entry, call_put_var, contracts_entry, premium_entry, strike_entry, close_date_entry, add_window), width=12
        ).pack(pady=8)

    def _submit_option(self, ticker_entry, call_put_var, contracts_entry, premium_entry, strike_entry, close_date_entry, window):
        ticker = ticker_entry.get().upper().strip()
        call_put = call_put_var.get().capitalize()
        contracts = contracts_entry.get().strip()
        premium = premium_entry.get().strip()
        strike_price = strike_entry.get().strip()
        close_date = close_date_entry.get().strip()

        if close_date and not close_date.count('/') == 1:
            messagebox.showerror("Error", "Close must be M/D format or empty.")
            return
        if ticker and call_put in ["Call", "Put"]:
            try:
                new_row = [
                    ticker,
                    close_date,
                    call_put,
                    int(contracts),
                    float(premium),
                    float(strike_price)
                ]
                self.data.append(new_row)
                self.populate_treeview()
                self.save_data()
                threading.Thread(target=self.fetch_initial_prices, daemon=True).start()
                window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Check Contracts, Premium & Strike.")
        else:
            messagebox.showerror("Error", "Invalid Ticker.")

    def fetch_initial_prices(self):
        unique_tickers = {row[0] for row in self.data if len(row) == 6}
        for ticker in unique_tickers:
            if ticker not in self.price_cache or self.price_cache[ticker] in ["....", "?"]:
                self.price_cache[ticker] = "...."
                for attempt in range(3):
                    try:
                        yf_ticker = yf.Ticker(ticker)
                        quote = yf_ticker.get_info().get('regularMarketPrice', None)
                        if quote is None or (isinstance(quote, float) and math.isnan(quote)):
                            history = yf_ticker.history(period="1d")
                            quote = round(history["Close"].iloc[-1], 2) if not history.empty and not math.isnan(history["Close"].iloc[-1]) else None
                        self.price_cache[ticker] = quote if quote is not None and not math.isnan(quote) else "?"
                        if self.DevMode == 1:
                            print(f"[{datetime.now().strftime('%H:%M:%S')}] Lookup {ticker}: {self.price_cache[ticker]}")
                        break
                    except Exception as e:
                        if self.DevMode == 1:
                            print(f"[{datetime.now().strftime('%H:%M:%S')}] Lookup {ticker}: Failed, Error: {str(e)}")
                        if attempt < 2:
                            time_module.sleep(1)
                        self.price_cache[ticker] = "?"
                self.root.after(0, self.populate_treeview)

    def populate_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        unique_tickers = {row[0] for row in self.data if len(row) == 6}
        for ticker in unique_tickers:
            if ticker not in self.price_cache:
                self.price_cache[ticker] = "...."

        for index, row in enumerate(self.data):
            if len(row) != 6:
                continue
            ticker, ends, option = row[0], row[1], row[2]
            contracts = row[3]
            premium = row[4]   # treated as total premium paid for the position
            strike_price = row[5]
            current_price = self.price_cache.get(ticker, "....")

            outcome = ""
            diff = float('nan')
            diff_fmt = ""
            value = ""
            value_fmt = ""

            if current_price not in ["....", "?"] and not math.isnan(current_price):
                if option == "Call":
                    diff = current_price - strike_price
                else:  # Put
                    diff = strike_price - current_price

                diff_fmt = f"+{round(diff, 2):.2f}" if diff > 0 else f"-{round(abs(diff), 2):.2f}" if diff < 0 else ""
                outcome = self.calculate_outcome(option, current_price, strike_price)
                intrinsic = max(0.0, diff)

                if outcome:
                    if option == "Put" and outcome == "Purchase":
                        # Calculate based on effective cost vs current price
                        effective_cost = strike_price - (premium / (contracts * 100))
                        value_num = (current_price - effective_cost) * (contracts * 100)
                    else:
                        value_num = (intrinsic * (contracts * 100)) - premium

                    value = round(value_num, 2)

                    if value > 0:
                        value_fmt = f"+{value:,.0f}" if value.is_integer() else f"+{value:,.2f}"
                    else:
                        value_fmt = f"{value:,.0f}" if value.is_integer() else f"{value:,.2f}"
                else:
                    value = ""
                    value_fmt = ""

            diff_tag = 'green_diff' if not math.isnan(diff) and diff > 0 else 'red_diff' if not math.isnan(diff) and diff < 0 else ''
            strike_price_fmt = int(strike_price) if strike_price == int(strike_price) else round(strike_price, 2)
            current_price_fmt = current_price if current_price in ["....", "?"] else (int(current_price) if current_price == int(current_price) else round(current_price, 2))
            tag = 'redrow' if outcome else ('oddrow' if index % 2 else 'evenrow')

            item = self.tree.insert("", "end", tags=(tag, f"list_index_{index}"), values=(
                ticker, ends, option, contracts, int(premium), strike_price_fmt, current_price_fmt, diff_fmt, outcome, value_fmt
            ))

            if diff_tag:
                self.tree.set(item, "Diff", diff_fmt)

        self.last_updated_label.config(text=f"{datetime.now().strftime('%H:%M:%S')}")

    def calculate_outcome(self, call_put, current_price, strike_price):
        # returns Sell for Calls (if ITM), Purchase for Puts (if ITM)
        if call_put == "Put":
            return "Purchase" if strike_price > current_price else ""
        else:  # Call
            return "Sell" if strike_price < current_price else ""

    def refresh_data(self):
        unique_tickers = {row[0] for row in self.data if len(row) == 6}
        for ticker in unique_tickers:
            if ticker not in self.price_cache or self.is_market_open():
                for attempt in range(3):
                    try:
                        yf_ticker = yf.Ticker(ticker)
                        quote = yf_ticker.get_info().get('regularMarketPrice', None)
                        if quote is None or (isinstance(quote, float) and math.isnan(quote)):
                            history = yf_ticker.history(period="1d")
                            quote = round(history["Close"].iloc[-1], 2) if not history.empty and not math.isnan(history["Close"].iloc[-1]) else None
                        self.price_cache[ticker] = quote if quote is not None and not math.isnan(quote) else "?"
                        if self.DevMode == 1:
                            print(f"[{datetime.now().strftime('%H:%M:%S')}] Lookup {ticker}: {self.price_cache[ticker]}")
                        break
                    except Exception as e:
                        if self.DevMode == 1:
                            print(f"[{datetime.now().strftime('%H:%M:%S')}] Lookup {ticker}: Failed, Error: {str(e)}")
                        if attempt < 2:
                            time_module.sleep(1)
                        self.price_cache[ticker] = "?"
                self.root.after(0, self.populate_treeview)

        self.root.after(0, self.populate_treeview)

    def remove_selected(self):
        selected = self.tree.selection()
        if selected:
            indices_to_drop = []
            for item in selected:
                tags = self.tree.item(item, "tags")
                for tag in tags:
                    if tag.startswith("list_index_"):
                        indices_to_drop.append(int(tag.replace("list_index_", "")))
                        break
            if indices_to_drop:
                self.data = [row for i, row in enumerate(self.data) if i not in indices_to_drop]
                self.populate_treeview()
                self.save_data()

    def confirm_remove_all(self):
        if messagebox.askyesno("Confirm", "Remove all options?"):
            self.remove_all()

    def remove_all(self):
        self.data = []
        self.price_cache = {}
        self.populate_treeview()
        self.save_data()

    def sort_column(self, col):
        self.current_sort_col = col
        self.current_sort_reverse = not self.sort_reverse.get(col, False)
        self.sort_reverse[col] = self.current_sort_reverse

        # Data rows: [Ticker, Ends, Option, Contracts, Premium, Strike]
        def get_sort_key(item):
            # item is one of the rows from self.data
            if len(item) != 6:
                return (float('inf'), "")
            ticker = item[0]

            if col == "Ticker":
                return (item[0], ticker)
            if col == "Ends":
                try:
                    month, day = map(int, item[1].split('/'))
                    current_year = datetime.now().year
                    return ((current_year, month, day), ticker)
                except Exception:
                    return ((9999, 12, 31), ticker)
            if col == "Option":
                return (item[2], ticker)
            if col == "Contracts":
                return (item[3], ticker)
            if col == "Premium":
                return (item[4], ticker)
            if col == "Strike":
                return (item[5], ticker)
            if col == "Current":
                current_price = self.price_cache.get(ticker, float('nan'))
                return (current_price, ticker) if current_price not in ["....", "?"] and not math.isnan(current_price) else (float('inf'), ticker)
            if col == "Diff":
                current_price = self.price_cache.get(ticker, float('nan'))
                strike_price = item[5]
                if item[2] == "Call":
                    if current_price not in ["....", "?"] and not math.isnan(current_price):
                        diff = current_price - strike_price
                        return (diff, ticker)
                else:  # Put
                    if current_price not in ["....", "?"] and not math.isnan(current_price):
                        diff = strike_price - current_price
                        return (diff, ticker)
                return (float('inf'), ticker)
            if col == "Value":
                current_price = self.price_cache.get(ticker, float('nan'))
                strike_price = item[5]
                contracts = item[3]
                premium = item[4]
                if current_price not in ["....", "?"] and not math.isnan(current_price):
                    if item[2] == "Call":
                        intrinsic = max(0.0, current_price - strike_price)
                    else:
                        intrinsic = max(0.0, strike_price - current_price)
                    value = (intrinsic * (contracts * 100)) - premium
                    return (value, ticker)
                return (float('inf'), ticker)
            if col == "Outcome":
                return (self.calculate_outcome(item[2], self.price_cache.get(ticker, float('nan')) if self.price_cache.get(ticker, float('nan')) not in ["....", "?"] else float('nan'), item[5]), ticker)

            return (str(item[0]), ticker)

        self.data = sorted(self.data, key=get_sort_key, reverse=self.current_sort_reverse)
        self.populate_treeview()

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            column = self.tree.identify_column(event.x)
            # allow editing Ticker, Ends, Option, Contracts, Premium, Strike (#1..#6)
            if column in ("#1", "#2", "#3", "#4", "#5", "#6"):
                self.editing_item = item
                self.start_editing(item, column)

    def start_editing(self, item, column):
        current_value = self.tree.set(item, column)
        entry = tk.Entry(self.root)
        bbox = self.tree.bbox(item, column)
        if bbox:
            # bbox returns (x, y, width, height) relative to tree; placing on root is usually fine in most setups
            entry.place(x=bbox[0], y=bbox[1], width=bbox[2], anchor="nw")
            entry.insert(0, current_value)
            entry.focus_set()

            def save_edit(event=None):
                new_value_raw = entry.get().strip()
                col_map = {"#1": 0, "#2": 1, "#3": 2, "#4": 3, "#5": 4, "#6": 5}
                col_index = col_map.get(column)
                if col_index is None:
                    entry.destroy()
                    return

                # Validation & casting
                try:
                    if col_index == 0:  # Ticker
                        if len(new_value_raw) > 5:
                            messagebox.showerror("Error", "Ticker cannot exceed 5 characters.")
                            entry.destroy()
                            return
                        new_value = new_value_raw.upper()
                    elif col_index == 1:  # Ends (M/D)
                        if new_value_raw and not new_value_raw.count('/') == 1:
                            messagebox.showerror("Error", "Close must be M/D format or empty.")
                            entry.destroy()
                            return
                        new_value = new_value_raw
                    elif col_index == 2:  # Option
                        new_value = new_value_raw.capitalize()
                        if new_value not in ["Call", "Put"]:
                            messagebox.showerror("Error", "Option must be Call or Put.")
                            entry.destroy()
                            return
                    elif col_index == 3:  # Contracts
                        new_value = int(new_value_raw)
                    elif col_index == 4:  # Premium
                        new_value = float(new_value_raw)
                    elif col_index == 5:  # Strike
                        new_value = float(new_value_raw)
                    else:
                        new_value = new_value_raw
                except ValueError:
                    messagebox.showerror("Error", "Invalid numeric value.")
                    entry.destroy()
                    return

                # find data list index
                tags = self.tree.item(item, "tags")
                list_index = None
                for tag in tags:
                    if tag.startswith("list_index_"):
                        list_index = int(tag.replace("list_index_", ""))
                        break
                if list_index is not None and list_index < len(self.data):
                    self.data[list_index][col_index] = new_value
                    self.populate_treeview()
                    self.save_data()
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("710x300") #Width x Height
    app = OptionsMonitor(root)
    root.mainloop()
