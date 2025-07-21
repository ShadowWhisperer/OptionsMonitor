import tkinter as tk
from tkinter import ttk, messagebox
import yfinance as yf
import pandas as pd
from datetime import datetime, time, timedelta
import os
import pytz

class OptionsMonitor:
    def __init__(self, root): 
        self.root = root
        self.root.title("Options Monitor - Ver 1.4")
        self.data_file = r"C:\ProgramData\ShadowWhisperer\OptionsMonitor\data.csv"
        self.df = self.load_data()
        self.sort_reverse = {}
        self.last_market_status = None
        self.refresh_interval = None
        self.last_interval = "5 Mins"
        self.last_updated_label = None
        self.DevMode = 0
        self.setup_gui()

    def load_data(self):
        os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
        try:
            df = pd.read_csv(self.data_file)
            for col in ['Open Date', 'Time Remaining', 'Current Price', 'Outcome', 'Diff']:
                if col in df.columns:
                    df = df.drop(columns=[col])
            return df
        except FileNotFoundError:
            return pd.DataFrame(columns=["Ticker", "Close Date", "Option", "Strike Price"])

    def save_data(self):
        self.df[["Ticker", "Close Date", "Option", "Strike Price"]].to_csv(self.data_file, index=False)

    def setup_gui(self):
        self.root.grid_rowconfigure(0, weight=1)    # Treeview row is expandable
        self.root.grid_rowconfigure(1, weight=0)    # Button frame row is fixed
        self.root.grid_rowconfigure(2, weight=0)    # Status frame row is fixed
        self.root.grid_columnconfigure(0, weight=1) # Main column is expandable

        self.tree = ttk.Treeview(self.root, columns=("Ticker", "Close Date", "Option", "Strike Price", "Current", "Diff", "Outcome"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c), anchor="w")
            self.tree.column(col, width=70, anchor="w")
            self.sort_reverse[col] = False
        self.tree.tag_configure('oddrow', background='#f0f0f0')
        self.tree.tag_configure('evenrow', background='#ffffff')
        self.tree.tag_configure('redrow', background='#ffcccc')
        self.tree.grid(row=0, column=0, pady=2, sticky="nsew")
        self.tree.bind("<Double-1>", self.on_double_click)

        button_frame = tk.Frame(self.root, height=30)
        button_frame.grid(row=1, column=0, pady=2, sticky="ew")  # Don't squeeze buttons, when adjusting window size
        self.refresh_button = tk.Button(button_frame, text="Refresh", command=self.refresh_data, width=10)
        self.refresh_button.pack(side="left", padx=(15, 5))
        tk.Button(button_frame, text="Add", command=self.add_option, width=10).pack(side="left", padx=(10, 5))
        tk.Button(button_frame, text="Remove", command=self.remove_selected, width=10).pack(side="left", padx=(10, 5))
        tk.Button(button_frame, text="Remove All", command=self.remove_all, width=10).pack(side="left", padx=(10, 5))
        
        self.last_updated_label = tk.Label(button_frame, text="")
        self.last_updated_label.pack(side="right", padx=5)

        self.interval_var = tk.StringVar(value="5 Mins")
        self.interval_combo = ttk.Combobox(button_frame, textvariable=self.interval_var, values=["Don't Update", "5 Mins", "10 Mins", "15 Mins", "30 Mins"], width=16, state="readonly")
        self.interval_combo.pack(side="right", padx=(0, 0))
        self.interval_combo.bind("<<ComboboxSelected>>", self.schedule_refresh)

        style = ttk.Style()
        style.configure("Red.TCombobox", foreground="red")
        style.map("Red.TCombobox", foreground=[("disabled", "red")])
        
        self.status_frame = tk.Frame(self.root)
        self.status_frame.grid(row=2, column=0, sticky="ew")
        
        self.update_market_status()
        self.populate_treeview()
        self.schedule_refresh()

    def is_market_open(self):
        if self.DevMode == 1: # Dev Mode -- Pretend the markets are open
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
                self.interval_combo.config(values=["Don't Update", "5 Mins", "10 Mins", "15 Mins", "30 Mins"], state="readonly", style="TCombobox")
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

        self.last_interval = interval  # Keep the full string with " Mins"
        
        try:
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
        add_window.geometry("300x200")

        form_frame = tk.Frame(add_window)
        form_frame.pack(pady=10, padx=10)

        tk.Label(form_frame, text="Ticker").grid(row=0, column=0, sticky="e", padx=5, pady=4)
        ticker_entry = tk.Entry(form_frame)
        ticker_entry.grid(row=0, column=1, pady=5)

        tk.Label(form_frame, text="Option").grid(row=1, column=0, sticky="e", padx=5, pady=4)
        call_put_var = tk.StringVar(value="Call")
        option_frame = tk.Frame(form_frame)
        tk.Radiobutton(option_frame, text="Call", variable=call_put_var, value="Call").pack(side="left")
        tk.Radiobutton(option_frame, text="Put", variable=call_put_var, value="Put").pack(side="left")
        option_frame.grid(row=1, column=1, pady=5)

        tk.Label(form_frame, text="Strike Price").grid(row=2, column=0, sticky="e", padx=5, pady=4)
        strike_entry = tk.Entry(form_frame)
        strike_entry.grid(row=2, column=1, pady=5)

        tk.Label(form_frame, text="Close Date (M/D)").grid(row=3, column=0, sticky="e", padx=5, pady=4)
        close_date_entry = tk.Entry(form_frame)
        close_date_entry.grid(row=3, column=1, pady=5)

        tk.Button(add_window, text="Add Option", command=lambda: self._submit_option(
            ticker_entry, call_put_var, strike_entry, close_date_entry, add_window), width=10
        ).pack(pady=10)

    def _submit_option(self, ticker_entry, call_put_var, strike_entry, close_date_entry, window):
        ticker = ticker_entry.get().upper()
        call_put = call_put_var.get().capitalize()
        strike_price = strike_entry.get()
        close_date = close_date_entry.get()

        try:
            float(strike_price)
            datetime.strptime(close_date, "%m/%d")
            if ticker and call_put in ["Call", "Put"]:
                new_row = pd.DataFrame({
                    "Ticker": [ticker],
                    "Close Date": [close_date],
                    "Option": [call_put],
                    "Strike Price": [float(strike_price)],
                })
                self.df = pd.concat([self.df, new_row], ignore_index=True)
                self.populate_treeview()
                self.save_data()
                window.destroy()
            else:
                messagebox.showerror("Error", "Invalid Ticker.")
        except (ValueError, Exception):
            messagebox.showerror("Error", "Invalid input. Check Strike Price and Close Date (M/D).")

    def populate_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        unique_tickers = self.df["Ticker"].unique()
        price_cache = {}
        for ticker in unique_tickers:
            try:
                yf_ticker = yf.Ticker(ticker)
                quote = yf_ticker.get_info().get('regularMarketPrice', None)
                if quote is None:
                    history = yf_ticker.history(period="1d")
                    quote = round(history["Close"].iloc[-1], 2) if not history.empty else None
                price_cache[ticker] = quote if quote is not None else float('nan')
            except Exception:
                price_cache[ticker] = float('nan')
        for index, row in self.df.iterrows():
            current_price = price_cache.get(row["Ticker"], float('nan'))
            outcome = self.calculate_outcome(row["Option"], current_price, row["Strike Price"]) if not pd.isna(current_price) else ""
            diff = current_price - row["Strike Price"] if not pd.isna(current_price) else float('nan')
            diff_fmt = f"+{round(diff, 2):.2f}" if not pd.isna(diff) and diff > 0 else f"-{round(abs(diff), 2):.2f}" if not pd.isna(diff) and diff < 0 else ""
            diff_tag = 'green_diff' if not pd.isna(diff) and diff > 0 else 'red_diff' if not pd.isna(diff) and diff < 0 else ''
            strike_price_fmt = int(row["Strike Price"]) if row["Strike Price"] == int(row["Strike Price"]) else round(row["Strike Price"], 2)
            current_price_fmt = int(current_price) if not pd.isna(current_price) and current_price == int(current_price) else round(current_price, 2) if not pd.isna(current_price) else ""
            tag = 'redrow' if outcome else ('oddrow' if index % 2 else 'evenrow')
            item = self.tree.insert("", "end", tags=(tag, f"df_index_{index}"), values=(
                row["Ticker"], row["Close Date"], row["Option"], strike_price_fmt, current_price_fmt, diff_fmt, outcome
            ))
            if diff_tag:
                self.tree.item(item, tags=(tag, diff_tag, f"df_index_{index}"))
                self.tree.set(item, "Diff", diff_fmt)

        self.last_updated_label.config(text=f"{datetime.now().strftime('%H:%M:%S')}")

    def calculate_outcome(self, call_put, current_price, strike_price):
        if call_put == "Call":
            return "Sell" if current_price > strike_price else ""
        else:
            return "Purchase" if current_price < strike_price else ""

    def refresh_data(self):
        self.populate_treeview()

    def remove_selected(self):
        selected = self.tree.selection()
        if selected:
            indices_to_drop = []
            for item in selected:
                tags = self.tree.item(item, "tags")
                for tag in tags:
                    if tag.startswith("df_index_"):
                        df_index = int(tag.replace("df_index_", ""))
                        indices_to_drop.append(df_index)
                        break
            if indices_to_drop:
                self.df = self.df.drop(indices_to_drop)
                self.df.reset_index(drop=True, inplace=True)
                self.populate_treeview()
                self.save_data()

    def remove_all(self):
        self.df = pd.DataFrame(columns=["Ticker", "Close Date", "Option", "Strike Price"])
        self.populate_treeview()
        self.save_data()

    def sort_column(self, col):
        self.sort_reverse[col] = not self.sort_reverse.get(col, False)
        items = list(self.tree.get_children())
        data = [(self.tree.index(item), item) for item in items]
        
        def get_sort_key(x):
            value = self.tree.set(x[1], col)
            ticker = self.tree.set(x[1], "Ticker") # Secondary sort key for tie-breaking
            if col == "Close Date":
                current_year = datetime.now().year
                try:
                    return (datetime.strptime(value + f"/{current_year}", "%m/%d/%Y"), ticker)
                except ValueError:
                    return (datetime.max, ticker)  # Handle invalid dates
            elif col == "Diff":
                try:
                    num = float(value.replace('+', '').replace('-', '')) * (-1 if '-' in value else 1)
                    return (num, ticker)
                except ValueError:
                    return (float('inf'), ticker)  # Handle empty or invalid Diff
            elif col == "Strike Price" or col == "Current":
                try:
                    return (float(value), ticker)  # Force numeric sort for Strike Price and Current
                except ValueError:
                    return (float('inf'), ticker)  # Handle non-numeric values
            else:
                return (value, ticker)             # Default string sort for other columns
        
        data.sort(key=get_sort_key, reverse=self.sort_reverse[col])
        
        for new_index, (orig_index, item) in enumerate(data):
            current_price = float(self.tree.set(item, "Current")) if self.tree.set(item, "Current") else float('nan')
            strike_price = float(self.tree.set(item, "Strike Price")) if self.tree.set(item, "Strike Price") else 0.0
            diff = current_price - strike_price if not pd.isna(current_price) else float('nan')
            diff_fmt = f"+{round(diff, 2):.2f}" if not pd.isna(diff) and diff > 0 else f"-{round(abs(diff), 2):.2f}" if not pd.isna(diff) and diff < 0 else ""
            diff_tag = 'green_diff' if not pd.isna(diff) and diff > 0 else 'red_diff' if not pd.isna(diff) and diff < 0 else ''
            option = self.tree.set(item, "Option")
            outcome = self.calculate_outcome(option, current_price, strike_price) if not pd.isna(current_price) else ""
            strike_price_fmt = int(strike_price) if strike_price == int(strike_price) else round(strike_price, 2)
            current_price_fmt = int(current_price) if not pd.isna(current_price) and current_price == int(current_price) else round(current_price, 2) if not pd.isna(current_price) else ""
            tags = [tag for tag in self.tree.item(item, "tags") if tag.startswith("df_index_")]
            tag = 'redrow' if outcome else ('oddrow' if new_index % 2 else 'evenrow')
            tags.append(tag)
            if diff_tag:
                tags.append(diff_tag)
            self.tree.item(item, tags=tuple(tags))
            self.tree.item(item, values=(
                self.tree.set(item, "Ticker"), 
                self.tree.set(item, "Close Date"), 
                option, 
                strike_price_fmt, 
                current_price_fmt, 
                diff_fmt,
                outcome
            ))
            self.tree.move(item, "", new_index)

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            column = self.tree.identify_column(event.x)
            if column in ("#1", "#2", "#3", "#4"):
                self.editing_item = item
                self.start_editing(item, column)

    def start_editing(self, item, column):
        current_value = self.tree.set(item, column)
        entry = tk.Entry(self.root)
        bbox = self.tree.bbox(item, column)
        if bbox:
            entry.place(x=bbox[0], y=bbox[1], width=bbox[2], anchor="nw")
            entry.insert(0, current_value)
            entry.focus_set()

            def save_edit(event=None):
                new_value = entry.get()
                if column == "#1" and len(new_value) > 5:
                    messagebox.showerror("Error", "Ticker cannot exceed 5 characters.")
                    entry.destroy()
                    return
                if column == "#3":
                    new_value = new_value.capitalize()
                    if new_value not in ["Call", "Put"]:
                        messagebox.showerror("Error", "Option must be Call or Put.")
                        entry.destroy()
                        return
                if column == "#4":
                    try:
                        float(new_value)
                    except ValueError:
                        messagebox.showerror("Error", "Strike Price must be a number.")
                        entry.destroy()
                        return
                if column == "#2":
                    try:
                        datetime.strptime(new_value, "%m/%d")
                    except ValueError:
                        messagebox.showerror("Error", "Close Date must be in M/D format.")
                        entry.destroy()
                        return
                self.tree.set(item, column, new_value)
                self.update_df_from_tree()
                self.save_data()
                if column in ["#3", "#4"]:
                    current_price = float(self.tree.set(item, "Current")) if self.tree.set(item, "Current") else float('nan')
                    strike_price = float(self.tree.set(item, "Strike Price")) if self.tree.set(item, "Strike Price") else 0.0
                    option = self.tree.set(item, "Option")
                    outcome = self.calculate_outcome(option, current_price, strike_price) if not pd.isna(current_price) else ""
                    self.tree.set(item, "Outcome", outcome)
                    tags = [tag for tag in self.tree.item(item, "tags") if tag.startswith("df_index_")]
                    tag = 'redrow' if outcome else ('oddrow' if int(self.tree.index(item)) % 2 else 'evenrow')
                    tags.append(tag)
                    self.tree.item(item, tags=tuple(tags))
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)

    def update_df_from_tree(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            tags = self.tree.item(item, "tags")
            df_index = None
            for tag in tags:
                if tag.startswith("df_index_"):
                    df_index = int(tag.replace("df_index_", ""))
                    break
            data.append({
                "index": df_index,
                "Ticker": values[0],
                "Close Date": values[1],
                "Option": values[2],
                "Strike Price": float(values[3]) if values[3] else 0.0
            })
        new_df = pd.DataFrame(data)
        new_df.set_index("index", inplace=True, drop=True)
        new_df = new_df[new_df.index.notnull()]
        new_df.index.name = None
        self.df = new_df[["Ticker", "Close Date", "Option", "Strike Price"]]
        self.df.reset_index(drop=True, inplace=True)
        self.populate_treeview()

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("620x300")
    app = OptionsMonitor(root)
    root.mainloop()