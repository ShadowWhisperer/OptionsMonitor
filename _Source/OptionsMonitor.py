import tkinter as tk
from tkinter import ttk, messagebox
import yfinance as yf
import pandas as pd
from datetime import datetime, time
import os
import pytz

class OptionsMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Options Monitor")
        self.data_file = r"C:\ProgramData\ShadowWhisperer\OptionsMonitor\data.csv"
        self.df = self.load_data()
        self.sort_reverse = {}  # Track sort direction
        self.editing_item = None
        self.setup_gui()

    def load_data(self):
        os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
        try:
            df = pd.read_csv(self.data_file)
            if 'Open Date' in df.columns:
                df = df.drop(columns=['Open Date'])
            if 'Time Remaining' in df.columns:
                df = df.drop(columns=['Time Remaining'])
            if 'Current Price' in df.columns:
                df = df.drop(columns=['Current Price'])
            if 'Outcome' in df.columns:
                df = df.drop(columns=['Outcome'])
            if 'Diff' in df.columns:
                df = df.drop(columns=['Diff'])
            return df
        except FileNotFoundError:
            return pd.DataFrame(columns=["Ticker", "Close Date", "Option", "Strike Price"])

    def save_data(self):
        self.df[["Ticker", "Close Date", "Option", "Strike Price"]].to_csv(self.data_file, index=False)

    def setup_gui(self):
        self.tree = ttk.Treeview(self.root, columns=("Ticker", "Close Date", "Option", "Strike Price", "Current", "Diff", "Outcome"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c), anchor="w")
            self.tree.column(col, width=70, anchor="w")
            self.sort_reverse[col] = False  # Initialize sort direction
        self.tree.tag_configure('oddrow', background='#f0f0f0')
        self.tree.tag_configure('evenrow', background='#ffffff')
        self.tree.tag_configure('redrow', background='#ffcccc')
        self.tree.pack(pady=10, fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Refresh", command=self.refresh_data).pack(side="left", padx=5)
        tk.Button(button_frame, text="Add", command=self.add_option).pack(side="left", padx=5)
        tk.Button(button_frame, text="Remove", command=self.remove_selected).pack(side="left", padx=5)
        tk.Button(button_frame, text="Remove All", command=self.remove_all).pack(side="left", padx=5)

        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(fill="x", side="bottom")
        self.last_updated_label = tk.Label(self.status_frame, anchor="e")
        self.last_updated_label.pack(side="right")
        self.market_status_label = tk.Label(self.status_frame, anchor="w")
        self.market_status_label.pack(side="left")
        
        self.update_market_status()
        self.populate_treeview()

    def is_market_open(self):
        now = datetime.now(pytz.timezone('US/Eastern'))
        market_open = time(9, 30)
        market_close = time(16, 0)
        is_weekday = now.weekday() < 5
        is_open_hours = market_open <= now.time() <= market_close
        return is_weekday and is_open_hours

    def update_market_status(self):
        status = "Markets Open" if self.is_market_open() else "Markets Closed"
        color = "green" if self.is_market_open() else "red"
        self.market_status_label.config(text=status, fg=color)
        self.root.after(60000, self.update_market_status)

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
            ticker_entry, call_put_var, strike_entry, close_date_entry, add_window)
        ).pack(pady=10)

    def _submit_option(self, ticker_entry, call_put_var, strike_entry, close_date_entry, window):
        ticker = ticker_entry.get().upper()
        call_put = call_put_var.get()
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
                self.df = pd.concat([self.df, new_row], ignore_index=True, verify_integrity=True)
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
                    quote = round(history["Close"].iloc[-1], 2) if not history.empty else 0.0
                price_cache[ticker] = quote
            except Exception:
                price_cache[ticker] = 0.0
        for index, row in self.df.iterrows():
            current_price = price_cache.get(row["Ticker"], 0.0)
            outcome = self.calculate_outcome(row["Option"], current_price, row["Strike Price"])
            diff = current_price - row["Strike Price"]
            diff_fmt = f"+{round(diff, 2)}" if diff > 0 else f"{round(diff, 2)}"
            diff_tag = 'green_diff' if diff > 0 else 'red_diff' if diff < 0 else ''
            strike_price_fmt = int(row["Strike Price"]) if row["Strike Price"] == int(row["Strike Price"]) else round(row["Strike Price"], 2)
            current_price_fmt = int(current_price) if current_price == int(current_price) else round(current_price, 2)
            tag = 'redrow' if outcome else ('oddrow' if index % 2 else 'evenrow')
            item = self.tree.insert("", "end", tags=(tag,), values=(
                row["Ticker"], row["Close Date"], row["Option"], strike_price_fmt, current_price_fmt, diff_fmt, outcome
            ))
            if diff_tag:
                self.tree.item(item, tags=(tag, diff_tag))
                self.tree.column("Diff", anchor="w")
                self.tree.set(item, "Diff", diff_fmt)

        self.last_updated_label.config(text=f"Updated  {datetime.now().strftime('%H:%M:%S')} ")

    def calculate_outcome(self, call_put, current_price, strike_price):
        if call_put == "Call":
            return "Sell" if current_price > strike_price else ""
        else:
            return "Purchase" if current_price < strike_price else ""

    def refresh_data(self):
        unique_tickers = self.df["Ticker"].unique()
        price_cache = {}
        for ticker in unique_tickers:
            try:
                yf_ticker = yf.Ticker(ticker)
                quote = yf_ticker.get_info().get('regularMarketPrice', None)
                if quote is None:
                    history = yf_ticker.history(period="1d")
                    quote = round(history["Close"].iloc[-1], 2) if not history.empty else 0.0
                price_cache[ticker] = quote
            except Exception:
                price_cache[ticker] = 0.0
        self.populate_treeview()

    def remove_selected(self):
        selected = self.tree.selection()
        if selected:
            for item in selected:
                index = int(self.tree.index(item))
                self.df = self.df.drop(index)
            self.df.reset_index(drop=True, inplace=True)
            self.populate_treeview()
            self.save_data()

    def remove_all(self):
        self.df = pd.DataFrame(columns=["Ticker", "Close Date", "Option", "Strike Price"])
        self.populate_treeview()
        self.save_data()

    def sort_column(self, col):
        self.sort_reverse[col] = not self.sort_reverse.get(col, False)  # Toggle sort direction
        items = list(self.tree.get_children())
        data = [(self.tree.index(item), item) for item in items]
        if col == "Close Date":
            current_year = datetime.now().year
            data.sort(key=lambda x: datetime.strptime(self.tree.set(x[1], col) + f"/{current_year}", "%m/%d/%Y"), reverse=self.sort_reverse[col])
        elif col == "Diff":
            data.sort(key=lambda x: float(self.tree.set(x[1], col).replace('+', '')) if self.tree.set(x[1], col) else 0.0, reverse=self.sort_reverse[col])
        else:
            try:
                data.sort(key=lambda x: float(self.tree.set(x[1], col)) if col in ["Strike Price", "Current"] else self.tree.set(x[1], col), reverse=self.sort_reverse[col])
            except ValueError:
                data.sort(key=lambda x: self.tree.set(x[1], col), reverse=self.sort_reverse[col])
        for new_index, (orig_index, item) in enumerate(data):
            current_price = float(self.tree.set(item, "Current")) if self.tree.set(item, "Current") else 0.0
            strike_price = float(self.tree.set(item, "Strike Price")) if self.tree.set(item, "Strike Price") else 0.0
            diff = current_price - strike_price
            diff_fmt = f"+{round(diff, 2)}" if diff > 0 else f"{round(diff, 2)}"
            diff_tag = 'green_diff' if diff > 0 else 'red_diff' if diff < 0 else ''
            option = self.tree.set(item, "Option")
            outcome = self.calculate_outcome(option, current_price, strike_price)
            strike_price_fmt = int(strike_price) if strike_price == int(strike_price) else round(strike_price, 2)
            current_price_fmt = int(current_price) if current_price == int(current_price) else round(current_price, 2)
            tag = 'redrow' if outcome else ('oddrow' if new_index % 2 else 'evenrow')
            self.tree.item(item, tags=(tag, diff_tag))
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
            if column in ("#1", "#2", "#3", "#4"):  # Editable columns: Ticker, Close Date, Option, Strike Price
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
                if column == "#3" and new_value not in ["Call", "Put"]:
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
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)

    def update_df_from_tree(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            data.append({
                "Ticker": values[0],
                "Close Date": values[1],
                "Option": values[2],
                "Strike Price": float(values[3]) if values[3] else 0.0
            })
        self.df = pd.DataFrame(data)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("800x420")
    app = OptionsMonitor(root)
    root.mainloop()
