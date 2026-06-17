"""
IP Info Checker - Main Application
Supports both ipinfo.io API and local database lookups.
"""

import os
import re
import json
from datetime import datetime
import requests
import xlsxwriter
import csv
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import logging
from threading import Thread
import time
import ipaddress

from local_ip_db import (
    LocalIPDatabase, ConfigManager, get_config, get_local_db,
    fetch_ip_data_local
)

# Headers for output files
EXPORT_HEADERS = ["IP", "ISP", "Hostname", "Region", "Country_code"]

# Export format options
EXPORT_FORMATS = {
    'xlsx': 'Excel (.xlsx)',
    'csv': 'CSV (.csv)',
    'json': 'JSON (.json)'
}

# Logging setup
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Load environment variables
load_dotenv()

# Get API token (optional if using local DB)
ipinfo_token = os.getenv('IPINFO_TOKEN')

# Global variables
result_list = []


class APILimitReached(Exception):
    """Custom exception for API rate limit errors"""
    pass


def is_valid_ipv4(ip):
    """Checks if a given string is a valid IPv4 address."""
    ip_regex = re.compile(r"^(?:\d{1,3}\.){3}\d{1,3}$")
    return ip_regex.match(ip) is not None


def is_valid_ipv6(ip):
    """Checks if a given string is a valid IPv6 address."""
    try:
        return isinstance(ipaddress.ip_address(ip), ipaddress.IPv6Address)
    except ValueError:
        return False


def fetch_ip_data_api(ip, max_retries=5):
    """Fetches data for a single IP address from the ipinfo.io API."""
    if not ipinfo_token:
        raise ValueError("API token not configured. Use local database or set IPINFO_TOKEN.")

    url = f"https://ipinfo.io/{ip}?token={ipinfo_token}"
    retry_count = 0
    delay = 1

    while retry_count <= max_retries:
        try:
            response = requests.get(url)
            if response.status_code == 429:
                if retry_count == max_retries:
                    raise APILimitReached("API rate limit reached after max retries")
                logging.warning(f"Rate limit hit for IP {ip}. Retrying in {delay} second(s)...")
                time.sleep(delay)
                retry_count += 1
                delay *= 2
                continue
            response.raise_for_status()
            return response.json()
        except APILimitReached:
            raise
        except requests.RequestException as e:
            logging.error(f"Error checking IP {ip}: {e}")
            return None

    raise APILimitReached("API rate limit reached after max retries")


def fetch_ip_data(ip):
    """
    Fetch IP data using either local database or API based on configuration.
    """
    config = get_config()

    if config.get_use_local_db():
        # Check if database path is configured
        db_path = config.get_ip_database_path()
        if not db_path:
            raise ValueError("Local database mode enabled but no database configured. "
                             "Please go to Settings and set up the database.")

        result = fetch_ip_data_local(ip)
        if result:
            return result
        logging.warning(f"IP {ip} not found in local database")
        return None
    else:
        return fetch_ip_data_api(ip)


def update_result_list(data, count=None):
    """Parses the API response and updates the result list."""
    ip = data.get("ip", "N/A")
    isp = data.get("org", "N/A")
    hostname = data.get("hostname", "N/A")
    region = data.get("region", "N/A")
    country_code = data.get("country", "N/A")

    row = [ip, isp, hostname, region, country_code]
    if count is not None:
        row.append(count)
    result_list.append(row)


def write_to_excel(save_directory, partial=False, has_count=False, export_format='xlsx'):
    """
    Writes the collected data to a file (Excel, CSV, or JSON).

    Args:
        save_directory: Directory to save the file
        partial: If True, indicates this is a partial save
        has_count: If True, includes Count as the last column
        export_format: 'xlsx', 'csv', or 'json'

    Returns:
        Path to the saved file or None if no data
    """
    if not result_list:
        logging.warning("No results to save")
        return None

    now = datetime.now().strftime("%d%m%Y-%H%M%S")
    partial_suffix = "_PARTIAL" if partial else ""

    # Determine file extension
    ext = export_format if export_format in ['xlsx', 'csv', 'json'] else 'xlsx'
    filename = f'ipinfo_export-{now}{partial_suffix}.{ext}'
    save_path = os.path.join(save_directory, filename)

    # Prepare headers
    headers = EXPORT_HEADERS.copy()
    if has_count:
        headers.append("Count")

    if export_format == 'csv':
        save_path = _write_to_csv(save_path, headers, has_count)
    elif export_format == 'json':
        save_path = _write_to_json(save_path, headers, has_count)
    else:  # xlsx
        save_path = _write_to_xlsx(save_path, headers, has_count)

    logging.info(f"File saved at: {save_path}")
    return save_path


def _write_to_xlsx(save_path, headers, has_count):
    """Write data to Excel file."""
    import xlsxwriter

    with xlsxwriter.Workbook(save_path) as workbook:
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        # Write headers
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header, bold)

        # Write data rows
        for row_num, row_data in enumerate(result_list, start=1):
            for col_num, cell_data in enumerate(row_data):
                if has_count and col_num == len(row_data) - 1:
                    try:
                        count_value = int(cell_data) if cell_data else 0
                        worksheet.write_number(row_num, col_num, count_value)
                    except (ValueError, TypeError):
                        worksheet.write(row_num, col_num, str(cell_data) if cell_data else "")
                else:
                    worksheet.write(row_num, col_num, str(cell_data))

    return save_path


def _write_to_csv(save_path, headers, has_count):
    """Write data to CSV file."""
    import csv

    with open(save_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(headers)

        for row_data in result_list:
            # Convert all values to strings, handle None
            row = [str(cell) if cell is not None else '' for cell in row_data]
            writer.writerow(row)

    return save_path


def _write_to_json(save_path, headers, has_count):
    """Write data to JSON file."""
    import json

    # Convert to list of dictionaries
    data = []
    for row_data in result_list:
        record = {}
        for i, header in enumerate(headers):
            if i < len(row_data):
                value = row_data[i]
                # Convert Count to integer if present
                if has_count and header == "Count":
                    try:
                        value = int(value) if value else 0
                    except (ValueError, TypeError):
                        value = 0
                record[header] = value if value is not None else ''
            else:
                record[header] = ''
        data.append(record)

    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    return save_path


def get_export_format_choice(parent):
    """
    Show a dialog to choose export format.
    Returns: 'xlsx', 'csv', 'json', or None if cancelled
    """
    dialog = tk.Toplevel(parent)
    dialog.title("Export Format")
    dialog.geometry("300x230")
    dialog.transient(parent)
    dialog.grab_set()

    result = {'format': None}

    ttk.Label(dialog, text="Choose export format:",
              font=("Helvetica", 12)).pack(pady=15)

    format_var = tk.StringVar(value='xlsx')

    formats_frame = ttk.Frame(dialog)
    formats_frame.pack(pady=10)

    ttk.Radiobutton(formats_frame, text="Excel (.xlsx) - Best for < 100K rows",
                    variable=format_var, value='xlsx').pack(anchor=tk.W, pady=3)
    ttk.Radiobutton(formats_frame, text="CSV (.csv) - Best for large datasets",
                    variable=format_var, value='csv').pack(anchor=tk.W, pady=3)
    ttk.Radiobutton(formats_frame, text="JSON (.json) - Best for APIs/programming",
                    variable=format_var, value='json').pack(anchor=tk.W, pady=3)

    def on_ok():
        result['format'] = format_var.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=15)

    ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

    # Center the dialog
    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - dialog.winfo_width()) // 2
    y = parent.winfo_y() + (parent.winfo_height() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")

    parent.wait_window(dialog)
    return result['format']


def extract_ips_from_file(file):
    """Extracts IP addresses and optional Count column from a CSV file."""
    try:
        with open(file, 'r') as f:
            reader = csv.reader(f)
            header = next(reader, None)

            count_index = None
            if header:
                for i, col in enumerate(header):
                    if col.strip().lower() == 'count':
                        count_index = i
                        break

            results = []
            for row in reader:
                if row:
                    ip = row[0].strip()
                    count = row[count_index].strip() if count_index is not None and len(row) > count_index else None
                    results.append({'ip': ip, 'count': count})

            return results
    except Exception as e:
        logging.error(f"Error reading file {file}: {e}")
        return []


class DatabaseSetupWindow:
    """Window for setting up or updating the local IP database."""

    def __init__(self, parent, on_complete=None):
        self.parent = parent
        self.on_complete = on_complete
        self.window = tk.Toplevel(parent)
        self.window.title("IP Database Setup")
        self.window.geometry("600x500")
        self.window.transient(parent)
        self.window.grab_set()

        self.config = get_config()
        self.db = None
        self.is_importing = False

        self._create_widgets()
        self._load_current_config()

    def _create_widgets(self):
        """Create the setup window widgets."""
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="Local IP Database Setup",
                                font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # Current database info frame
        info_frame = ttk.LabelFrame(main_frame, text="Current Database", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 15))

        self.db_status_label = ttk.Label(info_frame, text="No database configured")
        self.db_status_label.pack(anchor=tk.W)

        self.db_stats_label = ttk.Label(info_frame, text="")
        self.db_stats_label.pack(anchor=tk.W)

        # Database file selection
        file_frame = ttk.LabelFrame(main_frame, text="Database Location", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 15))

        db_path_frame = ttk.Frame(file_frame)
        db_path_frame.pack(fill=tk.X)

        ttk.Label(db_path_frame, text="Database file:").pack(side=tk.LEFT)

        self.db_path_var = tk.StringVar()
        self.db_path_entry = ttk.Entry(db_path_frame, textvariable=self.db_path_var, width=40)
        self.db_path_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        ttk.Button(db_path_frame, text="Browse...",
                   command=self._browse_db_path).pack(side=tk.LEFT)

        # Import section
        import_frame = ttk.LabelFrame(main_frame, text="Import/Update Database", padding="10")
        import_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(import_frame,
                  text="Select a CSV/TSV file with IP network ranges to import:").pack(anchor=tk.W)

        csv_frame = ttk.Frame(import_frame)
        csv_frame.pack(fill=tk.X, pady=5)

        self.csv_path_var = tk.StringVar()
        self.csv_path_entry = ttk.Entry(csv_frame, textvariable=self.csv_path_var, width=40)
        self.csv_path_entry.pack(side=tk.LEFT, padx=(0, 5), expand=True, fill=tk.X)

        ttk.Button(csv_frame, text="Select CSV...",
                   command=self._browse_csv).pack(side=tk.LEFT)

        # Delimiter selection
        delim_frame = ttk.Frame(import_frame)
        delim_frame.pack(fill=tk.X, pady=5)

        ttk.Label(delim_frame, text="Delimiter:").pack(side=tk.LEFT)
        self.delimiter_var = tk.StringVar(value="tab")
        ttk.Radiobutton(delim_frame, text="Tab", variable=self.delimiter_var,
                        value="tab").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(delim_frame, text="Comma", variable=self.delimiter_var,
                        value="comma").pack(side=tk.LEFT)

        # Import button
        self.import_btn = ttk.Button(import_frame, text="Import Database",
                                     command=self._start_import)
        self.import_btn.pack(pady=10)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(import_frame, variable=self.progress_var,
                                            maximum=100, length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.progress_label = ttk.Label(import_frame, text="")
        self.progress_label.pack()

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)

        ttk.Button(button_frame, text="Save & Close",
                   command=self._save_and_close).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel",
                   command=self.window.destroy).pack(side=tk.RIGHT)

    def _load_current_config(self):
        """Load and display current configuration."""
        db_path = self.config.get_ip_database_path()

        if db_path:
            self.db_path_var.set(db_path)
            self._update_db_stats(db_path)
        else:
            # Default path
            default_path = os.path.join(os.path.expanduser("~"), "ip_database.db")
            self.db_path_var.set(default_path)
            self.db_status_label.config(text="No database configured")

    def _update_db_stats(self, db_path):
        """Update the database statistics display."""
        if os.path.exists(db_path):
            try:
                db = LocalIPDatabase(db_path)
                if db.is_valid():
                    stats = db.get_stats()
                    self.db_status_label.config(
                        text=f"Database: {os.path.basename(db_path)}")

                    # Show IPv4/IPv6 breakdown
                    v4 = stats.get('ipv4_records', stats['total_records'])
                    v6 = stats.get('ipv6_records', 0)

                    self.db_stats_label.config(
                        text=f"Records: {stats['total_records']:,} (IPv4: {v4:,}, IPv6: {v6:,}) | "
                             f"Countries: {stats['unique_countries']} | "
                             f"Size: {stats['database_size_mb']} MB | "
                             f"Imported: {stats['import_date'][:10]}")
                else:
                    self.db_status_label.config(text="Database exists but is empty")
                    self.db_stats_label.config(text="")
                db.close()
            except Exception as e:
                self.db_status_label.config(text=f"Error reading database: {e}")
                self.db_stats_label.config(text="")
        else:
            self.db_status_label.config(text="Database file does not exist (will be created)")
            self.db_stats_label.config(text="")

    def _browse_db_path(self):
        """Browse for database file location."""
        path = filedialog.asksaveasfilename(
            title="Select Database Location",
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")],
            initialfile="ip_database.db"
        )
        if path:
            self.db_path_var.set(path)
            self._update_db_stats(path)

    def _browse_csv(self):
        """Browse for CSV file to import."""
        path = filedialog.askopenfilename(
            title="Select IP Database CSV",
            filetypes=[("CSV Files", "*.csv"), ("TSV Files", "*.tsv"), ("All Files", "*.*")]
        )
        if path:
            self.csv_path_var.set(path)

    def _start_import(self):
        """Start the database import process."""
        csv_path = self.csv_path_var.get()
        db_path = self.db_path_var.get()

        if not csv_path:
            messagebox.showwarning("Warning", "Please select a CSV file to import.")
            return

        if not db_path:
            messagebox.showwarning("Warning", "Please specify a database location.")
            return

        if not os.path.exists(csv_path):
            messagebox.showerror("Error", f"CSV file not found: {csv_path}")
            return

        # Confirm if database exists
        if os.path.exists(db_path):
            if not messagebox.askyesno("Confirm",
                                       "Database already exists. This will replace all existing data.\n\nContinue?"):
                return

        self.is_importing = True
        self.import_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="Starting import...")

        # Run import in background thread
        delimiter = '\t' if self.delimiter_var.get() == "tab" else ','
        thread = Thread(target=self._run_import, args=(csv_path, db_path, delimiter))
        thread.start()

    def _run_import(self, csv_path, db_path, delimiter):
        """Run the import process in a background thread."""
        try:
            self.db = LocalIPDatabase(db_path)

            def progress_callback(current, total):
                percent = (current / total) * 100 if total > 0 else 0
                self.window.after(0, lambda: self._update_progress(current, total, percent))

            total = self.db.import_from_csv(csv_path, delimiter=delimiter,
                                            progress_callback=progress_callback)

            self.window.after(0, lambda: self._import_complete(True, total))

        except Exception as e:
            logging.error(f"Import error: {e}")
            self.window.after(0, lambda: self._import_complete(False, str(e)))

    def _update_progress(self, current, total, percent):
        """Update progress bar from main thread."""
        self.progress_var.set(percent)
        self.progress_label.config(text=f"Importing: {current:,} / {total:,} records ({percent:.1f}%)")

    def _import_complete(self, success, result):
        """Handle import completion."""
        self.is_importing = False
        self.import_btn.config(state=tk.NORMAL)

        if success:
            self.progress_var.set(100)
            self.progress_label.config(text=f"Import complete! {result:,} records imported.")
            self._update_db_stats(self.db_path_var.get())

            # Auto-save the database path and enable local mode
            db_path = self.db_path_var.get()
            self.config.set_ip_database_path(db_path)
            self.config.set_use_local_db(True)

            messagebox.showinfo("Success",
                                f"Database imported successfully!\n\n"
                                f"{result:,} records added.\n\n"
                                f"Local database mode has been enabled.")
        else:
            self.progress_label.config(text=f"Import failed: {result}")
            messagebox.showerror("Import Failed", f"Error during import:\n\n{result}")

    def _save_and_close(self):
        """Save configuration and close window."""
        if self.is_importing:
            messagebox.showwarning("Warning", "Please wait for import to complete.")
            return

        db_path = self.db_path_var.get()

        if db_path and os.path.exists(db_path):
            self.config.set_ip_database_path(db_path)
            logging.info(f"Database path saved: {db_path}")

        if self.db:
            self.db.close()

        if self.on_complete:
            self.on_complete()

        self.window.destroy()


class SettingsWindow:
    """Window for application settings."""

    def __init__(self, parent, on_settings_changed=None):
        self.parent = parent
        self.on_settings_changed = on_settings_changed
        self.window = tk.Toplevel(parent)
        self.window.title("Settings")
        self.window.geometry("500x600")
        self.window.transient(parent)
        self.window.grab_set()

        self.config = get_config()

        self._create_widgets()
        self._load_settings()

    def _create_widgets(self):
        """Create settings window widgets."""
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main_frame, text="Settings",
                  font=("Helvetica", 16, "bold")).pack(pady=(0, 20))

        # Data source selection
        source_frame = ttk.LabelFrame(main_frame, text="IP Data Source", padding="10")
        source_frame.pack(fill=tk.X, pady=(0, 15))

        self.source_var = tk.StringVar(value="api")

        ttk.Radiobutton(source_frame, text="Use ipinfo.io API (requires token)",
                        variable=self.source_var, value="api",
                        command=self._on_source_change).pack(anchor=tk.W, pady=2)

        ttk.Radiobutton(source_frame, text="Use Local Database (offline)",
                        variable=self.source_var, value="local",
                        command=self._on_source_change).pack(anchor=tk.W, pady=2)

        # API settings
        self.api_frame = ttk.LabelFrame(main_frame, text="API Settings", padding="10")
        self.api_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(self.api_frame, text="API Token status:").pack(anchor=tk.W)
        self.api_status_label = ttk.Label(self.api_frame, text="")
        self.api_status_label.pack(anchor=tk.W, pady=5)

        ttk.Label(self.api_frame,
                  text="Set IPINFO_TOKEN in .env file to use API",
                  foreground="gray").pack(anchor=tk.W)

        # Local DB settings
        self.local_frame = ttk.LabelFrame(main_frame, text="Local Database Settings", padding="10")
        self.local_frame.pack(fill=tk.X, pady=(0, 15))

        # Database path selection
        db_path_frame = ttk.Frame(self.local_frame)
        db_path_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(db_path_frame, text="Database file:").pack(side=tk.LEFT)

        self.db_path_var = tk.StringVar()
        self.db_path_entry = ttk.Entry(db_path_frame, textvariable=self.db_path_var, width=30)
        self.db_path_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        ttk.Button(db_path_frame, text="Browse...",
                   command=self._browse_database).pack(side=tk.LEFT)

        self.local_status_label = ttk.Label(self.local_frame, text="")
        self.local_status_label.pack(anchor=tk.W, pady=5)

        # Buttons for database management
        db_btn_frame = ttk.Frame(self.local_frame)
        db_btn_frame.pack(anchor=tk.W, pady=5)

        ttk.Button(db_btn_frame, text="Create/Update Database...",
                   command=self._open_db_setup).pack(side=tk.LEFT, padx=(0, 5))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)

        ttk.Button(button_frame, text="Save",
                   command=self._save_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel",
                   command=self.window.destroy).pack(side=tk.RIGHT)

    def _browse_database(self):
        """Browse for an existing database file."""
        file_path = filedialog.askopenfilename(
            title="Select IP Database File",
            filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")],
            initialdir=os.path.dirname(self.db_path_var.get()) if self.db_path_var.get() else None
        )
        if file_path:
            self.db_path_var.set(file_path)
            self._update_local_db_status()

    def _load_settings(self):
        """Load current settings."""
        # Load data source
        use_local = self.config.get_use_local_db()
        self.source_var.set("local" if use_local else "api")

        # Load database path
        db_path = self.config.get('ip_database_path', '')
        self.db_path_var.set(db_path if db_path else '')

        # Update API status
        if ipinfo_token:
            self.api_status_label.config(text="✓ Token configured", foreground="green")
        else:
            self.api_status_label.config(text="✗ Token not found", foreground="red")

        # Update local DB status
        self._update_local_db_status()

        self._on_source_change()

    def _update_local_db_status(self):
        """Update local database status display."""
        db_path = self.db_path_var.get()

        if db_path and os.path.exists(db_path):
            try:
                db = LocalIPDatabase(db_path)
                if db.is_valid():
                    stats = db.get_stats()
                    v4 = stats.get('ipv4_records', stats['total_records'])
                    v6 = stats.get('ipv6_records', 0)
                    self.local_status_label.config(
                        text=f"✓ Database ready ({stats['total_records']:,} records: {v4:,} IPv4, {v6:,} IPv6)",
                        foreground="green")
                else:
                    self.local_status_label.config(
                        text="⚠ Database exists but is empty",
                        foreground="orange")
                db.close()
            except Exception as e:
                self.local_status_label.config(
                    text=f"✗ Database error: {e}",
                    foreground="red")
        elif db_path:
            self.local_status_label.config(
                text="✗ Database file not found",
                foreground="red")
        else:
            self.local_status_label.config(
                text="✗ No database selected",
                foreground="red")

    def _on_source_change(self):
        """Handle data source selection change."""
        is_local = self.source_var.get() == "local"

        # Visual feedback for selection
        if is_local:
            self.local_frame.config(style="Selected.TLabelframe")
        else:
            self.api_frame.config(style="Selected.TLabelframe")

    def _open_db_setup(self):
        """Open the database setup window."""
        DatabaseSetupWindow(self.window, on_complete=self._update_local_db_status)

    def _save_settings(self):
        """Save settings and close."""
        use_local = self.source_var.get() == "local"
        db_path = self.db_path_var.get().strip()

        # Validate selection
        if use_local:
            if not db_path:
                messagebox.showwarning("Warning",
                                       "Please select a database file or create a new one.")
                return
            if not os.path.exists(db_path):
                messagebox.showwarning("Warning",
                                       f"Database file not found:\n{db_path}\n\n"
                                       "Please select an existing database or create a new one.")
                return

            # Verify database is valid
            try:
                db = LocalIPDatabase(db_path)
                if not db.is_valid():
                    messagebox.showwarning("Warning",
                                           "Selected database is empty. Please import data first.")
                    db.close()
                    return
                db.close()
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read database:\n{e}")
                return
        else:
            if not ipinfo_token:
                messagebox.showwarning("Warning",
                                       "API token not configured. Please set IPINFO_TOKEN in .env file.")
                return

        # Save settings
        if db_path:
            self.config.set_ip_database_path(db_path)
        self.config.set_use_local_db(use_local)

        if self.on_settings_changed:
            self.on_settings_changed()

        self.window.destroy()


def run_in_thread(func, *args):
    """Runs a function in a separate thread."""
    thread = Thread(target=func, args=args)
    thread.start()


def process_file(file_path, save_directory, in_process_window, root, export_format='xlsx'):
    """Processes the selected file and fetches data for each IP."""
    global result_list
    result_list = []
    ip_data = extract_ips_from_file(file_path)
    total_ips = len(ip_data)
    api_limit_hit = False
    scanned_count = 0

    has_count = any(item['count'] is not None for item in ip_data)

    logging.info(f"Found {total_ips} IPs to process.")

    for index, item in enumerate(ip_data, start=1):
        ip = item['ip']
        count = item['count']

        if is_valid_ipv4(ip) or is_valid_ipv6(ip):
            logging.info(f"Processing IP {index}/{total_ips}: {ip}")
            try:
                response = fetch_ip_data(ip)
                if response:
                    update_result_list(response, count if has_count else None)
                    scanned_count += 1
            except APILimitReached:
                api_limit_hit = True
                logging.warning(f"API limit reached after scanning {scanned_count} IPs.")
                break
        else:
            logging.warning(f"Invalid IP address skipped: {ip}")

    in_process_window.destroy()

    if result_list:
        save_path = write_to_excel(save_directory, partial=api_limit_hit,
                                   has_count=has_count, export_format=export_format)

        if api_limit_hit:
            message = (f"API limit reached!\n\n"
                       f"Successfully scanned: {scanned_count}/{total_ips} IPs\n"
                       f"Partial results saved at:\n{save_path}\n\n"
                       f"Do you want to close the program?")
            if messagebox.askyesno("Partial Results Saved", message):
                root.quit()
                root.destroy()
        else:
            if messagebox.askyesno("Success", f"Results saved at: {save_path}\nDo you want to close the program?"):
                root.quit()
                root.destroy()
    else:
        messagebox.showerror("Error", "No IPs were successfully scanned. No file was saved.")


def process_ips(ip_list, save_directory, in_process_window, root, export_format='xlsx'):
    """Process a list of IPs (one or more)."""
    global result_list
    result_list = []
    total_ips = len(ip_list)
    api_limit_hit = False
    scanned_count = 0
    invalid_count = 0

    for index, ip in enumerate(ip_list, start=1):
        if is_valid_ipv4(ip) or is_valid_ipv6(ip):
            logging.info(f"Scanning IP {index}/{total_ips}: {ip}")
            try:
                response = fetch_ip_data(ip)
                if response:
                    update_result_list(response)
                    scanned_count += 1
            except APILimitReached:
                api_limit_hit = True
                break
        else:
            invalid_count += 1
            logging.warning(f"Invalid IP address skipped: {ip}")

    in_process_window.destroy()

    if result_list:
        save_path = write_to_excel(save_directory, partial=api_limit_hit, export_format=export_format)

        if api_limit_hit:
            message = (f"API limit reached!\n\n"
                       f"Successfully scanned: {scanned_count}/{total_ips} IPs\n"
                       f"Invalid IPs skipped: {invalid_count}\n"
                       f"Partial results saved at:\n{save_path}")
            if messagebox.askyesno("Partial Results Saved", message + "\n\nClose program?"):
                root.quit()
                root.destroy()
        else:
            success_msg = f"Results saved at: {save_path}"
            if invalid_count > 0:
                success_msg += f"\n\nNote: {invalid_count} invalid IP(s) were skipped."
            if messagebox.askyesno("Success", success_msg + "\n\nClose program?"):
                root.quit()
                root.destroy()
    else:
        messagebox.showerror("Error", "No IPs were successfully scanned.")


class IPCheckerApp:
    """Main application class."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IP Info Checker")
        self.root.geometry("500x350")

        self.config = get_config()

        self._create_widgets()
        self._update_mode_display()

        # Check for first-time setup
        self.root.after(100, self._check_first_run)

    def _create_widgets(self):
        """Create main window widgets."""
        self.font = Font(self.root, size=12)

        # Main frame
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(expand=True, fill=tk.BOTH)

        # Title
        tk.Label(frame, text="IP Info Checker",
                 font=("Helvetica", 20, "bold"), fg="blue").pack(pady=10)

        # Mode indicator
        self.mode_frame = tk.Frame(frame, bg="#e0e0e0", padx=10, pady=5)
        self.mode_frame.pack(fill=tk.X, pady=10)

        self.mode_label = tk.Label(self.mode_frame, text="Mode: ",
                                   font=("Helvetica", 11), bg="#e0e0e0")
        self.mode_label.pack(side=tk.LEFT)

        self.mode_value_label = tk.Label(self.mode_frame, text="API",
                                         font=("Helvetica", 11, "bold"), bg="#e0e0e0")
        self.mode_value_label.pack(side=tk.LEFT)

        tk.Button(self.mode_frame, text="⚙ Settings",
                  command=self._open_settings, font=("Helvetica", 10)).pack(side=tk.RIGHT)

        # Separator
        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=10)

        # File selection
        tk.Label(frame, text="Select a CSV file with IP addresses:",
                 font=self.font).pack(pady=5)
        tk.Button(frame, text="Select File", command=self._select_file,
                  font=self.font).pack(pady=5)

        # Enter IPs manually (one or more)
        tk.Label(frame, text="Or enter IP addresses manually:",
                 font=self.font).pack(pady=5)
        tk.Button(frame, text="Enter IPs", command=self._input_ips,
                  font=self.font).pack(pady=5)

    def _update_mode_display(self):
        """Update the mode indicator display."""
        use_local = self.config.get_use_local_db()

        if use_local:
            self.mode_value_label.config(text="Local Database", fg="green")
        else:
            self.mode_value_label.config(text="API (ipinfo.io)", fg="blue")

    def _check_first_run(self):
        """Check if this is the first run and prompt for setup."""
        use_local = self.config.get_use_local_db()
        db_path = self.config.get_ip_database_path()

        # Check if neither API nor local DB is configured
        if not ipinfo_token and (not db_path or not os.path.exists(db_path)):
            result = messagebox.askyesno(
                "First Time Setup",
                "Welcome to IP Info Checker!\n\n"
                "No data source is configured.\n\n"
                "Would you like to set up a local IP database now?\n\n"
                "(Click 'No' to configure API token instead)"
            )

            if result:
                DatabaseSetupWindow(self.root, on_complete=self._on_first_setup_complete)
            else:
                messagebox.showinfo(
                    "API Setup",
                    "To use the API mode:\n\n"
                    "1. Create a .env file in the application directory\n"
                    "2. Add: IPINFO_TOKEN=your_token_here\n"
                    "3. Restart the application"
                )

    def _on_first_setup_complete(self):
        """Handle completion of first-time setup."""
        db_path = self.config.get_ip_database_path()
        if db_path and os.path.exists(db_path):
            self.config.set_use_local_db(True)
            self._update_mode_display()
            messagebox.showinfo("Setup Complete",
                                "Local database configured successfully!\n\n"
                                "You can now scan IP addresses offline.")

    def _open_settings(self):
        """Open settings window."""
        SettingsWindow(self.root, on_settings_changed=self._update_mode_display)

    def _select_file(self):
        """Open file selection dialog."""
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            save_directory = filedialog.askdirectory(title="Select Directory to Save Results")
            if save_directory:
                # Ask for export format
                export_format = get_export_format_choice(self.root)
                if export_format:
                    in_process_window = tk.Toplevel(self.root)
                    in_process_window.title("Processing")
                    in_process_window.geometry("300x150")
                    tk.Label(in_process_window, text="In process... Please wait",
                             font=("Helvetica", 14)).pack(expand=True)
                    run_in_thread(process_file, file_path, save_directory,
                                  in_process_window, self.root, export_format)

    def _input_ips(self):
        """Open IP input dialog (supports one or more IPs)."""

        def on_submit():
            text_content = ip_text_area.get("1.0", tk.END)
            ip_list = [ip.strip() for ip in text_content.split('\n') if ip.strip()]

            if not ip_list:
                messagebox.showwarning("Warning", "Please enter at least one IP address.")
                return

            save_directory = filedialog.askdirectory(title="Select Directory to Save Results")
            if save_directory:
                # Ask for export format
                export_format = get_export_format_choice(self.root)
                if export_format:
                    ip_input_window.destroy()
                    in_process_window = tk.Toplevel(self.root)
                    in_process_window.title("Processing")
                    in_process_window.geometry("300x150")
                    tk.Label(in_process_window,
                             text=f"Scanning {len(ip_list)} IP(s)...\nPlease wait",
                             font=("Helvetica", 14)).pack(expand=True)
                    run_in_thread(process_ips, ip_list, save_directory,
                                  in_process_window, self.root, export_format)

        def clear_text():
            ip_text_area.delete("1.0", tk.END)

        ip_input_window = tk.Toplevel(self.root)
        ip_input_window.title("Enter IP Addresses")
        ip_input_window.geometry("400x350")

        tk.Label(ip_input_window, text="Enter IP addresses (one per line):",
                 font=("Helvetica", 12)).pack(pady=10)

        text_frame = tk.Frame(ip_input_window)
        text_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        ip_text_area = tk.Text(text_frame, font=("Helvetica", 11), height=10,
                               width=35, yscrollcommand=scrollbar.set)
        ip_text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=ip_text_area.yview)

        button_frame = tk.Frame(ip_input_window)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="Clear", command=clear_text,
                  font=("Helvetica", 11), width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Submit", command=on_submit,
                  font=("Helvetica", 11), width=10).pack(side=tk.LEFT, padx=5)

    def run(self):
        """Run the application."""
        self.root.mainloop()


if __name__ == "__main__":
    app = IPCheckerApp()
    app.run()