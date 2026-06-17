import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
import re
import datetime
import queue
import threading

# ---------------------------------------------------------------------------
# Coordinate parsing
# ---------------------------------------------------------------------------

def dms_to_decimal(degree, minute, second, direction):
    """
    Converts coordinates from Degrees, Minutes, Seconds (DMS) format to Decimal Degrees (DD) format.

    Args:
        degree (str or float): Degrees component of the coordinate.
        minute (str or float): Minutes component of the coordinate.
        second (str or float): Seconds component of the coordinate.
        direction (str): Cardinal direction ('N', 'S', 'E', 'W').

    Returns:
        float: The coordinate in Decimal Degrees format, rounded to 7 decimal places.
    """
    decimal = float(degree) + float(minute) / 60 + float(second) / 3600
    if direction in ['S', 'W']:
        decimal *= -1
    return round(decimal, 7)


def parse_coordinate(coord):
    """
    Parses a coordinate string and converts it to Decimal Degrees (DD) format.

    Supports various formats including:
    - Decimal Degrees with optional cardinal direction (e.g., "45.123N", "-120.456").
    - Degrees, Minutes, Seconds (DMS) with cardinal direction (e.g., "45°30'15\"N").
    - Degrees and Minutes with cardinal direction (e.g., "45°30'N").

    Args:
        coord (str): The coordinate string to parse.

    Returns:
        float or None: The coordinate in Decimal Degrees format, or None if parsing fails.
    """
    if pd.isnull(coord):
        return None

    # If already a number, just return as float
    if isinstance(coord, (int, float)):
        return float(coord)

    coord = str(coord)
    coord = coord.replace('’', "'").replace('″', '"').replace('“', '"').replace("''", '"')
    coord = coord.strip()

    try:
        # Handle Decimal Degrees format
        if re.match(r'^-?\d+(\.\d+)?[NSEW]?$', coord):
            if coord[-1] in ['N', 'S', 'E', 'W']:
                val = float(coord[:-1])
                if coord[-1] in ['S', 'W']:
                    val *= -1
                return round(val, 7)
            else:
                return round(float(coord), 7)
    except ValueError:
        pass

    # Handle Degrees, Minutes, Seconds (DMS) format
    match = re.match(
        r'(?P<deg>\d+)[°:]?\s*(?P<min>\d+)[\':]?\s*(?P<sec>\d+(?:\.\d+)?)[\"’]?\s*(?P<dir>[NSEW])',
        coord, re.IGNORECASE)
    if match:
        parts = match.groupdict()
        return dms_to_decimal(parts['deg'], parts['min'], parts['sec'], parts['dir'].upper())

    # Handle Degrees and Minutes format
    match = re.match(
        r'(?P<deg>\d+)[°:]?\s*(?P<min>\d+)[\'’]?\s*(?P<dir>[NSEW])',
        coord, re.IGNORECASE)
    if match:
        parts = match.groupdict()
        return dms_to_decimal(parts['deg'], parts['min'], 0, parts['dir'].upper())

    return None


def convert_coordinates(df, x_col, y_col):
    lon_converted = df[x_col].apply(parse_coordinate)
    lat_converted = df[y_col].apply(parse_coordinate)
    df['Longitude_Converted'] = lon_converted
    df['Latitude_Converted'] = lat_converted
    df['Convert_Status'] = [
        'Yes' if (not pd.isnull(lon) and not pd.isnull(lat)) else 'No'
        for lon, lat in zip(lon_converted, lat_converted)
    ]
    return df


def process_file(filepath, x_col, y_col):
    df = pd.read_excel(filepath)
    df = convert_coordinates(df, x_col, y_col)
    date_stamp = datetime.datetime.now().strftime('%Y%m%d')
    out_path = os.path.splitext(filepath)[0] + f'_converted_{date_stamp}.xlsx'
    df.to_excel(out_path, index=False)
    total = len(df)
    converted = int((df['Convert_Status'] == 'Yes').sum())
    failed = total - converted
    return out_path, total, converted, failed


def detect_coordinate_columns(columns):
    # Ensure all column names are strings
    columns = [str(col) for col in columns]
    # Latitude candidates: prioritize 'lat', fallback to 'y'
    lat_candidates = [col for col in columns if 'lat' in col.lower()]
    if not lat_candidates:
        lat_candidates = [col for col in columns if col.lower() == 'y']
    # Longitude candidates: prioritize 'lon'/'lng', fallback to 'x'
    lon_candidates = [col for col in columns if 'lon' in col.lower() or 'lng' in col.lower()]
    if not lon_candidates:
        lon_candidates = [col for col in columns if col.lower() == 'x']
    # Fallbacks
    lat_default = lat_candidates[0] if lat_candidates else columns[0] if columns else ''
    lon_default = lon_candidates[0] if lon_candidates else columns[1] if len(columns) > 1 else columns[0] if columns else ''
    return lat_default, lon_default


# ---------------------------------------------------------------------------
# Theme / palette
# ---------------------------------------------------------------------------

COLORS = {
    'bg': '#f1f5fb',
    'card': '#ffffff',
    'primary': '#4f46e5',
    'primary_active': '#4338ca',
    'on_primary': '#ffffff',
    'text': '#1e293b',
    'muted': '#64748b',
    'heading': '#1e2a52',
    'border': '#d8def0',
    'info_bg': '#eef2ff',
    'accent': '#0ea5e9',
    'success': '#16a34a',
    'danger': '#dc2626',
}

FONT_FAMILY = 'Segoe UI'


# ---------------------------------------------------------------------------
# UI callbacks
# ---------------------------------------------------------------------------

def set_status(message, kind='info'):
    color = {
        'info': COLORS['muted'],
        'success': COLORS['success'],
        'error': COLORS['danger'],
    }.get(kind, COLORS['muted'])
    status_var.set(message)
    status_label.configure(foreground=color)


# Background work (reading / converting Excel) runs on worker threads so the
# window stays responsive and the progress bar can animate. Workers never touch
# widgets directly — they push a callable onto this queue, which the Tk main
# loop drains via pump_ui_queue().
_ui_queue = queue.Queue()


def pump_ui_queue():
    while True:
        try:
            callback = _ui_queue.get_nowait()
        except queue.Empty:
            break
        callback()
    root.after(80, pump_ui_queue)


def post_to_ui(callback):
    _ui_queue.put(callback)


def set_busy(busy, message=None, kind='info'):
    """Show/hide the activity indicator and lock controls while working."""
    if message is not None:
        set_status(message, kind)
    if busy:
        progress.grid()
        progress.start(12)
        browse_btn.state(['disabled'])
        convert_btn.state(['disabled'])
    else:
        progress.stop()
        progress.grid_remove()
        browse_btn.state(['!disabled'])
        convert_btn.state(['!disabled'] if entry_file_var.get() else ['disabled'])


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
    if not file_path:
        return
    entry_file_var.set(file_path)
    name = os.path.basename(file_path)
    set_busy(True, f'Reading {name}…')

    def worker():
        try:
            df = pd.read_excel(file_path)
            columns = [str(c) for c in df.columns]
            lat_default, lon_default = detect_coordinate_columns(columns)
            row_count = len(df)

            def done():
                combo_x['values'] = columns
                combo_y['values'] = columns
                combo_x.set(lon_default)
                combo_y.set(lat_default)
                set_busy(False, f'Loaded {name} — {len(columns)} columns, {row_count} rows.', 'info')

            post_to_ui(done)
        except Exception as e:
            def fail():
                entry_file_var.set('')
                set_busy(False, 'Could not read that file.', 'error')
                messagebox.showerror('Error', f'Failed to read Excel file: {e}')

            post_to_ui(fail)

    threading.Thread(target=worker, daemon=True).start()


def run_conversion():
    file_path = entry_file_var.get()
    x_col = combo_x.get()
    y_col = combo_y.get()
    if not (file_path and x_col and y_col):
        messagebox.showerror('Error', 'Please choose a file and both coordinate columns.')
        set_status('Missing inputs.', 'error')
        return
    set_busy(True, 'Converting…')

    def worker():
        try:
            out_path, total, converted, failed = process_file(file_path, x_col, y_col)

            def done():
                set_busy(False, f'Done — {converted} of {total} rows converted, {failed} failed.', 'success')
                messagebox.showinfo(
                    'Success',
                    f'Converted {converted} of {total} rows ({failed} failed).\n\nSaved as:\n{out_path}'
                )

            post_to_ui(done)
        except Exception as e:
            def fail():
                set_busy(False, 'Conversion failed.', 'error')
                messagebox.showerror('Error', str(e))

            post_to_ui(fail)

    threading.Thread(target=worker, daemon=True).start()


def close_app():
    root.destroy()


def center_window(window, width, height):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


# ---------------------------------------------------------------------------
# UI construction
# ---------------------------------------------------------------------------

def main():
    global root, status_var, status_label, entry_file_var, combo_x, combo_y
    global convert_btn, browse_btn, progress

    root = tk.Tk()
    root.title('Converta — Coordinate Cleaner')
    root.configure(bg=COLORS['bg'])
    center_window(root, 600, 560)
    root.minsize(560, 520)

    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except tk.TclError:
        pass

    style.configure('TFrame', background=COLORS['bg'])
    style.configure('Card.TFrame', background=COLORS['card'])
    style.configure('TLabel', background=COLORS['card'], foreground=COLORS['text'], font=(FONT_FAMILY, 10))
    style.configure('Bg.TLabel', background=COLORS['bg'], foreground=COLORS['text'], font=(FONT_FAMILY, 10))
    style.configure('Brand.TLabel', background=COLORS['bg'], foreground=COLORS['heading'], font=(FONT_FAMILY, 18, 'bold'))
    style.configure('Org.TLabel', background=COLORS['bg'], foreground=COLORS['muted'], font=(FONT_FAMILY, 9))
    style.configure('Heading.TLabel', background=COLORS['card'], foreground=COLORS['heading'], font=(FONT_FAMILY, 10, 'bold'))
    style.configure('Info.TLabel', background=COLORS['info_bg'], foreground=COLORS['text'], font=(FONT_FAMILY, 9))
    style.configure('Status.TLabel', background=COLORS['bg'], foreground=COLORS['muted'], font=(FONT_FAMILY, 9))

    style.configure('TEntry', fieldbackground=COLORS['card'], bordercolor=COLORS['border'], padding=6)
    style.configure('TCombobox', fieldbackground=COLORS['card'], bordercolor=COLORS['border'], padding=6)
    style.map('TCombobox', fieldbackground=[('readonly', COLORS['card'])])

    style.configure('Primary.TButton', background=COLORS['primary'], foreground=COLORS['on_primary'],
                    font=(FONT_FAMILY, 10, 'bold'), borderwidth=0, focusthickness=0, padding=(18, 9))
    style.map('Primary.TButton',
              background=[('active', COLORS['primary_active']), ('disabled', '#b9bdf0')],
              foreground=[('disabled', '#eef0ff')])
    style.configure('Secondary.TButton', background=COLORS['card'], foreground=COLORS['primary'],
                    font=(FONT_FAMILY, 10, 'bold'), borderwidth=1, bordercolor=COLORS['primary'], padding=(18, 9))
    style.map('Secondary.TButton', background=[('active', COLORS['info_bg'])])
    style.configure('Browse.TButton', background=COLORS['info_bg'], foreground=COLORS['primary'],
                    font=(FONT_FAMILY, 9, 'bold'), borderwidth=0, padding=(12, 6))
    style.map('Browse.TButton', background=[('active', '#e0e6ff')])

    style.configure('Converta.Horizontal.TProgressbar', troughcolor=COLORS['info_bg'],
                    background=COLORS['primary'], bordercolor=COLORS['info_bg'],
                    lightcolor=COLORS['primary'], darkcolor=COLORS['primary'])

    # --- Header / brand bar ---
    header = ttk.Frame(root, style='TFrame', padding=(24, 20, 24, 8))
    header.pack(fill='x')

    brand_canvas = tk.Canvas(header, width=40, height=40, bg=COLORS['bg'], highlightthickness=0)
    brand_canvas.create_oval(6, 4, 34, 32, fill=COLORS['primary'], outline='')
    brand_canvas.create_polygon(20, 36, 11, 22, 29, 22, fill=COLORS['primary'], outline='')
    brand_canvas.create_oval(15, 11, 25, 21, fill='#ffffff', outline='')
    brand_canvas.grid(row=0, column=0, rowspan=2, padx=(0, 12))
    ttk.Label(header, text='Converta', style='Brand.TLabel').grid(row=0, column=1, sticky='w')
    ttk.Label(header, text='Coordinate Cleaner · a maplumi tool', style='Org.TLabel').grid(row=1, column=1, sticky='w')

    # --- Card body ---
    card = ttk.Frame(root, style='Card.TFrame', padding=24)
    card.pack(fill='both', expand=True, padx=24, pady=(8, 12))
    card.columnconfigure(1, weight=1)

    info_text = (
        'Supported formats:\n'
        '   • Decimal degrees:  45.123N,  -120.456,  120.456W\n'
        "   • DMS:  45°30'15\"N,  120°30'15\"W\n"
        "   • Degrees & minutes:  45°30'N,  120°30'W\n"
        '   • Plain numbers:  45.123,  -120.456\n'
        'Rows that cannot be converted are flagged in the output.'
    )
    info_frame = tk.Frame(card, bg=COLORS['info_bg'], highlightbackground=COLORS['primary'],
                          highlightthickness=0, bd=0)
    info_frame.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 18))
    tk.Frame(info_frame, bg=COLORS['primary'], width=4).pack(side='left', fill='y')
    ttk.Label(info_frame, text=info_text, style='Info.TLabel', justify='left').pack(
        side='left', fill='x', padx=14, pady=12)

    # Excel file input
    ttk.Label(card, text='Excel file', style='Heading.TLabel').grid(row=1, column=0, sticky='w', pady=(2, 4))
    entry_file_var = tk.StringVar()
    entry_file = ttk.Entry(card, textvariable=entry_file_var)
    entry_file.grid(row=2, column=0, columnspan=2, sticky='ew', padx=(0, 8))
    browse_btn = ttk.Button(card, text='Browse…', style='Browse.TButton', command=select_file)
    browse_btn.grid(row=2, column=2, sticky='e')

    # Longitude column
    ttk.Label(card, text='Longitude / X column', style='Heading.TLabel').grid(
        row=3, column=0, sticky='w', pady=(16, 4))
    combo_x = ttk.Combobox(card, state='readonly')
    combo_x.grid(row=4, column=0, columnspan=3, sticky='ew')

    # Latitude column
    ttk.Label(card, text='Latitude / Y column', style='Heading.TLabel').grid(
        row=5, column=0, sticky='w', pady=(16, 4))
    combo_y = ttk.Combobox(card, state='readonly')
    combo_y.grid(row=6, column=0, columnspan=3, sticky='ew')

    # Buttons
    button_row = ttk.Frame(card, style='Card.TFrame')
    button_row.grid(row=7, column=0, columnspan=3, sticky='ew', pady=(26, 0))
    button_row.columnconfigure(0, weight=1)
    convert_btn = ttk.Button(button_row, text='Convert', style='Primary.TButton', command=run_conversion)
    convert_btn.grid(row=0, column=1, padx=(0, 10))
    convert_btn.state(['disabled'])
    close_btn = ttk.Button(button_row, text='Close', style='Secondary.TButton', command=close_app)
    close_btn.grid(row=0, column=2)

    # Activity indicator — hidden until a read/convert is in progress
    progress = ttk.Progressbar(card, mode='indeterminate', style='Converta.Horizontal.TProgressbar')
    progress.grid(row=8, column=0, columnspan=3, sticky='ew', pady=(16, 0))
    progress.grid_remove()

    # --- Status bar ---
    status_var = tk.StringVar(value='Choose an Excel file to get started.')
    status_label = ttk.Label(root, textvariable=status_var, style='Status.TLabel', anchor='w')
    status_label.pack(fill='x', padx=26, pady=(0, 14))

    root.after(80, pump_ui_queue)
    root.mainloop()


if __name__ == '__main__':
    main()
