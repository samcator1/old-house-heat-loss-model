"""
Configuration, styling theme, formats, and baseline assumptions for the Old House Heat Loss Model.
"""

from typing import Dict, Any

# Visual Design Theme (Modern Clean Engineering Palette)
THEME = {
    # Headers
    "PRIMARY_HEADER_BG": {"red": 0.09, "green": 0.16, "blue": 0.24},   # Deep Slate/Navy (#172A3A)
    "SECONDARY_HEADER_BG": {"red": 0.13, "green": 0.23, "blue": 0.35}, # Steel Navy (#223A59)
    "SECTION_HEADER_BG": {"red": 0.18, "green": 0.38, "blue": 0.58},   # Architectural Blue (#2E6194)
    "HEADER_TEXT": {"red": 1.0, "green": 1.0, "blue": 1.0},            # Pure White
    
    # Cells
    "INPUT_BG": {"red": 0.88, "green": 0.94, "blue": 0.99},            # Soft Sky Blue / User Editable (#E0F0FD)
    "CALC_BG": {"red": 0.94, "green": 0.97, "blue": 0.94},             # Soft Sage / Calculated (#F0F7F0)
    "ACCENT_BG": {"red": 1.0, "green": 0.97, "blue": 0.88},            # Soft Amber / Key Metric (#FFF7E0)
    "TOTAL_BG": {"red": 0.92, "green": 0.95, "blue": 0.98},             # Muted Steel Gray (#EBF2FA)
    "ZEBRA_BG": {"red": 0.98, "green": 0.99, "blue": 1.0},             # Ultra Light Slate (#F8FAFC)
    "CARD_BG": {"red": 0.95, "green": 0.97, "blue": 0.99},              # Dashboard Card (#F1F5F9)
    "CARD_HEADER_BG": {"red": 0.18, "green": 0.28, "blue": 0.38},       # Dark Card Header
    
    # Text
    "DARK_TEXT": {"red": 0.06, "green": 0.09, "blue": 0.16},           # Off-Black (#0F172A)
    "MUTED_TEXT": {"red": 0.39, "green": 0.45, "blue": 0.55},          # Muted Slate (#64748B)
    "ACCENT_TEXT": {"red": 0.02, "green": 0.44, "blue": 0.70},         # Blue Accent
    "SUCCESS_TEXT": {"red": 0.06, "green": 0.50, "blue": 0.28},        # Emerald Green
}

# Standard Number Format Strings for Google Sheets
FORMATS = {
    "CURRENCY_GBP": "£#,##0",
    "CURRENCY_GBP_DEC": "£#,##0.00",
    "PERCENT": "0.0%",
    "PERCENT_INT": "0%",
    "DECIMAL_1": "0.0",
    "DECIMAL_2": "0.00",
    "INTEGER": "#,##0",
    "WATTS": '#,##0 "W"',
    "KW": '0.0 "kW"',
    "KWH": '#,##0 "kWh"',
    "M2": '0.0 "m²"',
    "M3": '0.0 "m³"',
    "W_M2": '0.0 "W/m²"',
    "U_VAL": '0.00 "W/m²K"',
    "ACH": '0.00 "ACH"',
    "TEMP_C": '0.0 "°C"',
    "TONNES_CO2": '0.0 "t CO₂"',
}

# Default Building Zones for Final Build
DEFAULT_ZONES = [
    {
        "name": "Georgian end",
        "length": 14.75,
        "width": 6.25,
        "floors": 2.0,
        "h1": 2.80,
        "h2": 2.80,
        "pct_ext_wall": 0.65,
        "u_wall": 1.40,
        "pct_glazing": 0.20,
        "u_window": 4.80,
        "u_ground": 0.80,
        "u_roof": 0.18,
        "zone_type": "Old House"
    },
    {
        "name": "Thatched gable ended",
        "length": 6.00,
        "width": 16.50,
        "floors": 2.5,
        "h1": 2.80,
        "h2": 2.60,
        "pct_ext_wall": 0.37,
        "u_wall": 1.80,
        "pct_glazing": 0.20,
        "u_window": 4.80,
        "u_ground": 0.60,
        "u_roof": 0.30,
        "zone_type": "Old House"
    },
    {
        "name": "New build",
        "length": 10.00,
        "width": 13.00,
        "floors": 2.0,
        "h1": 2.80,
        "h2": 2.60,
        "pct_ext_wall": 0.28,
        "u_wall": 0.22,
        "pct_glazing": 0.25,
        "u_window": 1.40,
        "u_ground": 0.15,
        "u_roof": 0.15,
        "zone_type": "New Build"
    },
    {
        "name": "Orangery",
        "length": 6.00,
        "width": 10.00,
        "floors": 1.0,
        "h1": 2.80,
        "h2": 0.00,
        "pct_ext_wall": 0.50,
        "u_wall": 0.25,
        "pct_glazing": 0.70,
        "u_window": 1.40,
        "u_ground": 0.15,
        "u_roof": 1.50,
        "zone_type": "New Build"
    },
    {
        "name": "Lean to West",
        "length": 4.00,
        "width": 6.00,
        "floors": 1.0,
        "h1": 2.60,
        "h2": 0.00,
        "pct_ext_wall": 0.50,
        "u_wall": 2.20,
        "pct_glazing": 0.30,
        "u_window": 4.80,
        "u_ground": 1.10,
        "u_roof": 2.00,
        "zone_type": "Old House"
    },
    {
        "name": "Lean to North",
        "length": 3.00,
        "width": 6.00,
        "floors": 1.0,
        "h1": 2.60,
        "h2": 2.00,
        "pct_ext_wall": 0.17,
        "u_wall": 2.20,
        "pct_glazing": 0.30,
        "u_window": 4.80,
        "u_ground": 1.10,
        "u_roof": 2.00,
        "zone_type": "Old House"
    }
]
