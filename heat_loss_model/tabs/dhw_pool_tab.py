"""
Domestic Hot Water (DHW) & Swimming Pool Thermal Requirements Tab ('3_DHW_and_Pool').
Models water heating volume, standing vessel losses, secondary pumped circulation loops,
Legionella pasteurization cycles, and outdoor swimming pool seasonal requirements.
"""

from typing import Tuple, List, Dict, Any
import gspread
from ..config import THEME, FORMATS
from ..formatting import (
    create_repeat_cell_request,
    create_set_column_width_request,
    create_freeze_pane_request,
    create_merge_cells_request
)

def build_dhw_pool_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 3_DHW_and_Pool worksheet."""
    tab_name = "3_DHW_and_Pool"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=55, cols=8)

    total_cols = 5
    total_rows = 52
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Banners
    grid[0][0] = "3_DHW_and_Pool: Domestic Hot Water & Swimming Pool Thermal Analysis"
    grid[1][0] = "Detailed DHW storage vessel recharge, standing/circulation losses, Legionella boost, and seasonal pool thermal demand."

    # Row 3: Section 1 Header
    grid[2][0] = "SECTION A: DOMESTIC HOT WATER (DHW) THERMAL DEMAND & STORAGE"
    grid[3][0] = "Parameter Description"
    grid[3][1] = "Value"
    grid[3][2] = "Unit"
    grid[3][3] = "Formula / Reference"
    grid[3][4] = "Technical Notes & Guidance"

    dhw_rows = [
        ("Occupancy", "='1_Inputs'!$C$20", "people", "Inputs C20", "Design building occupancy"),
        ("Hot water consumption per person", "='1_Inputs'!$C$21", "L/person/day", "Inputs C21", "UK average daily consumption"),
        ("Total daily hot water consumption", "=B5*B6", "L/day", "=B5*B6", "Whole-house daily volume"),
        ("Total annual hot water volume", "=B7*365", "L/year", "=B7*365", "Annual draw-off volume"),
        ("Hot water storage temperature", "='1_Inputs'!$C$22", "°C", "Inputs C22", "Optimal heat pump cylinder setpoint"),
        ("Mains cold water temperature", "='1_Inputs'!$C$23", "°C", "Inputs C23", "UK average mains feed temp"),
        ("Hot water temperature lift (ΔT)", "=B9-B10", "K", "=B9-B10", "Thermal lift required"),
        ("Specific heat capacity of water", "4186", "J/kg·K", "Constant", "Water physical property"),
        ("Theoretical net water heating energy", "=(B8*B12*B11)/3600000", "kWh/year", "=(B8*4186*ΔT)/3.6M", "Base energy delivered to tap"),
        ("Cylinder standing loss allowance", "='1_Inputs'!$C$25", "%", "Inputs C25", "Standing heat loss through vessel insulation"),
        ("Primary cylinder standing losses", "=B13*B14", "kWh/year", "=B13*B14", "Annual cylinder standing heat loss"),
        ("Secondary circulation loop installed", "='1_Inputs'!$C$26", "Yes/No", "Inputs C26", "Continuous or timed pumped loop"),
        ("Secondary circulation loop losses", "=IF(B16=\"Yes\", '1_Inputs'!$C$27*365, 0)", "kWh/year", "=IF(B16=\"Yes\", C27*365, 0)", "Pipework distribution standing losses"),
        ("Legionella weekly pasteurization boost", "='1_Inputs'!$C$28", "kWh/year", "Inputs C28", "Weekly 60-65°C electric immersion pasteurization"),
        ("TOTAL ANNUAL DELIVERED DHW ENERGY", "=B13+B15+B17+B18", "kWh/year", "=SUM(B13,B15,B17,B18)", "Total heat pump/boiler DHW thermal requirement"),
        ("DHW cylinder storage capacity", "='1_Inputs'!$C$24", "Litres", "Inputs C24", "Recommended buffer for heat pump smooth cycling"),
        ("Cold vessel full reheat energy", "=(B20*B12*B11)/3600000", "kWh", "=(B20*4186*ΔT)/3.6M", "Thermal storage content of cylinder"),
        ("Target vessel reheat duration", "='1_Inputs'!$C$29", "Hours", "Inputs C29", "Target reheat time from empty"),
        ("Peak DHW Reheat Plant Capacity Required", "=B21/B22", "kW", "=B21/B22", "Dedicated heat pump output required during DHW priority")
    ]

    for i, r_data in enumerate(dhw_rows):
        r_idx = 4 + i
        grid[r_idx][0] = r_data[0]
        grid[r_idx][1] = r_data[1]
        grid[r_idx][2] = r_data[2]
        desc = r_data[3]
        grid[r_idx][3] = f"'{desc}" if desc.startswith("=") else desc
        grid[r_idx][4] = r_data[4]

    # Row 25: Section 2 Header
    grid[24][0] = "SECTION B: SWIMMING POOL THERMAL DEMAND & LOSS ANALYSIS"
    grid[25][0] = "Pool Parameter Description"
    grid[25][1] = "Value"
    grid[25][2] = "Unit"
    grid[25][3] = "Formula / Reference"
    grid[25][4] = "Technical Notes & Guidance"

    pool_rows = [
        ("Pool length", "='1_Inputs'!$C$32", "m", "Inputs C32", "Pool physical length"),
        ("Pool width", "='1_Inputs'!$C$33", "m", "Inputs C33", "Pool physical width"),
        ("Pool average depth", "='1_Inputs'!$C$34", "m", "Inputs C34", "Pool average depth"),
        ("Pool surface area", "=B27*B28", "m²", "B27 × B28", "Evaporation & surface loss area"),
        ("Pool water volume", "=B30*B29", "m³", "B30 × B29", "Total water volume"),
        ("Pool water mass", "=B31*1000", "kg", "B31 × 1000", "Mass of water (1000 kg/m³)"),
        ("Target pool water temperature", "='1_Inputs'!$C$35", "°C", "Inputs C35", "Heated pool setpoint"),
        ("Initial fill cold water temperature", "='1_Inputs'!$C$23", "°C", "Inputs C23", "Mains supply temp at fill"),
        ("Initial seasonal fill energy", "=(B32*B12*(B33-B34))/3600000", "kWh", "(Mass × 4186 × ΔT)/3.6M", "Energy to bring fresh fill to 24°C"),
        ("Pool heating season length", "='1_Inputs'!$C$36", "Days", "Inputs C36", "May to September season (~3 months)"),
        ("Daily covered hours", "='1_Inputs'!$C$39", "Hours/day", "Inputs C39", "Cover active duration"),
        ("Daily uncovered hours", "=24-B37", "Hours/day", "24 - B37", "Swimming / uncovered duration"),
        ("Covered heat loss rate", "='1_Inputs'!$C$37", "kW/m²", "Inputs C37", "Loss rate with thermal cover (~5-7 kW total)"),
        ("Uncovered heat loss rate", "='1_Inputs'!$C$38", "kW/m²", "Inputs C38", "Loss rate when open (~15 kW peak)"),
        ("Daily surface heat loss", "=B30*(B39*B37+B40*B38)", "kWh/day", "Area × (Cov_W × Hrs + Unc_W × Hrs)", "Daily maintenance thermal loss"),
        ("Seasonal surface heat loss", "=B41*B36", "kWh/year", "B41 × B36", "Maintenance heat over 90-day season"),
        ("TOTAL ANNUAL DELIVERED POOL HEAT", "=B35+B42", "kWh/year", "B35 + B42", "Total delivered energy to pool heat exchanger"),
        ("Peak uncovered pool heat loss rate", "=B30*B40", "kW", "B30 × B40", "Peak open pool heat exchanger requirement"),
        ("Covered steady-state heat loss rate", "=B30*B39", "kW", "B30 × B39", "Covered pool steady-state heat requirement")
    ]

    for i, r_data in enumerate(pool_rows):
        r_idx = 26 + i
        grid[r_idx][0] = r_data[0]
        grid[r_idx][1] = r_data[1]
        grid[r_idx][2] = r_data[2]
        desc = r_data[3]
        grid[r_idx][3] = f"'{desc}" if desc.startswith("=") else desc
        grid[r_idx][4] = r_data[4]

    # Row 47: Section C Header
    grid[46][0] = "SECTION C: CONSOLIDATED NON-SPACE HEATING THERMAL DEMAND"
    grid[47][0] = "Thermal End-Use"
    grid[47][1] = "Delivered Heat (kWh/yr)"
    grid[47][2] = "Daily Avg (kWh/day)"
    grid[47][3] = "Peak Plant (kW)"
    grid[47][4] = "Operating Season"

    grid[48] = ["Domestic Hot Water (DHW)", "=B19", "=B49/365", "=B23", "Year-round (365 days)"]
    grid[49] = ["Swimming Pool Heating", "=B43", "=B50/B36", "=B44", "Summer season (90 days)"]
    grid[50] = ["TOTAL NON-SPACE HEATING DEMAND", "=B49+B50", "=B51/365", "=B23+B44", "Combined hot water & leisure"]

    # Write grid to sheet
    ws.update(values=grid, range_name=f"A1:E{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze header
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 300))  # Description
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 130))  # Value
    fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 110))  # Unit
    fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 180))  # Formula
    fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 360))  # Notes

    # Banner
    fmt_reqs.append(create_merge_cells_request(ws.id, 0, 1, 0, total_cols))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 0, 1, 0, total_cols,
        bg_color=THEME["PRIMARY_HEADER_BG"],
        font_color=THEME["HEADER_TEXT"],
        bold=True,
        font_size=12,
        align="LEFT"
    ))

    fmt_reqs.append(create_merge_cells_request(ws.id, 1, 2, 0, total_cols))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 1, 2, 0, total_cols,
        bg_color=THEME["ZEBRA_BG"],
        font_color=THEME["MUTED_TEXT"],
        italic=True,
        font_size=9,
        align="LEFT"
    ))

    # Section Headers
    section_rows = [2, 24, 46]
    for sr in section_rows:
        fmt_reqs.append(create_merge_cells_request(ws.id, sr, sr + 1, 0, total_cols))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr, sr + 1, 0, total_cols,
            bg_color=THEME["SECTION_HEADER_BG"],
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=10,
            align="LEFT"
        ))
        # Column headers right below
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr + 1, sr + 2, 0, total_cols,
            bg_color=THEME["SECONDARY_HEADER_BG"],
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=9,
            align="CENTER"
        ))

    # Data rows formatting
    for r in range(4, 23):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 2, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="CENTER", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 5, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    for r in range(26, 45):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 2, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="CENTER", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 5, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    for r in range(48, 51):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 4, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 4, 5, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    # Key Totals Styling
    highlight_rows = [18, 42, 50]  # Total DHW (row 19), Total Pool (row 43), Grand Total (row 51)
    for hr in highlight_rows:
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, hr, hr + 1, 0, total_cols,
            bg_color=THEME["ACCENT_BG"],
            bold=True,
            font_size=10
        ))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, hr, hr + 1, 1, 2,
            font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
            number_format=FORMATS["INTEGER"]
        ))

    return ws, fmt_reqs
