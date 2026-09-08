"""
Executive KPI Dashboard Tab ('0_Executive_Dashboard').
High-level executive overview with KPI metric cards, comparative system table,
zone contribution breakdown, and renewable energy self-sufficiency balance.
"""

from typing import Tuple, List, Dict, Any
import gspread
from ..config import THEME, FORMATS, DEFAULT_ZONES
from ..formatting import (
    create_repeat_cell_request,
    create_set_column_width_request,
    create_freeze_pane_request,
    create_merge_cells_request
)

def build_dashboard_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 0_Executive_Dashboard worksheet."""
    tab_name = "0_Executive_Dashboard"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=55, cols=8)

    total_cols = 7
    total_rows = 50
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Main Banner
    grid[0][0] = "0_Executive_Dashboard: Old House Heat Loss & Heating Decarbonization Model"
    grid[1][0] = "Final Build Specification: Whole-House Heat Loss, Hot Water, Pool, and Low-Carbon Energy Economics."

    # Row 3-6: KPI Cards
    # Card 1: Peak Heat Loss (Cols A-B, Rows 4-6)
    grid[2][0] = "PEAK HEAT LOSS (-4°C)"
    grid[3][0] = "='2_Room_Heat_Loss'!$AD$28/1000"
    grid[4][0] = "=TEXT('2_Room_Heat_Loss'!$AE$28, \"0.0\") & \" W/m² whole-house average\""

    # Card 2: Total Delivered Heat (Cols C-D, Rows 4-6)
    grid[2][2] = "ANNUAL DELIVERED HEAT"
    grid[3][2] = "='4_Heating_and_Renewables'!$B$8"
    grid[4][2] = "Space + DHW + Swimming Pool"

    # Card 3: Oil Boiler Operating Cost (Cols E-G, Rows 4-6)
    grid[2][4] = "OIL BOILER ANNUAL COST"
    grid[3][4] = "='4_Heating_and_Renewables'!$E$46"
    grid[4][4] = "Oil fuel + domestic electricity baseline"

    # Row 7-9: Second Row of KPI Cards
    # Card 4: GSHP + Solar + Battery Cost (Cols A-B, Rows 8-10)
    grid[6][0] = "GSHP + SOLAR + BATTERY COST"
    grid[7][0] = "='4_Heating_and_Renewables'!$E$51"
    grid[8][0] = "Smart time-of-use tariff & export credit"

    # Card 5: Net Annual Savings (Cols C-D, Rows 8-10)
    grid[6][2] = "ANNUAL OPERATING SAVING"
    grid[7][2] = "='4_Heating_and_Renewables'!$F$51"
    grid[8][2] = "=TEXT('4_Heating_and_Renewables'!$F$51/'4_Heating_and_Renewables'!$E$46, \"0.0%\") & \" cost reduction vs Oil\""

    # Card 6: Carbon Reduction (Cols E-G, Rows 8-10)
    grid[6][4] = "ANNUAL CO2 REDUCTION"
    grid[7][4] = "=('4_Heating_and_Renewables'!$G$46-'4_Heating_and_Renewables'!$G$51)/1000"
    grid[8][4] = "=TEXT(('4_Heating_and_Renewables'!$G$46-'4_Heating_and_Renewables'!$G$51)/'4_Heating_and_Renewables'!$G$46, \"0.0%\") & \" emissions reduction\""

    # Row 12: Section 1 - Technology Options Matrix
    grid[11][0] = "EXECUTIVE COMPARISON: HEATING & POWER OPTIONS"
    headers_tech = [
        "Heating & Power Option",
        "Turnkey Cap-Ex (£)",
        "Annual Fuel / Power",
        "Net Energy Cost (£/yr)",
        "Annual Savings (£/yr)",
        "10-Yr Total Cost (£)",
        "Emissions (t CO2/yr)"
    ]
    for c_i, h in enumerate(headers_tech):
        grid[12][c_i] = h

    tech_rows = [
        ("Oil Boiler Baseline (No Solar)", "='4_Heating_and_Renewables'!B46", "=TEXT('4_Heating_and_Renewables'!D29, \"#,##0\") & \" L oil\"", "='4_Heating_and_Renewables'!E46", "='4_Heating_and_Renewables'!F46", "='4_Heating_and_Renewables'!D55", "='4_Heating_and_Renewables'!G46/1000"),
        ("Oil Boiler + Solar PV + Battery", "='4_Heating_and_Renewables'!B47", "=TEXT('4_Heating_and_Renewables'!D29, \"#,##0\") & \" L + PV\"", "='4_Heating_and_Renewables'!E47", "='4_Heating_and_Renewables'!F47", "='4_Heating_and_Renewables'!D56", "='4_Heating_and_Renewables'!G47/1000"),
        ("ASHP (Flat Tariff, No Solar)", "='4_Heating_and_Renewables'!B48", "=TEXT('4_Heating_and_Renewables'!C35, \"#,##0\") & \" kWh el\"", "='4_Heating_and_Renewables'!E48", "='4_Heating_and_Renewables'!F48", "='4_Heating_and_Renewables'!D57", "='4_Heating_and_Renewables'!G48/1000"),
        ("ASHP + Solar PV + Battery (Smart Tariff)", "='4_Heating_and_Renewables'!B49", "=TEXT('4_Heating_and_Renewables'!C39, \"#,##0\") & \" kWh net\"", "='4_Heating_and_Renewables'!E49", "='4_Heating_and_Renewables'!F49", "='4_Heating_and_Renewables'!D58", "='4_Heating_and_Renewables'!G49/1000"),
        ("GSHP (Flat Tariff, No Solar)", "='4_Heating_and_Renewables'!B50", "=TEXT('4_Heating_and_Renewables'!B35, \"#,##0\") & \" kWh el\"", "='4_Heating_and_Renewables'!E50", "='4_Heating_and_Renewables'!F50", "='4_Heating_and_Renewables'!D59", "='4_Heating_and_Renewables'!G50/1000"),
        ("GSHP + Solar PV + Battery (Smart Tariff)", "='4_Heating_and_Renewables'!B51", "=TEXT('4_Heating_and_Renewables'!B39, \"#,##0\") & \" kWh net\"", "='4_Heating_and_Renewables'!E51", "='4_Heating_and_Renewables'!F51", "='4_Heating_and_Renewables'!D60", "='4_Heating_and_Renewables'!G51/1000")
    ]

    for r_offset, tr in enumerate(tech_rows):
        r_num = 13 + r_offset
        for c_i in range(7):
            grid[r_num][c_i] = tr[c_i]

    # Row 22: Section 2 - Bottom-Up Architectural Wing & Room Breakdown
    grid[21][0] = "BOTTOM-UP ARCHITECTURAL WING & ROOM HEAT LOSS BREAKDOWN (23 ROOMS)"
    headers_zone = [
        "Zone / Wing",
        "Room Count & Type",
        "Floor Area (m²)",
        "Area Share (%)",
        "Peak Heat Loss (kW)",
        "Loss Share (%)",
        "Intensity (W/m²)"
    ]
    for c_i, h in enumerate(headers_zone):
        grid[22][c_i] = h

    zone_types = {
        "Georgian end": "Solid Brick",
        "Thatched gable ended": "Solid Stone",
        "New build": "Insulated Cavity",
        "Orangery": "Glazed Lantern",
        "Lean to West": "Uninsulated Solid",
        "Lean to North": "Plant Room"
    }

    for z_idx, z in enumerate(DEFAULT_ZONES):
        z_name = z["name"]
        z_desc = zone_types.get(z_name, z.get("zone_type", "Standard"))
        dash_r = 23 + z_idx
        grid[dash_r][0] = z_name
        grid[dash_r][1] = f'=COUNTIF(\'2_Room_Heat_Loss\'!$D$5:$D$27, A{dash_r+1}) & " rms (" & "{z_desc})"'
        grid[dash_r][2] = f"=SUMIF('2_Room_Heat_Loss'!$D$5:$D$27, A{dash_r+1}, '2_Room_Heat_Loss'!$H$5:$H$27)"
        grid[dash_r][3] = f"=C{dash_r+1}/$C$30"
        grid[dash_r][4] = f"=SUMIF('2_Room_Heat_Loss'!$D$5:$D$27, A{dash_r+1}, '2_Room_Heat_Loss'!$AD$5:$AD$27)/1000"
        grid[dash_r][5] = f"=E{dash_r+1}/$E$30"
        grid[dash_r][6] = f"=(E{dash_r+1}*1000)/C{dash_r+1}"

    # Total Zone Row (Row 30)
    grid[29][0] = "Total Whole Building"
    grid[29][1] = '="All " & COUNT(\'2_Room_Heat_Loss\'!$H$5:$H$27) & " Rooms"'
    grid[29][2] = "='2_Room_Heat_Loss'!$H$28"
    grid[29][3] = "=SUM(D24:D29)"
    grid[29][4] = "='2_Room_Heat_Loss'!$AD$28/1000"
    grid[29][5] = "=SUM(F24:F29)"
    grid[29][6] = "='2_Room_Heat_Loss'!$AE$28"

    # Row 33: Section 3 - Solar PV & Battery Self-Sufficiency Summary
    grid[32][0] = "SOLAR PV & BATTERY DISPATCH SUMMARY (GSHP SCENARIO)"
    headers_solar = [
        "Renewable Metric",
        "Annual Value (kWh)",
        "Daily Avg (kWh)",
        "Share of Generation (%)",
        "Tariff Slot Rate (£/kWh)",
        "Annual Cost / Credit (£)",
        "Status / Strategy"
    ]
    for c_i, h in enumerate(headers_solar):
        grid[33][c_i] = h

    solar_summary = [
        ("Solar PV Total Generation", "='4_Heating_and_Renewables'!B36", "=B35/365", "100.0%", "-", "-", "Clean zero-carbon on-site generation"),
        ("Solar Used On-Site (Heat Pump + Household)", "='4_Heating_and_Renewables'!B37", "=B36/365", "=B36/$B$35", "-", "-", "Direct solar consumption boosted by battery"),
        ("Surplus Solar Exported to Grid", "='4_Heating_and_Renewables'!B38", "=B37/365", "=B37/$B$35", "='1_Inputs'!$C$57", "='4_Heating_and_Renewables'!D51", "Export revenue earned via SEG"),
        ("Net Grid Electricity Purchased", "='4_Heating_and_Renewables'!B39", "=B38/365", "-", "Weighted Smart Rate", "='4_Heating_and_Renewables'!C51", "Shifted to 65% cheap overnight tariff via battery"),
        ("Site Self-Sufficiency Ratio", "=B36/'4_Heating_and_Renewables'!B35", "-", "-", "-", "-", "Percentage of total electricity generated on-site")
    ]

    for s_idx, ss_row in enumerate(solar_summary):
        dash_r = 34 + s_idx
        for c_i in range(7):
            grid[dash_r][c_i] = ss_row[c_i]

    # Write to sheet
    ws.update(values=grid, range_name=f"A1:G{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze banner
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=2, frozen_cols=0))

    # Widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 290))  # Col A
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 140))  # Col B
    fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 140))  # Col C
    fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 150))  # Col D
    fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 150))  # Col E
    fmt_reqs.append(create_set_column_width_request(ws.id, 5, 6, 150))  # Col F
    fmt_reqs.append(create_set_column_width_request(ws.id, 6, 7, 240))  # Col G

    # Banner
    fmt_reqs.append(create_merge_cells_request(ws.id, 0, 1, 0, total_cols))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 0, 1, 0, total_cols,
        bg_color=THEME["PRIMARY_HEADER_BG"],
        font_color=THEME["HEADER_TEXT"],
        bold=True,
        font_size=13,
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

    # KPI Cards Formatting: (start_row, end_row, start_col, end_col, title_bg, format_str)
    kpi_card_spans = [
        (2, 5, 0, 2, THEME["CARD_HEADER_BG"], FORMATS["KW"]),          # Card 1: Peak Heat Loss
        (2, 5, 2, 4, THEME["SECTION_HEADER_BG"], FORMATS["KWH"]),       # Card 2: Annual Heat Delivered
        (2, 5, 4, 7, {"red": 0.65, "green": 0.20, "blue": 0.20}, FORMATS["CURRENCY_GBP"]), # Card 3: Oil Boiler
        (6, 9, 0, 2, THEME["CARD_HEADER_BG"], FORMATS["CURRENCY_GBP"]), # Card 4: GSHP + Solar
        (6, 9, 2, 4, {"red": 0.08, "green": 0.45, "blue": 0.25}, FORMATS["CURRENCY_GBP"]), # Card 5: Savings
        (6, 9, 4, 7, {"red": 0.12, "green": 0.40, "blue": 0.35}, FORMATS["TONNES_CO2"]),   # Card 6: CO2
    ]

    for sr, er, sc, ec, h_bg, fmt in kpi_card_spans:
        # Card Header
        fmt_reqs.append(create_merge_cells_request(ws.id, sr, sr + 1, sc, ec))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr, sr + 1, sc, ec,
            bg_color=h_bg,
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=9,
            align="CENTER"
        ))
        # Card Value
        fmt_reqs.append(create_merge_cells_request(ws.id, sr + 1, sr + 2, sc, ec))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr + 1, sr + 2, sc, ec,
            bg_color=THEME["CARD_BG"],
            font_color=THEME["DARK_TEXT"],
            bold=True,
            font_size=15,
            align="CENTER",
            number_format=fmt
        ))
        # Card Subtitle
        fmt_reqs.append(create_merge_cells_request(ws.id, sr + 2, er, sc, ec))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr + 2, er, sc, ec,
            bg_color=THEME["CARD_BG"],
            font_color=THEME["MUTED_TEXT"],
            italic=True,
            font_size=9,
            align="CENTER"
        ))

    # Section Headers
    section_rows = [11, 21, 32]
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
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, sr + 1, sr + 2, 0, total_cols,
            bg_color=THEME["SECONDARY_HEADER_BG"],
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=9,
            align="CENTER"
        ))

    # Format Tech Matrix Rows (Rows 14 to 19)
    for r in range(13, 19):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 2, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="RIGHT", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 6, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 6, 7, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_1"]))

    # Highlight Best Option (Row 19: GSHP + Solar + Battery)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 18, 19, 0, total_cols, bg_color=THEME["CALC_BG"], bold=True))

    # Format Zone Rows (Rows 24 to 29)
    for r in range(23, 29):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 2, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 4, align="RIGHT", font_size=9, number_format=FORMATS["PERCENT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 4, 5, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_1"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="RIGHT", font_size=9, number_format=FORMATS["PERCENT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 6, 7, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_1"]))

    # Total Zone Row (Row 30)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 0, total_cols, bg_color=THEME["TOTAL_BG"], bold=True, font_size=10))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 2, 3, align="RIGHT", number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 3, 4, align="RIGHT", number_format=FORMATS["PERCENT"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 4, 5, align="RIGHT", number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 5, 6, align="RIGHT", number_format=FORMATS["PERCENT"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 29, 30, 6, 7, align="RIGHT", number_format=FORMATS["DECIMAL_1"]))

    # Format Solar Summary Rows (Rows 35 to 39)
    for r in range(34, 39):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 3, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 4, align="RIGHT", font_size=9, number_format=FORMATS["PERCENT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 4, 5, align="RIGHT", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 6, 7, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    return ws, fmt_reqs
