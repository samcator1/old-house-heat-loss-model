"""
Heating Systems, Renewables & Tariff Economics Tab ('4_Heating_and_Renewables').
Compares Ground Source Heat Pump (GSHP), Air Source Heat Pump (ASHP), and Oil Boiler,
with Solar PV, Battery storage dispatch, and Smart Time-of-Use tariffs.
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

def build_systems_tab(ss: gspread.Spreadsheet, room_total_row: int = 29, num_rooms: int = 24) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 4_Heating_and_Renewables worksheet."""
    tab_name = "4_Heating_and_Renewables"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=70, cols=10)

    total_cols = 7
    total_rows = 68
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Banner
    grid[0][0] = "4_Heating_and_Renewables: Heating Systems Comparison, Solar PV & Tariff Economics"
    grid[1][0] = "Side-by-side performance, capital costs, Solar PV + Battery dispatch, and smart time-of-use tariff economics."

    # Row 3: Section 1 Header
    grid[2][0] = "SECTION A: WHOLE-BUILDING ANNUAL DELIVERED HEAT SUMMARY"
    grid[3][0] = "Thermal End-Use Demand"
    grid[3][1] = "Delivered Heat (kWh/yr)"
    grid[3][2] = "Share (%)"
    grid[3][3] = "Calculation / Sourcing"
    grid[3][4] = "Engineering Rationale"

    grid[4] = [
        "Space heating annual demand",
        f"=(('2_Room_Heat_Loss'!$AH${room_total_row}/('2_Room_Heat_Loss'!$E${room_total_row}-'1_Inputs'!$C$5))*'1_Inputs'!$C$7*24/1000)*'1_Inputs'!$C$8",
        "=B5/$B$8",
        "(HLC × HDD × 24 / 1000) × f_usage",
        "Degree-day method using bottom-up room schedule heat loss coefficient"
    ]
    grid[5] = [
        "Domestic hot water (DHW) demand",
        "='3_DHW_and_Pool'!$B$19",
        "=B6/$B$8",
        "='3_DHW_and_Pool'!B19",
        "Water heating, standing vessel losses, secondary loop & Legionella"
    ]
    grid[6] = [
        "Swimming pool heating demand",
        "='3_DHW_and_Pool'!$B$43",
        "=B7/$B$8",
        "='3_DHW_and_Pool'!B43",
        "Seasonal maintenance surface losses over 90 days + initial fill"
    ]
    grid[7] = [
        "TOTAL ANNUAL DELIVERED HEAT",
        "=SUM(B5:B7)",
        "=SUM(C5:C7)",
        "=SUM(B5:B7)",
        "Total useful heat required by house, hot water, and pool"
    ]
    grid[8] = [
        "Domestic baseline electricity load",
        "='1_Inputs'!$C$61",
        "-",
        "Inputs C61",
        "Household non-heating electricity (appliances, cooking, lighting)"
    ]

    # Row 11: Section B Header
    grid[10][0] = "SECTION B: SYSTEM SIZING & TURNKEY CAPITAL INSTALLATION COSTS"
    grid[11][0] = "Heating & Renewable Technology"
    grid[11][1] = "Design Plant Sizing"
    grid[11][2] = "Unit"
    grid[11][3] = "Cost Rate"
    grid[11][4] = "Turnkey Cap-Ex (£)"
    grid[11][5] = "Scope of Supply & Installation"

    grid[12] = [
        "Ground Source Heat Pump (GSHP)",
        f"=(('2_Room_Heat_Loss'!$AH${room_total_row}/1000)*(1+'1_Inputs'!$C$16))",
        "kW heat",
        "='1_Inputs'!$C$64",
        "=B13*D13",
        "Ecoforest heat pump, ground borehole collector array & plantroom"
    ]
    grid[13] = [
        "Air Source Heat Pump (ASHP)",
        f"=(('2_Room_Heat_Loss'!$AH${room_total_row}/1000)*(1+'1_Inputs'!$C$16))",
        "kW heat",
        "='1_Inputs'!$C$65",
        "=B14*D14",
        "External outdoor heat pump units, indoor buffer & integration"
    ]
    grid[14] = [
        "Commercial Oil Boiler",
        f"=(('2_Room_Heat_Loss'!$AH${room_total_row}/1000)*(1+'1_Inputs'!$C$17))",
        "kW heat",
        "='1_Inputs'!$C$66",
        "=B15*D15",
        "High-output condensing oil boiler with intermittent boost margin"
    ]
    grid[15] = [
        "Rooftop Solar PV Array",
        "='1_Inputs'!$C$58",
        "kWp",
        "='1_Inputs'!$C$67",
        "=B16*D16",
        "37 kWp high-efficiency solar photovoltaic array & inverters"
    ]
    grid[16] = [
        "Home Battery Storage System",
        "='1_Inputs'!$C$60",
        "kWh",
        "='1_Inputs'!$C$68",
        "=B17*D17",
        "Modular lithium battery storage for off-peak charging & solar buffering"
    ]

    # Row 19: Section C Header
    grid[18][0] = "SECTION C: HEATING SYSTEM PERFORMANCE & ANNUAL FUEL CONSUMPTION"
    grid[19][0] = "Performance Metric"
    grid[19][1] = "GSHP (Ground Source)"
    grid[19][2] = "ASHP (Air Source)"
    grid[19][3] = "Oil Boiler Baseline"
    grid[19][4] = "Unit"
    grid[19][5] = "Notes / Source"

    grid[20] = [
        "Space heating efficiency / SCOP (45°C flow)",
        "='1_Inputs'!$C$44",
        "='1_Inputs'!$C$47",
        "='1_Inputs'!$C$42",
        "SCOP / %",
        "Medium-temperature emitter design flow"
    ]
    grid[21] = [
        "DHW generation efficiency / SCOP (55°C flow)",
        "='1_Inputs'!$C$45",
        "='1_Inputs'!$C$48",
        "='1_Inputs'!$C$42",
        "SCOP / %",
        "Cylinder heating to 55°C"
    ]
    grid[22] = [
        "Pool heating efficiency / SCOP (35°C flow)",
        "='1_Inputs'!$C$46",
        "='1_Inputs'!$C$49",
        "='1_Inputs'!$C$42",
        "SCOP / %",
        "Low-temperature heat exchanger flow"
    ]
    grid[23] = [
        "Weighted Annual System SCOP / Efficiency",
        "=$B$8/((B5/B21)+(B6/B22)+(B7/B23))",
        "=$B$8/((B5/C21)+(B6/C22)+(B7/C23))",
        "='1_Inputs'!$C$42",
        "SCOP / %",
        "Annual delivered heat divided by energy input"
    ]
    grid[24] = [
        "Space heating input energy",
        "=B5/B21",
        "=B5/C21",
        "=B5/D21",
        "kWh/year",
        "Purchased electricity or oil fuel input"
    ]
    grid[25] = [
        "DHW generation input energy",
        "=B6/B22",
        "=B6/C22",
        "=B6/D22",
        "kWh/year",
        "Purchased electricity or oil fuel input"
    ]
    grid[26] = [
        "Pool heating input energy",
        "=B7/B23",
        "=B7/C23",
        "=B7/D23",
        "kWh/year",
        "Purchased electricity or oil fuel input"
    ]
    grid[27] = [
        "TOTAL ANNUAL HEATING INPUT ENERGY",
        "=SUM(B25:B27)",
        "=SUM(C25:C27)",
        "=SUM(D25:D27)",
        "kWh/year",
        "Total plant input requirement"
    ]
    grid[28] = [
        "Annual heating oil volume required",
        "-",
        "-",
        "=D28/'1_Inputs'!$C$43",
        "Litres/year",
        "At 10 kWh/Litre kerosene calorific value"
    ]

    # Row 31: Section D Header
    grid[30][0] = "SECTION D: SOLAR PV & BATTERY DISPATCH WITH ELECTRICITY LOAD"
    grid[31][0] = "Electricity Balance Parameter"
    grid[31][1] = "GSHP System"
    grid[31][2] = "ASHP System"
    grid[31][3] = "Oil Boiler System"
    grid[31][4] = "Unit"
    grid[31][5] = "Basis & Operating Strategy"

    grid[32] = [
        "Annual heating electricity consumption",
        "=B28",
        "=C28",
        "0",
        "kWh/year",
        "From Section C heating input"
    ]
    grid[33] = [
        "Domestic baseline electricity consumption",
        "=$B$9",
        "=$B$9",
        "=$B$9",
        "kWh/year",
        "Household non-heating appliances & lighting"
    ]
    grid[34] = [
        "Total annual household electricity load",
        "=B33+B34",
        "=C33+C34",
        "=D33+D34",
        "kWh/year",
        "Combined site electricity demand"
    ]
    grid[35] = [
        "Annual solar PV generation",
        "='1_Inputs'!$C$58*'1_Inputs'!$C$59",
        "='1_Inputs'!$C$58*'1_Inputs'!$C$59",
        "='1_Inputs'!$C$58*'1_Inputs'!$C$59",
        "kWh/year",
        "37 kWp × 900 kWh/kWp specific yield"
    ]
    grid[36] = [
        "Solar electricity self-consumed (with Battery)",
        "=MIN(B35, B36*0.70)",
        "=MIN(C35, C36*0.70)",
        "=MIN(D35, D36*0.50)",
        "kWh/year",
        "70% self-use for heat pumps + battery, 50% for domestic only"
    ]
    grid[37] = [
        "Surplus solar electricity exported to grid",
        "=B36-B37",
        "=C36-C37",
        "=D36-D37",
        "kWh/year",
        "Exported at Smart Export Guarantee (SEG) rate"
    ]
    grid[38] = [
        "Net grid electricity imported",
        "=B35-B37",
        "=C35-C37",
        "=D35-D37",
        "kWh/year",
        "Electricity purchased from grid"
    ]
    grid[39] = [
        "Smart Tariff - Off-Peak import (8 hrs overnight)",
        "=B39*0.65",
        "=C39*0.65",
        "=D39*0.50",
        "kWh/year",
        "65% of import shifted to cheap 15p overnight slot via battery"
    ]
    grid[40] = [
        "Smart Tariff - Day / Standard import (13 hrs)",
        "=B39*0.30",
        "=C39*0.30",
        "=D39*0.40",
        "kWh/year",
        "Daytime background consumption at 27p"
    ]
    grid[41] = [
        "Smart Tariff - Peak import (3 hrs, 4pm-7pm)",
        "=B39*0.05",
        "=C39*0.05",
        "=D39*0.10",
        "kWh/year",
        "Minimized to 5% as battery discharges during 40p peak"
    ]

    # Row 44: Section E Header
    grid[43][0] = "SECTION E: COMPREHENSIVE ECONOMIC & CARBON COMPARISON"
    grid[44][0] = "Heating & Power Scenario"
    grid[44][1] = "Turnkey Cap-Ex (£)"
    grid[44][2] = "Gross Energy Cost (£/yr)"
    grid[44][3] = "Solar Export Credit (£/yr)"
    grid[44][4] = "Net Energy Cost (£/yr)"
    grid[44][5] = "Annual Savings (£/yr)"
    grid[44][6] = "Carbon Emissions (kg CO2/yr)"

    scenarios = [
        ("1. Oil Boiler Baseline (No Solar/Battery)", "=E15", "=D29*'1_Inputs'!$C$52+D34*'1_Inputs'!$C$53", "0", "=C46-D46", "0", "=D28*'1_Inputs'!$C$70+D34*'1_Inputs'!$C$69"),
        ("2. Oil Boiler + Solar PV + Battery (Smart Tariff)", "=E15+E16+E17", "=D29*'1_Inputs'!$C$52+(D40*'1_Inputs'!$C$54+D41*'1_Inputs'!$C$55+D42*'1_Inputs'!$C$56)", "=D38*'1_Inputs'!$C$57", "=C47-D47", "=$E$46-E47", "=D28*'1_Inputs'!$C$70+D39*'1_Inputs'!$C$69"),
        ("3. ASHP Baseline (Flat Tariff, No Solar)", "=E14", "=(C35*'1_Inputs'!$C$53)", "0", "=C48-D48", "=$E$46-E48", "=C35*'1_Inputs'!$C$69"),
        ("4. ASHP + Solar PV + Battery (Smart Tariff)", "=E14+E16+E17", "=(C40*'1_Inputs'!$C$54+C41*'1_Inputs'!$C$55+C42*'1_Inputs'!$C$56)", "=C38*'1_Inputs'!$C$57", "=C49-D49", "=$E$46-E49", "=C39*'1_Inputs'!$C$69"),
        ("5. GSHP Baseline (Flat Tariff, No Solar)", "=E13", "=(B35*'1_Inputs'!$C$53)", "0", "=C50-D50", "=$E$46-E50", "=B35*'1_Inputs'!$C$69"),
        ("6. GSHP + Solar PV + Battery (Smart Tariff)", "=E13+E16+E17", "=(B40*'1_Inputs'!$C$54+B41*'1_Inputs'!$C$55+B42*'1_Inputs'!$C$56)", "=B38*'1_Inputs'!$C$57", "=C51-D51", "=$E$46-E51", "=B39*'1_Inputs'!$C$69")
    ]

    for s_idx, sc in enumerate(scenarios):
        r_i = 45 + s_idx
        grid[r_i][0] = sc[0]
        grid[r_i][1] = sc[1]
        grid[r_i][2] = sc[2]
        grid[r_i][3] = sc[3]
        grid[r_i][4] = sc[4]
        grid[r_i][5] = sc[5]
        grid[r_i][6] = sc[6]

    # Row 53: 10-Year Cumulative Total Cost of Ownership (TCO)
    grid[52][0] = "SECTION F: 10-YEAR TOTAL COST OF OWNERSHIP (TCO)"
    grid[53][0] = "Heating Option"
    grid[53][1] = "Turnkey Cap-Ex (£)"
    grid[53][2] = "10-Yr Operating (£)"
    grid[53][3] = "10-Yr Total Cost (£)"
    grid[53][4] = "10-Yr Net Savings vs Oil (£)"
    grid[53][5] = "Simple Payback (Years)"

    tco_options = [
        ("Oil Boiler Baseline", "=B46", "=E46*10", "=B55+C55", "0", "-"),
        ("Oil Boiler + Solar + Battery", "=B47", "=E47*10", "=B56+C56", "=D$55-D56", "=(B47-B$46)/MAX(1, F47)"),
        ("ASHP (Flat Tariff)", "=B48", "=E48*10", "=B57+C57", "=D$55-D57", "=(B48-B$46)/MAX(1, F48)"),
        ("ASHP + Solar + Battery", "=B49", "=E49*10", "=B58+C58", "=D$55-D58", "=(B49-B$46)/MAX(1, F49)"),
        ("GSHP (Flat Tariff)", "=B50", "=E50*10", "=B59+C59", "=D$55-D59", "=(B50-B$46)/MAX(1, F50)"),
        ("GSHP + Solar + Battery", "=B51", "=E51*10", "=B60+C60", "=D$55-D60", "=(B51-B$46)/MAX(1, F51)")
    ]

    for t_idx, to in enumerate(tco_options):
        r_i = 54 + t_idx
        grid[r_i][0] = to[0]
        grid[r_i][1] = to[1]
        grid[r_i][2] = to[2]
        grid[r_i][3] = to[3]
        grid[r_i][4] = to[4]
        grid[r_i][5] = to[5]

    # Write to sheet
    ws.update(values=grid, range_name=f"A1:G{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze top
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 310))  # Metric / Description
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 145))  # Col B
    fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 145))  # Col C
    fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 145))  # Col D
    fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 145))  # Col E
    fmt_reqs.append(create_set_column_width_request(ws.id, 5, 6, 145))  # Col F
    fmt_reqs.append(create_set_column_width_request(ws.id, 6, 7, 180))  # Col G

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
    section_rows = [2, 10, 18, 30, 43, 52]
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

    # Data formatting Section A (rows 5 to 9)
    for r in range(4, 9):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 2, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="RIGHT", font_size=9, number_format=FORMATS["PERCENT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 5, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    # Accent on Total Heat Delivered (Row 8)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 7, 8, 0, total_cols, bg_color=THEME["ACCENT_BG"], bold=True))

    # Data formatting Section B (rows 13 to 17)
    for r in range(12, 17):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 2, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_1"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 2, 3, align="CENTER", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 3, 5, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    # Data formatting Section C (rows 21 to 29)
    for r in range(20, 29):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        if r in [20, 21, 22, 23]:
            fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 4, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_2"]))
        else:
            fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 4, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 4, 5, align="CENTER", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    # Data formatting Section D (rows 33 to 42)
    for r in range(32, 42):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 4, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 4, 5, align="CENTER", font_size=9, font_color=THEME["MUTED_TEXT"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="LEFT", font_size=9, font_color=THEME["MUTED_TEXT"]))

    # Data formatting Section E (rows 46 to 51)
    for r in range(45, 51):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 6, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 6, 7, align="RIGHT", font_size=10, number_format=FORMATS["INTEGER"]))

    # Highlight GSHP + Solar + Battery (Row 51)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 50, 51, 0, total_cols,
        bg_color=THEME["CALC_BG"],
        bold=True
    ))

    # Data formatting Section F (rows 55 to 60)
    for r in range(54, 60):
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 0, 1, bg_color=THEME["ZEBRA_BG"], font_size=9))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 1, 5, align="RIGHT", font_size=10, number_format=FORMATS["CURRENCY_GBP"]))
        fmt_reqs.append(create_repeat_cell_request(ws.id, r, r + 1, 5, 6, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_1"]))

    return ws, fmt_reqs
