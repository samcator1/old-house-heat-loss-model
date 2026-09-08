"""
Building Fabric and Ventilation Heat Loss Tab ('2_Building_Heat_Loss').
Calculates zone-by-zone fabric heat loss with separate transparent columns for
walls, windows, ground floor (corrected 10°C ground temp), roof, and infiltration.
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

def build_fabric_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 2_Building_Heat_Loss worksheet."""
    tab_name = "2_Building_Heat_Loss"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=30, cols=32)

    total_cols = 31
    total_rows = 15
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1: Title Banner
    grid[0][0] = "2_Building_Heat_Loss: Zone Fabric & Infiltration Heat Loss Analysis"

    # Row 2: Explainer
    grid[1][0] = "Transparent element-by-element peak heat loss at -4°C external design temperature and 10°C ground temperature."

    # Row 3: Group Headers
    grid[2][0] = "ZONE IDENTIFICATION & GEOMETRY"   # A-I (Cols 0-8)
    grid[2][9] = "EXTERNAL ENVELOPE AREAS"          # J-O (Cols 9-14)
    grid[2][15] = "ELEMENT U-VALUES (W/m²K)"       # P-S (Cols 15-18)
    grid[2][19] = "PEAK FABRIC HEAT LOSS (W)"       # T-X (Cols 19-23)
    grid[2][24] = "VENTILATION & INFILTRATION (W)"  # Y-AB (Cols 24-27)
    grid[2][28] = "TOTAL HEAT LOSS & INTENSITY"    # AC-AE (Cols 28-30)

    # Row 4: Column Headers
    headers = [
        "Zone Name",              # Col A (0)
        "Zone Type",              # Col B (1)
        "Length (m)",             # Col C (2)
        "Width (m)",              # Col D (3)
        "Footprint (m²)",         # Col E (4)
        "Floors",                 # Col F (5)
        "Floor Area (m²)",        # Col G (6)
        "Wall Ht (m)",            # Col H (7)
        "Volume (m³)",            # Col I (8)
        "Perimeter (m)",          # Col J (9)
        "% Ext Wall",             # Col K (10)
        "Gross Wall (m²)",        # Col L (11)
        "% Glazing",              # Col M (12)
        "Window Area (m²)",       # Col N (13)
        "Net Wall (m²)",          # Col O (14)
        "U Wall",                 # Col P (15)
        "U Window",               # Col Q (16)
        "U Ground",               # Col R (17)
        "U Roof",                 # Col S (18)
        "Wall Loss (W)",          # Col T (19)
        "Window Loss (W)",        # Col U (20)
        "Ground Loss (W)",        # Col V (21)
        "Roof Loss (W)",          # Col W (22)
        "Fabric Loss (W)",        # Col X (23)
        "Peak ACH",               # Col Y (24)
        "Peak Vent (W)",          # Col Z (25)
        "Avg ACH",                # Col AA (26)
        "Avg Vent (W)",           # Col AB (27)
        "Peak Heat Loss (W)",     # Col AC (28)
        "Peak Loss (kW)",         # Col AD (29)
        "Intensity (W/m²)"        # Col AE (30)
    ]
    for c_idx, h in enumerate(headers):
        grid[3][c_idx] = h

    # Populate Zone Data (Rows 5 to 10, indices 4 to 9)
    for z_idx, z in enumerate(DEFAULT_ZONES):
        r = z_idx + 5  # 1-indexed row number in spreadsheet (5 to 10)
        row_arr = [
            z["name"],
            z["zone_type"],
            str(z["length"]),
            str(z["width"]),
            f"=C{r}*D{r}",                                     # Footprint (Col E)
            str(z["floors"]),
            f"=E{r}*F{r}",                                     # Total Floor Area (Col G)
            str(round(z["h1"] + z["h2"], 2)),                 # Wall Height (Col H)
            f"=E{r}*H{r}",                                     # Volume (Col I)
            f"=(C{r}+D{r})*2",                                 # Perimeter (Col J)
            str(z["pct_ext_wall"]),                            # % Ext Wall (Col K)
            f"=J{r}*H{r}*K{r}",                                # Gross Wall Area (Col L)
            str(z["pct_glazing"]),                             # % Glazing (Col M)
            f"=L{r}*M{r}",                                     # Window Area (Col N)
            f"=L{r}-N{r}",                                     # Net Wall Area (Col O)
            str(z["u_wall"]),                                  # U Wall (Col P)
            str(z["u_window"]),                                # U Window (Col Q)
            str(z["u_ground"]),                                # U Ground (Col R)
            str(z["u_roof"]),                                  # U Roof (Col S)
            f"=O{r}*P{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$5)",    # Wall Loss W (Col T)
            f"=N{r}*Q{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$5)",    # Window Loss W (Col U)
            f"=E{r}*R{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$6)",    # Ground Loss W (Col V: corrected with Ground Temp C6!)
            f"=E{r}*S{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$5)",    # Roof Loss W (Col W)
            f"=SUM(T{r}:W{r})",                                # Total Fabric Loss W (Col X)
            f"=IF(B{r}=\"Old House\", '1_Inputs'!$C$12, '1_Inputs'!$C$14)", # Peak ACH (Col Y)
            f"='1_Inputs'!$C$11*Y{r}*I{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$5)", # Peak Vent Loss W (Col Z)
            f"=IF(B{r}=\"Old House\", '1_Inputs'!$C$13, '1_Inputs'!$C$15)", # Avg ACH (Col AA)
            f"='1_Inputs'!$C$11*AA{r}*I{r}*('1_Inputs'!$C$4-'1_Inputs'!$C$5)", # Avg Vent Loss W (Col AB)
            f"=X{r}+Z{r}",                                     # Peak Heat Loss W (Col AC)
            f"=AC{r}/1000",                                    # Peak Heat Loss kW (Col AD)
            f"=AC{r}/G{r}"                                     # Intensity W/m² (Col AE)
        ]
        grid[z_idx + 4] = row_arr

    # Row 11: Total Row (index 10)
    grid[10] = [
        "Total / Weighted Building",                           # Col A
        "Whole House",                                         # Col B
        "-",                                                   # Col C
        "-",                                                   # Col D
        "=SUM(E5:E10)",                                        # Total Footprint (Col E)
        "-",                                                   # Col F
        "=SUM(G5:G10)",                                        # Total Floor Area (Col G)
        "-",                                                   # Col H
        "=SUM(I5:I10)",                                        # Total Volume (Col I)
        "-",                                                   # Col J
        "-",                                                   # Col K
        "=SUM(L5:L10)",                                        # Total Gross Wall (Col L)
        "-",                                                   # Col M
        "=SUM(N5:N10)",                                        # Total Window Area (Col N)
        "=SUM(O5:O10)",                                        # Total Net Wall Area (Col O)
        "=SUMPRODUCT(O5:O10, P5:P10)/SUM(O5:O10)",             # Area-weighted U Wall (Col P)
        "=SUMPRODUCT(N5:N10, Q5:Q10)/SUM(N5:N10)",             # Area-weighted U Window (Col Q)
        "=SUMPRODUCT(E5:E10, R5:R10)/SUM(E5:E10)",             # Area-weighted U Ground (Col R)
        "=SUMPRODUCT(E5:E10, S5:S10)/SUM(E5:E10)",             # Area-weighted U Roof (Col S)
        "=SUM(T5:T10)",                                        # Total Wall Loss W (Col T)
        "=SUM(U5:U10)",                                        # Total Window Loss W (Col U)
        "=SUM(V5:V10)",                                        # Total Ground Loss W (Col V)
        "=SUM(W5:W10)",                                        # Total Roof Loss W (Col W)
        "=SUM(X5:X10)",                                        # Total Fabric Loss W (Col X)
        "=SUMPRODUCT(Y5:Y10, I5:I10)/SUM(I5:I10)",             # Volume-weighted Peak ACH (Col Y)
        "=SUM(Z5:Z10)",                                        # Total Peak Vent Loss W (Col Z)
        "=SUMPRODUCT(AA5:AA10, I5:I10)/SUM(I5:I10)",           # Volume-weighted Avg ACH (Col AA)
        "=SUM(AB5:AB10)",                                      # Total Avg Vent Loss W (Col AB)
        "=SUM(AC5:AC10)",                                      # Total Peak Heat Loss W (Col AC)
        "=AC11/1000",                                          # Total Peak Heat Loss kW (Col AD)
        "=AC11/G11"                                            # Whole Building Intensity W/m² (Col AE)
    ]

    # Write values into worksheet
    ws.update(values=grid, range_name=f"A1:AE{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze panes (rows 1-4 frozen)
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Column widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 170))  # Zone Name
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 95))   # Type
    for col_i in range(2, total_cols):
        fmt_reqs.append(create_set_column_width_request(ws.id, col_i, col_i + 1, 105))

    # Banner (Row 1 & 2)
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

    # Group Headers (Row 3)
    group_spans = [
        (0, 9, THEME["SECONDARY_HEADER_BG"]),
        (9, 15, THEME["SECTION_HEADER_BG"]),
        (15, 19, THEME["SECONDARY_HEADER_BG"]),
        (19, 24, {"red": 0.15, "green": 0.35, "blue": 0.45}),
        (24, 28, {"red": 0.12, "green": 0.30, "blue": 0.40}),
        (28, 31, THEME["PRIMARY_HEADER_BG"])
    ]
    for start_c, end_c, bg in group_spans:
        fmt_reqs.append(create_merge_cells_request(ws.id, 2, 3, start_c, end_c))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, 2, 3, start_c, end_c,
            bg_color=bg,
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=9,
            align="CENTER"
        ))

    # Column Headers (Row 4)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 3, 4, 0, total_cols,
        bg_color=THEME["CARD_HEADER_BG"],
        font_color=THEME["HEADER_TEXT"],
        bold=True,
        font_size=9,
        align="CENTER",
        wrap_strategy="WRAP"
    ))

    # Data Rows (Rows 5 to 10)
    for r_i in range(4, 10):
        bg = THEME["ZEBRA_BG"] if r_i % 2 == 1 else {"red": 1.0, "green": 1.0, "blue": 1.0}
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, total_cols,
            bg_color=bg,
            font_color=THEME["DARK_TEXT"],
            font_size=9,
            align="RIGHT"
        ))
        # Left align Zone Name and Type
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, 2,
            font_color=THEME["DARK_TEXT"],
            bold=(r_i == 4),
            align="LEFT"
        ))

    # Number formats for data rows (Rows 5 to 10)
    # Footprint, Floor Area, Volume (Cols E, G, I -> indices 4, 6, 8)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 4, 5, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 6, 7, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 8, 9, number_format=FORMATS["INTEGER"]))
    # % Ext Wall, % Glazing (Cols K, M -> indices 10, 12)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 10, 11, number_format=FORMATS["PERCENT_INT"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 12, 13, number_format=FORMATS["PERCENT_INT"]))
    # Areas (Cols L, N, O -> indices 11, 13, 14)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 11, 12, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 13, 15, number_format=FORMATS["INTEGER"]))
    # U-values (Cols P, Q, R, S -> indices 15-18)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 15, 19, number_format=FORMATS["DECIMAL_2"]))
    # Fabric Losses (Cols T, U, V, W, X -> indices 19-23)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 19, 24, number_format=FORMATS["INTEGER"]))
    # ACH (Cols Y, AA -> indices 24, 26)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 24, 25, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 26, 27, number_format=FORMATS["DECIMAL_2"]))
    # Vent Losses (Cols Z, AB -> indices 25, 27)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 25, 26, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 27, 28, number_format=FORMATS["INTEGER"]))
    # Peak Loss W, kW, W/m² (Cols AC, AD, AE -> indices 28-30)
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 28, 29, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 29, 30, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 10, 30, 31, number_format=FORMATS["DECIMAL_1"]))

    # Total Row (Row 11, index 10)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 10, 11, 0, total_cols,
        bg_color=THEME["TOTAL_BG"],
        font_color=THEME["DARK_TEXT"],
        bold=True,
        font_size=10,
        align="RIGHT"
    ))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 10, 11, 0, 2,
        align="LEFT"
    ))
    # Number formats for total row
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 4, 5, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 6, 7, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 8, 9, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 11, 12, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 13, 15, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 15, 19, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 19, 24, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 24, 25, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 25, 26, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 26, 27, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 27, 28, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 28, 29, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 29, 30, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 10, 11, 30, 31, number_format=FORMATS["DECIMAL_1"]))

    # Accent on Total Peak kW (Col AD, row 11)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 10, 11, 29, 30,
        bg_color=THEME["ACCENT_BG"],
        font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
        bold=True,
        font_size=11,
        align="RIGHT"
    ))

    return ws, fmt_reqs
