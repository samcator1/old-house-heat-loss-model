"""
Room-by-Room Heat Loss Assessment & Emitter Sizing Tab ('5_Room_Heat_Loss').
MCS / CIBSE BS EN 12831 compliant room-by-room schedule for all 23 active rooms:
- 10 Ground Floor rooms (including double-height Hall)
- 13 First Floor rooms (bathrooms, bedrooms, corridors, laundry)
- Unoccupied attic boundary calculations
Provides room-by-room Watts, W/m², low-flow 45°C radiator sizing, and builder notes.
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

# 23 Rooms Specification (10 Ground Floor, 13 First Floor)
DEFAULT_ROOMS = [
    # Ground Floor (10 rooms)
    {"code": "GF-01", "name": "Entrance & Stair Hall (GF to FF)", "floor": "Ground Floor", "zone": "Georgian end", "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 5.60, "ext_wall": 10.75, "u_wall": 1.40, "win_area": 6.5, "u_win": 4.80, "fl_area": 28.1, "u_fl": 0.80, "roof_area": 0.0, "u_roof": 0.18, "ach": 1.8, "notes": "Double-height open stair volume connecting GF to FF. Check draught seal on front door."},
    {"code": "GF-02", "name": "Drawing Room / Main Living", "floor": "Ground Floor", "zone": "Georgian end", "temp": 21.0, "len": 8.50, "wid": 6.25, "ht": 2.80, "ext_wall": 14.75, "u_wall": 1.40, "win_area": 9.2, "u_win": 4.80, "fl_area": 53.1, "u_fl": 0.80, "roof_area": 0.0, "u_roof": 0.18, "ach": 1.8, "notes": "Solid Georgian brick wall. Open fireplace needs chimney balloon or flue damper."},
    {"code": "GF-03", "name": "Dining Room", "floor": "Ground Floor", "zone": "Thatched gable ended", "temp": 21.0, "len": 6.00, "wid": 5.50, "ht": 2.80, "ext_wall": 11.50, "u_wall": 1.80, "win_area": 5.0, "u_win": 4.80, "fl_area": 33.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30, "ach": 1.8, "notes": "Check floor void beneath timber boards. Good candidate for underfloor insulation."},
    {"code": "GF-04", "name": "Kitchen & Breakfast Area", "floor": "Ground Floor", "zone": "New build", "temp": 18.0, "len": 7.50, "wid": 6.50, "ht": 2.80, "ext_wall": 14.00, "u_wall": 0.22, "win_area": 8.0, "u_win": 1.40, "fl_area": 48.8, "u_fl": 0.15, "roof_area": 0.0, "u_roof": 0.15, "ach": 0.5, "notes": "New insulated cavity wall & insulated slab. High thermal performance."},
    {"code": "GF-05", "name": "Orangery / Garden Room", "floor": "Ground Floor", "zone": "Orangery", "temp": 20.0, "len": 10.00, "wid": 6.00, "ht": 2.80, "ext_wall": 22.00, "u_wall": 0.25, "win_area": 35.0, "u_win": 1.40, "fl_area": 60.0, "u_fl": 0.15, "roof_area": 60.0, "u_roof": 1.50, "ach": 0.6, "notes": "Large glazed roof lantern and double glazed perimeter. High heat loss density."},
    {"code": "GF-06", "name": "Snug / Family Sitting Room", "floor": "Ground Floor", "zone": "Thatched gable ended", "temp": 21.0, "len": 6.00, "wid": 5.00, "ht": 2.80, "ext_wall": 11.00, "u_wall": 1.80, "win_area": 4.5, "u_win": 4.80, "fl_area": 30.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30, "ach": 1.8, "notes": "Timber beam ceiling. Space available for Type 22 or Type 33 radiator."},
    {"code": "GF-07", "name": "Study / Home Office", "floor": "Ground Floor", "zone": "Thatched gable ended", "temp": 20.0, "len": 6.00, "wid": 4.00, "ht": 2.80, "ext_wall": 10.00, "u_wall": 1.80, "win_area": 3.8, "u_win": 4.80, "fl_area": 24.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30, "ach": 1.8, "notes": "Continuous occupancy during workdays. Comfort requires 20°C."},
    {"code": "GF-08", "name": "Utility & Boot Room", "floor": "Ground Floor", "zone": "Lean to West", "temp": 18.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 7.50, "u_wall": 2.20, "win_area": 2.2, "u_win": 4.80, "fl_area": 14.0, "u_fl": 1.10, "roof_area": 14.0, "u_roof": 2.00, "ach": 1.8, "notes": "Solid uninsulated wall & sloping roof. Prime candidate for roof & wall insulation."},
    {"code": "GF-09", "name": "Ground Floor Cloakroom / WC", "floor": "Ground Floor", "zone": "Lean to West", "temp": 18.0, "len": 2.50, "wid": 4.00, "ht": 2.60, "ext_wall": 2.50, "u_wall": 2.20, "win_area": 1.0, "u_win": 4.80, "fl_area": 10.0, "u_fl": 1.10, "roof_area": 10.0, "u_roof": 2.00, "ach": 1.8, "notes": "Small heated towel rail or compact radiator needed."},
    {"code": "GF-10", "name": "Plant Room / Back Scullery", "floor": "Ground Floor", "zone": "Lean to North", "temp": 16.0, "len": 6.00, "wid": 3.00, "ht": 2.60, "ext_wall": 6.00, "u_wall": 2.20, "win_area": 1.8, "u_win": 4.80, "fl_area": 18.0, "u_fl": 1.10, "roof_area": 18.0, "u_roof": 2.00, "ach": 1.8, "notes": "Buffer cylinder & manifolds location. Unintentional heat gains from pipework."},

    # First Floor (13 rooms)
    {"code": "FF-01", "name": "Upper Stair Landing & Main Gallery", "floor": "First Floor", "zone": "Georgian end", "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 2.80, "ext_wall": 10.75, "u_wall": 1.40, "win_area": 6.0, "u_win": 4.80, "fl_area": 28.1, "u_fl": 0.00, "roof_area": 28.1, "u_roof": 0.18, "ach": 1.4, "notes": "Connects with GF-01 hall. Ceiling below insulated attic."},
    {"code": "FF-02", "name": "First Floor East Corridor", "floor": "First Floor", "zone": "Thatched gable ended", "temp": 18.0, "len": 12.00, "wid": 1.80, "ht": 2.60, "ext_wall": 12.00, "u_wall": 1.80, "win_area": 3.0, "u_win": 4.80, "fl_area": 21.6, "u_fl": 0.00, "roof_area": 21.6, "u_roof": 0.30, "ach": 1.4, "notes": "Internal corridor with external end wall. Long pipe run access in floor void."},
    {"code": "FF-03", "name": "Master Bedroom", "floor": "First Floor", "zone": "Georgian end", "temp": 18.0, "len": 6.25, "wid": 5.50, "ht": 2.80, "ext_wall": 11.75, "u_wall": 1.40, "win_area": 5.8, "u_win": 4.80, "fl_area": 34.4, "u_fl": 0.00, "roof_area": 34.4, "u_roof": 0.18, "ach": 1.4, "notes": "Dual aspect windows. Large Type 22 radiator recommended."},
    {"code": "FF-04", "name": "Master En-Suite Bathroom", "floor": "First Floor", "zone": "Georgian end", "temp": 22.0, "len": 4.50, "wid": 3.00, "ht": 2.80, "ext_wall": 4.50, "u_wall": 1.40, "win_area": 1.8, "u_win": 4.80, "fl_area": 13.5, "u_fl": 0.00, "roof_area": 13.5, "u_roof": 0.18, "ach": 1.5, "notes": "CIBSE 22°C design temp. Underfloor heating mat or high-output towel radiator."},
    {"code": "FF-05", "name": "Bedroom 2 (Guest Suite)", "floor": "First Floor", "zone": "New build", "temp": 18.0, "len": 5.50, "wid": 5.00, "ht": 2.60, "ext_wall": 10.50, "u_wall": 0.22, "win_area": 4.2, "u_win": 1.40, "fl_area": 27.5, "u_fl": 0.00, "roof_area": 27.5, "u_roof": 0.15, "ach": 0.4, "notes": "New build high insulation section. Modest emitter sizing needed."},
    {"code": "FF-06", "name": "Bedroom 2 En-Suite", "floor": "First Floor", "zone": "New build", "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00, "u_wall": 0.22, "win_area": 1.2, "u_win": 1.40, "fl_area": 7.5, "u_fl": 0.00, "roof_area": 7.5, "u_roof": 0.15, "ach": 0.5, "notes": "Warm 22°C design requirement. Extract fan with backdraught shutter."},
    {"code": "FF-07", "name": "Bedroom 3", "floor": "First Floor", "zone": "Thatched gable ended", "temp": 18.0, "len": 5.50, "wid": 4.50, "ht": 2.60, "ext_wall": 10.00, "u_wall": 1.80, "win_area": 3.8, "u_win": 4.80, "fl_area": 24.8, "u_fl": 0.00, "roof_area": 24.8, "u_roof": 0.30, "ach": 1.4, "notes": "Dormer window and sloping ceiling. Ensure loft insulation reaches eaves."},
    {"code": "FF-08", "name": "Bedroom 4", "floor": "First Floor", "zone": "Thatched gable ended", "temp": 18.0, "len": 5.00, "wid": 4.50, "ht": 2.60, "ext_wall": 9.50, "u_wall": 1.80, "win_area": 3.5, "u_win": 4.80, "fl_area": 22.5, "u_fl": 0.00, "roof_area": 22.5, "u_roof": 0.30, "ach": 1.4, "notes": "Gable end wall exposed. Internal woodfibre insulation would reduce loss by 60%."},
    {"code": "FF-09", "name": "Bedroom 5 / Nursery", "floor": "First Floor", "zone": "New build", "temp": 18.0, "len": 4.50, "wid": 4.00, "ht": 2.60, "ext_wall": 8.50, "u_wall": 0.22, "win_area": 3.2, "u_win": 1.40, "fl_area": 18.0, "u_fl": 0.00, "roof_area": 18.0, "u_roof": 0.15, "ach": 0.4, "notes": "Modern insulated timber/cavity construction."},
    {"code": "FF-10", "name": "Family Bathroom", "floor": "First Floor", "zone": "Thatched gable ended", "temp": 22.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 4.00, "u_wall": 1.80, "win_area": 2.0, "u_win": 4.80, "fl_area": 14.0, "u_fl": 0.00, "roof_area": 14.0, "u_roof": 0.30, "ach": 1.5, "notes": "High target temp (22°C). Needs dedicated radiator plus heated towel rail."},
    {"code": "FF-11", "name": "Laundry & Linen Room", "floor": "First Floor", "zone": "New build", "temp": 18.0, "len": 3.50, "wid": 3.00, "ht": 2.60, "ext_wall": 3.50, "u_wall": 0.22, "win_area": 1.5, "u_win": 1.40, "fl_area": 10.5, "u_fl": 0.00, "roof_area": 10.5, "u_roof": 0.15, "ach": 0.5, "notes": "Washing/drying equipment. Internal heat gains help offset space heating."},
    {"code": "FF-12", "name": "Dressing Room / Walk-in Robe", "floor": "First Floor", "zone": "Georgian end", "temp": 18.0, "len": 4.00, "wid": 3.00, "ht": 2.80, "ext_wall": 4.00, "u_wall": 1.40, "win_area": 1.8, "u_win": 4.80, "fl_area": 12.0, "u_fl": 0.00, "roof_area": 12.0, "u_roof": 0.18, "ach": 1.4, "notes": "Fitted wardrobes on external wall create thermal shadowing. Keep background heat."},
    {"code": "FF-13", "name": "First Floor Shower Room / WC", "floor": "First Floor", "zone": "Thatched gable ended", "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00, "u_wall": 1.80, "win_area": 1.0, "u_win": 4.80, "fl_area": 7.5, "u_fl": 0.00, "roof_area": 7.5, "u_roof": 0.30, "ach": 1.5, "notes": "22°C bathroom comfort criteria. Compact high-efficiency emitter."}
]

def build_room_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 5_Room_Heat_Loss schedule worksheet."""
    tab_name = "5_Room_Heat_Loss"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=40, cols=32)

    total_cols = 30
    total_rows = 30
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Banner
    grid[0][0] = "5_Room_Heat_Loss: Room-by-Room Heat Loss Assessment & Low-Flow Radiator Sizing"
    grid[1][0] = "MCS / CIBSE BS EN 12831 Room Schedule: 10 Ground Floor & 13 First Floor Rooms, Sized for 45°C Heat Pump Flow Temperature."

    # Row 3: Category Groups
    grid[2][0] = "ROOM IDENTIFICATION"           # Cols A-D (0-3)
    grid[2][4] = "DESIGN & GEOMETRY"            # Cols E-J (4-9)
    grid[2][10] = "FABRIC LOSS CALCULATIONS (W)" # Cols K-V (10-21)
    grid[2][22] = "VENTILATION (W)"             # Cols W-X (22-23)
    grid[2][24] = "TOTAL HEAT LOSS & EMITTERS"  # Cols Y-AB (24-27)
    grid[2][28] = "SURVEY NOTES & INSPECTION"   # Cols AC-AD (28-29)

    # Row 4: Column Headers
    headers = [
        "Code",               # Col A (0)
        "Room Name",          # Col B (1)
        "Floor Level",        # Col C (2)
        "Zone / Wing",        # Col D (3)
        "Design Ti (°C)",     # Col E (4)
        "Length (m)",         # Col F (5)
        "Width (m)",          # Col G (6)
        "Area (m²)",          # Col H (7)
        "Height (m)",         # Col I (8)
        "Volume (m³)",        # Col J (9)
        "Ext Wall L (m)",     # Col K (10)
        "U Wall",             # Col L (11)
        "Wall Loss (W)",      # Col M (12)
        "Window Area (m²)",   # Col N (13)
        "U Window",           # Col O (14)
        "Window Loss (W)",    # Col P (15)
        "Exp Floor (m²)",     # Col Q (16)
        "U Floor",            # Col R (17)
        "Floor Loss (W)",     # Col S (18)
        "Ceiling Area (m²)",  # Col T (19)
        "U Ceiling",          # Col U (20)
        "Ceiling Loss (W)",   # Col V (21)
        "Infil (ACH)",        # Col W (22)
        "Vent Loss (W)",      # Col X (23)
        "Room Loss (W)",      # Col Y (24)
        "Intensity (W/m²)",   # Col Z (25)
        "Rad 45°C (ΔT30 W)",  # Col AA (26)
        "Boiler Rad (ΔT50 W)",# Col AB (27)
        "Recommended Emitter",# Col AC (28)
        "On-Site Survey Notes"# Col AD (29)
    ]
    for c_i, h in enumerate(headers):
        grid[3][c_i] = h

    # Populate 23 Rooms (Rows 5 to 27, indices 4 to 26)
    for r_idx, rm in enumerate(DEFAULT_ROOMS):
        r = 5 + r_idx  # 1-indexed row in spreadsheet (5 to 27)
        row_arr = [
            rm["code"],                                       # Col A
            rm["name"],                                       # Col B
            rm["floor"],                                      # Col C
            rm["zone"],                                       # Col D
            str(rm["temp"]),                                  # Col E
            str(rm["len"]),                                   # Col F
            str(rm["wid"]),                                   # Col G
            f"=F{r}*G{r}",                                    # Col H (Area)
            str(rm["ht"]),                                    # Col I
            f"=H{r}*I{r}",                                    # Col J (Volume)
            str(rm["ext_wall"]),                              # Col K
            str(rm["u_wall"]),                                # Col L
            f"=MAX(0, (K{r}*I{r}-N{r})*L{r}*(E{r}-'1_Inputs'!$C$5))", # Col M (Wall Loss W)
            str(rm["win_area"]),                              # Col N
            str(rm["u_win"]),                                 # Col O
            f"=N{r}*O{r}*(E{r}-'1_Inputs'!$C$5)",             # Col P (Window Loss W)
            str(rm["fl_area"]),                               # Col Q
            str(rm["u_fl"]),                                  # Col R
            f"=Q{r}*R{r}*(E{r}-'1_Inputs'!$C$6)",             # Col S (Floor Loss with Ground Temp C6!)
            str(rm["roof_area"]),                             # Col T
            str(rm["u_roof"]),                                # Col U
            f"=T{r}*U{r}*(E{r}-'1_Inputs'!$C$5)",             # Col V (Ceiling Loss W)
            str(rm["ach"]),                                   # Col W
            f"='1_Inputs'!$C$11*W{r}*J{r}*(E{r}-'1_Inputs'!$C$5)", # Col X (Vent Loss W)
            f"=SUM(M{r}, P{r}, S{r}, V{r}, X{r})",           # Col Y (Total Room Loss W)
            f"=Y{r}/H{r}",                                    # Col Z (W/m²)
            f"=Y{r}",                                         # Col AA (Req Rad at ΔT30)
            f"=ROUND(Y{r}*1.89, 0)",                          # Col AB (Boiler equivalent at ΔT50)
            f"=IF(Z{r}>100, \"Type 33 or 2x Type 22\", IF(Z{r}>65, \"Type 22 High-Output\", \"Type 21 / Underfloor\"))", # Col AC
            rm["notes"]                                       # Col AD
        ]
        grid[4 + r_idx] = row_arr

    # Row 28: Total Building Schedule Row (index 27)
    grid[27] = [
        "Total Whole House",
        "23 Assessed Rooms",
        "Ground & First",
        "All Wings",
        "=AVERAGE(E5:E27)",
        "-",
        "-",
        "=SUM(H5:H27)",
        "-",
        "=SUM(J5:J27)",
        "-",
        "-",
        "=SUM(M5:M27)",
        "=SUM(N5:N27)",
        "-",
        "=SUM(P5:P27)",
        "=SUM(Q5:Q27)",
        "-",
        "=SUM(S5:S27)",
        "=SUM(T5:T27)",
        "-",
        "=SUM(V5:V27)",
        "-",
        "=SUM(X5:X27)",
        "=SUM(Y5:Y27)",
        "=Y28/H28",
        "=SUM(AA5:AA27)",
        "=SUM(AB5:AB27)",
        "Whole House Emitters",
        "Survey Complete"
    ]

    # Write values into worksheet
    ws.update(values=grid, range_name=f"A1:AD{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze top 4 header rows
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Set Column widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 75))   # Code
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 230))  # Room Name
    fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 100))  # Floor
    fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 140))  # Zone
    fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 95))   # Ti
    for c_i in range(5, 24):
        fmt_reqs.append(create_set_column_width_request(ws.id, c_i, c_i + 1, 95))
    fmt_reqs.append(create_set_column_width_request(ws.id, 24, 25, 110)) # Room Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 25, 26, 105)) # Intensity
    fmt_reqs.append(create_set_column_width_request(ws.id, 26, 27, 120)) # Rad 45C
    fmt_reqs.append(create_set_column_width_request(ws.id, 27, 28, 120)) # Boiler Rad
    fmt_reqs.append(create_set_column_width_request(ws.id, 28, 29, 165)) # Recommended Emitter
    fmt_reqs.append(create_set_column_width_request(ws.id, 29, 30, 360)) # Notes

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

    # Category Group Headers (Row 3)
    group_spans = [
        (0, 4, THEME["SECONDARY_HEADER_BG"]),
        (4, 10, THEME["SECTION_HEADER_BG"]),
        (10, 22, {"red": 0.15, "green": 0.35, "blue": 0.45}),
        (22, 24, {"red": 0.12, "green": 0.30, "blue": 0.40}),
        (24, 28, THEME["PRIMARY_HEADER_BG"]),
        (28, 30, THEME["CARD_HEADER_BG"])
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

    # Data Rows (Rows 5 to 27)
    for r_i in range(4, 27):
        bg = THEME["ZEBRA_BG"] if r_i % 2 == 1 else {"red": 1.0, "green": 1.0, "blue": 1.0}
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, total_cols,
            bg_color=bg,
            font_color=THEME["DARK_TEXT"],
            font_size=9,
            align="RIGHT"
        ))
        # Left align Room Code, Name, Floor, Zone
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, 4,
            align="LEFT",
            font_size=9
        ))
        # Recommended Emitter & Notes (Cols 28, 29)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 28, 30,
            align="LEFT",
            font_size=9,
            font_color=THEME["MUTED_TEXT"]
        ))

    # Number formats
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 4, 5, number_format=FORMATS["TEMP_C"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 5, 8, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 7, 8, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 12, 13, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 15, 16, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 18, 19, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 21, 22, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 23, 25, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 25, 26, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 26, 28, number_format=FORMATS["INTEGER"]))

    # Total Row (Row 28, index 27)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 27, 28, 0, total_cols,
        bg_color=THEME["TOTAL_BG"],
        font_color=THEME["DARK_TEXT"],
        bold=True,
        font_size=10,
        align="RIGHT"
    ))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 27, 28, 0, 4,
        align="LEFT"
    ))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 7, 8, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 12, 13, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 15, 16, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 18, 19, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 21, 22, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 23, 28, number_format=FORMATS["INTEGER"]))

    # Accent on Total Room Heat Loss
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 27, 28, 24, 25,
        bg_color=THEME["ACCENT_BG"],
        font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
        bold=True,
        font_size=11,
        align="RIGHT"
    ))

    return ws, fmt_reqs
