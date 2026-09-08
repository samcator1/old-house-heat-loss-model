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
from ..config import THEME, FORMATS, INFILTRATION_QUESTIONNAIRE, WINDOW_SPECIFICATIONS
from ..formatting import (
    create_repeat_cell_request,
    create_set_column_width_request,
    create_freeze_pane_request,
    create_merge_cells_request,
    create_data_validation_request
)

# 23 Rooms Specification (10 Ground Floor, 13 First Floor)
DEFAULT_ROOMS = [
    # Ground Floor (10 rooms)
    {
        "code": "GF-01", "name": "Entrance & Stair Hall (GF to FF)", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 5.60, "ext_wall": 10.75, "u_wall": 1.40,
        "win_area": 6.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 28.1, "u_fl": 0.80, "roof_area": 0.0, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Double-height open stair volume connecting GF to FF. Check draught seal on front door."
    },
    {
        "code": "GF-02", "name": "Drawing Room / Main Living", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 21.0, "len": 8.50, "wid": 6.25, "ht": 2.80, "ext_wall": 14.75, "u_wall": 1.40,
        "win_area": 9.2, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 53.1, "u_fl": 0.80, "roof_area": 0.0, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "Open Fireplace (Unsealed Flue)",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Solid Georgian brick wall. Open fireplace needs chimney balloon or flue damper to cut 0.6 ACH."
    },
    {
        "code": "GF-03", "name": "Dining Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 6.00, "wid": 5.50, "ht": 2.80, "ext_wall": 11.50, "u_wall": 1.80,
        "win_area": 5.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 33.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Check floor void beneath timber boards. Good candidate for underfloor insulation."
    },
    {
        "code": "GF-04", "name": "Kitchen & Breakfast Area", "floor": "Ground Floor", "zone": "New build",
        "temp": 18.0, "len": 7.50, "wid": 6.50, "ht": 2.80, "ext_wall": 14.00, "u_wall": 0.22,
        "win_area": 8.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 48.8, "u_fl": 0.15, "roof_area": 0.0, "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "New insulated cavity wall & insulated slab. High thermal and airtight performance."
    },
    {
        "code": "GF-05", "name": "Orangery / Garden Room", "floor": "Ground Floor", "zone": "Orangery",
        "temp": 20.0, "len": 10.00, "wid": 6.00, "ht": 2.80, "ext_wall": 22.00, "u_wall": 0.25,
        "win_area": 35.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 60.0, "u_fl": 0.15, "roof_area": 60.0, "u_roof": 1.50,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Large glazed roof lantern and double glazed perimeter. High heat loss density."
    },
    {
        "code": "GF-06", "name": "Snug / Family Sitting Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 6.00, "wid": 5.00, "ht": 2.80, "ext_wall": 11.00, "u_wall": 1.80,
        "win_area": 4.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 30.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "Room-Sealed Stove with Ext Air",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Timber beam ceiling. Space available for Type 22 or Type 33 radiator."
    },
    {
        "code": "GF-07", "name": "Study / Home Office", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 20.0, "len": 6.00, "wid": 4.00, "ht": 2.80, "ext_wall": 10.00, "u_wall": 1.80,
        "win_area": 3.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 24.0, "u_fl": 0.60, "roof_area": 0.0, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Continuous occupancy during workdays. Comfort requires 20°C."
    },
    {
        "code": "GF-08", "name": "Utility & Boot Room", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 18.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 7.50, "u_wall": 2.20,
        "win_area": 2.2, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 14.0, "u_fl": 1.10, "roof_area": 14.0, "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Solid uninsulated wall & sloping roof. Prime candidate for roof & wall insulation."
    },
    {
        "code": "GF-09", "name": "Ground Floor Cloakroom / WC", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 18.0, "len": 2.50, "wid": 4.00, "ht": 2.60, "ext_wall": 2.50, "u_wall": 2.20,
        "win_area": 1.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 10.0, "u_fl": 1.10, "roof_area": 10.0, "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Small heated towel rail or compact radiator needed."
    },
    {
        "code": "GF-10", "name": "Plant Room / Back Scullery", "floor": "Ground Floor", "zone": "Lean to North",
        "temp": 16.0, "len": 6.00, "wid": 3.00, "ht": 2.60, "ext_wall": 6.00, "u_wall": 2.20,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 18.0, "u_fl": 1.10, "roof_area": 18.0, "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Buffer cylinder & manifolds location. Unintentional heat gains from pipework."
    },

    # First Floor (13 rooms)
    {
        "code": "FF-01", "name": "Upper Stair Landing & Main Gallery", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 2.80, "ext_wall": 10.75, "u_wall": 1.40,
        "win_area": 6.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 28.1, "u_fl": 0.00, "roof_area": 28.1, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Connects with GF-01 hall. Ceiling below insulated attic."
    },
    {
        "code": "FF-02", "name": "First Floor East Corridor", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 12.00, "wid": 1.80, "ht": 2.60, "ext_wall": 12.00, "u_wall": 1.80,
        "win_area": 3.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 21.6, "u_fl": 0.00, "roof_area": 21.6, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Internal corridor with external end wall. Long pipe run access in floor void."
    },
    {
        "code": "FF-03", "name": "Master Bedroom", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 6.25, "wid": 5.50, "ht": 2.80, "ext_wall": 11.75, "u_wall": 1.40,
        "win_area": 5.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 34.4, "u_fl": 0.00, "roof_area": 34.4, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Retrofitted Brush / Pile Draught Seals", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Dual aspect windows. Large Type 22 radiator recommended."
    },
    {
        "code": "FF-04", "name": "Master En-Suite Bathroom", "floor": "First Floor", "zone": "Georgian end",
        "temp": 22.0, "len": 4.50, "wid": 3.00, "ht": 2.80, "ext_wall": 4.50, "u_wall": 1.40,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 13.5, "u_fl": 0.00, "roof_area": 13.5, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Retrofitted Brush / Pile Draught Seals", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Attic Downlights / Unsealed Loft Hatch",
        "notes": "CIBSE 22°C design temp. Underfloor heating mat or high-output towel radiator."
    },
    {
        "code": "FF-05", "name": "Bedroom 2 (Guest Suite)", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 5.50, "wid": 5.00, "ht": 2.60, "ext_wall": 10.50, "u_wall": 0.22,
        "win_area": 4.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 27.5, "u_fl": 0.00, "roof_area": 27.5, "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "New build high insulation section. Modest emitter sizing needed."
    },
    {
        "code": "FF-06", "name": "Bedroom 2 En-Suite", "floor": "First Floor", "zone": "New build",
        "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00, "u_wall": 0.22,
        "win_area": 1.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 7.5, "u_fl": 0.00, "roof_area": 7.5, "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Warm 22°C design requirement. Extract fan with backdraught shutter."
    },
    {
        "code": "FF-07", "name": "Bedroom 3", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 5.50, "wid": 4.50, "ht": 2.60, "ext_wall": 10.00, "u_wall": 1.80,
        "win_area": 3.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 24.8, "u_fl": 0.00, "roof_area": 24.8, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Dormer window and sloping ceiling. Ensure loft insulation reaches eaves."
    },
    {
        "code": "FF-08", "name": "Bedroom 4", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 5.00, "wid": 4.50, "ht": 2.60, "ext_wall": 9.50, "u_wall": 1.80,
        "win_area": 3.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 22.5, "u_fl": 0.00, "roof_area": 22.5, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Gable end wall exposed. Internal woodfibre insulation would reduce loss by 60%."
    },
    {
        "code": "FF-09", "name": "Bedroom 5 / Nursery", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 4.50, "wid": 4.00, "ht": 2.60, "ext_wall": 8.50, "u_wall": 0.22,
        "win_area": 3.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 18.0, "u_fl": 0.00, "roof_area": 18.0, "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Modern insulated timber/cavity construction."
    },
    {
        "code": "FF-10", "name": "Family Bathroom", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 22.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 4.00, "u_wall": 1.80,
        "win_area": 2.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 14.0, "u_fl": 0.00, "roof_area": 14.0, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "High target temp (22°C). Needs dedicated radiator plus heated towel rail."
    },
    {
        "code": "FF-11", "name": "Laundry & Linen Room", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 3.50, "wid": 3.00, "ht": 2.60, "ext_wall": 3.50, "u_wall": 0.22,
        "win_area": 1.5, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40, "fl_area": 10.5, "u_fl": 0.00, "roof_area": 10.5, "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Washing/drying equipment. Internal heat gains help offset space heating."
    },
    {
        "code": "FF-12", "name": "Dressing Room / Walk-in Robe", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 4.00, "wid": 3.00, "ht": 2.80, "ext_wall": 4.00, "u_wall": 1.40,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 12.0, "u_fl": 0.00, "roof_area": 12.0, "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Fitted wardrobes on external wall create thermal shadowing. Keep background heat."
    },
    {
        "code": "FF-13", "name": "First Floor Shower Room / WC", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00, "u_wall": 1.80,
        "win_area": 1.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80, "fl_area": 7.5, "u_fl": 0.00, "roof_area": 7.5, "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "22°C bathroom comfort criteria. Compact high-efficiency emitter."
    }
]
def build_room_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
    """Builds and formats the 2_Room_Heat_Loss master schedule worksheet."""
    tab_name = "2_Room_Heat_Loss"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=40, cols=36)

    total_cols = 36
    total_rows = 30
    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Banner
    grid[0][0] = "2_Room_Heat_Loss: Room-by-Room Heat Loss Assessment & Low-Flow Radiator Sizing"
    grid[1][0] = "Master MCS / CIBSE BS EN 12831 Room Schedule: 10 Ground Floor & 13 First Floor Rooms with Dynamic Infiltration Questionnaire."

    # Row 3: Category Groups
    grid[2][0] = "ROOM IDENTIFICATION"                   # Cols A-D (0-3)
    grid[2][4] = "DESIGN & GEOMETRY"                    # Cols E-J (4-9)
    grid[2][10] = "FABRIC LOSS CALCULATIONS (W)"         # Cols K-W (10-22)
    grid[2][23] = "INFILTRATION & VENTILATION QUESTIONNAIRE" # Cols X-AD (23-29)
    grid[2][30] = "TOTAL HEAT LOSS & EMITTERS"          # Cols AE-AH (30-33)
    grid[2][34] = "SPECIFICATION & NOTES"               # Cols AI-AJ (34-35)

    # Row 4: Column Headers
    headers = [
        "Code",                           # Col A (0)
        "Room Name",                      # Col B (1)
        "Floor Level",                    # Col C (2)
        "Zone / Wing",                    # Col D (3)
        "Design Ti (°C)",                 # Col E (4)
        "Length (m)",                     # Col F (5)
        "Width (m)",                      # Col G (6)
        "Area (m²)",                      # Col H (7)
        "Height (m)",                     # Col I (8)
        "Volume (m³)",                    # Col J (9)
        "Ext Wall L (m)",                 # Col K (10)
        "U Wall",                         # Col L (11)
        "Wall Loss (W)",                  # Col M (12)
        "Window Area (m²)",               # Col N (13)
        "Window Specification",           # Col O (14) - Dropdown Glazing
        "U Window",                       # Col P (15) - Formula VLOOKUP
        "Window Loss (W)",                # Col Q (16)
        "Exp Floor (m²)",                 # Col R (17)
        "U Floor",                        # Col S (18)
        "Floor Loss (W)",                 # Col T (19)
        "Ceiling Area (m²)",              # Col U (20)
        "U Ceiling",                      # Col V (21)
        "Ceiling Loss (W)",               # Col W (22)
        "Base Construction",              # Col X (23) - Dropdown Q1
        "Chimney / Flue",                 # Col Y (24) - Dropdown Q2
        "Windows & Doors",                # Col Z (25) - Dropdown Q3
        "Floor Construction",             # Col AA (26) - Dropdown Q4
        "Ceiling Boundary",               # Col AB (27) - Dropdown Q5
        "Calculated ACH",                 # Col AC (28) - Formula
        "Vent Loss (W)",                  # Col AD (29) - Formula
        "Room Loss (W)",                  # Col AE (30) - Formula
        "Intensity (W/m²)",               # Col AF (31) - Formula
        "Rad 45°C (ΔT30 W)",              # Col AG (32) - Formula
        "Boiler Rad (ΔT50 W)",            # Col AH (33) - Formula
        "Recommended Emitter",            # Col AI (34) - Formula
        "On-Site Survey Notes"            # Col AJ (35) - Text
    ]
    for c_i, h in enumerate(headers):
        grid[3][c_i] = h

    # Populate 23 Rooms (Rows 5 to 27, indices 4 to 26)
    for r_idx, rm in enumerate(DEFAULT_ROOMS):
        r = 5 + r_idx  # 1-indexed row in spreadsheet (5 to 27)
        row_arr = [
            rm["code"],                                       # Col A (0)
            rm["name"],                                       # Col B (1)
            rm["floor"],                                      # Col C (2)
            rm["zone"],                                       # Col D (3)
            str(rm["temp"]),                                  # Col E (4)
            str(rm["len"]),                                   # Col F (5)
            str(rm["wid"]),                                   # Col G (6)
            f"=F{r}*G{r}",                                    # Col H (7) (Area)
            str(rm["ht"]),                                    # Col I (8)
            f"=H{r}*I{r}",                                    # Col J (9) (Volume)
            str(rm["ext_wall"]),                              # Col K (10)
            str(rm["u_wall"]),                                # Col L (11)
            f"=MAX(0, (K{r}*I{r}-N{r})*L{r}*(E{r}-'1_Inputs'!$C$5))", # Col M (12) (Wall Loss W)
            str(rm["win_area"]),                              # Col N (13)
            rm["win_spec"],                                   # Col O (14) (Window Specification)
            f"=VLOOKUP(O{r}, '1_Inputs'!$B$98:$C$108, 2, FALSE)", # Col P (15) (U Window Formula)
            f"=N{r}*P{r}*(E{r}-'1_Inputs'!$C$5)",             # Col Q (16) (Window Loss W)
            str(rm["fl_area"]),                               # Col R (17)
            str(rm["u_fl"]),                                  # Col S (18)
            f"=R{r}*S{r}*(E{r}-'1_Inputs'!$C$6)",             # Col T (19) (Floor Loss with Ground Temp C6!)
            str(rm["roof_area"]),                             # Col U (20)
            str(rm["u_roof"]),                                # Col V (21)
            f"=U{r}*V{r}*(E{r}-'1_Inputs'!$C$5)",             # Col W (22) (Ceiling Loss W)
            rm["q_base"],                                     # Col X (23) - Dropdown Q1
            rm["q_chimney"],                                  # Col Y (24) - Dropdown Q2
            rm["q_win"],                                      # Col Z (25) - Dropdown Q3
            rm["q_floor"],                                    # Col AA (26) - Dropdown Q4
            rm["q_ceil"],                                     # Col AB (27) - Dropdown Q5
            f"=VLOOKUP(X{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(Y{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(Z{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AA{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AB{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE)", # Col AC (28)
            f"='1_Inputs'!$C$11*AC{r}*J{r}*(E{r}-'1_Inputs'!$C$5)", # Col AD (29) (Vent Loss W)
            f"=SUM(M{r}, Q{r}, T{r}, W{r}, AD{r})",           # Col AE (30) (Total Room Loss W)
            f"=AE{r}/H{r}",                                   # Col AF (31) (W/m²)
            f"=AE{r}",                                        # Col AG (32) (Req Rad at ΔT30)
            f"=ROUND(AE{r}*1.89, 0)",                         # Col AH (33) (Boiler equivalent at ΔT50)
            f"=IF(AF{r}>100, \"Type 33 or 2x Type 22\", IF(AF{r}>65, \"Type 22 High-Output\", \"Type 21 / Underfloor\"))", # Col AI (34)
            rm["notes"]                                       # Col AJ (35)
        ]
        grid[4 + r_idx] = row_arr

    # Row 28: Total Building Schedule Row (index 27)
    grid[27] = [
        "Total Whole House",
        "23 Assessed Rooms",
        "Ground & First",
        "All Wings",
        "=SUMPRODUCT(E5:E27, H5:H27)/SUM(H5:H27)",
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
        "-",
        "=SUM(Q5:Q27)",
        "=SUM(R5:R27)",
        "-",
        "=SUM(T5:T27)",
        "=SUM(U5:U27)",
        "-",
        "=SUM(W5:W27)",
        "-",
        "-",
        "-",
        "-",
        "-",
        "=AVERAGE(AC5:AC27)",
        "=SUM(AD5:AD27)",
        "=SUM(AE5:AE27)",
        "=AE28/H28",
        "=SUM(AG5:AG27)",
        "=SUM(AH5:AH27)",
        "Whole House Emitters",
        "Survey Complete"
    ]

    # Write values into worksheet
    ws.update(values=grid, range_name=f"A1:AJ{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze top 4 header rows
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Set Column widths
    fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 75))    # Code
    fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 230))   # Room Name
    fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 100))   # Floor
    fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 140))   # Zone
    fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 95))    # Ti
    for c_i in range(5, 14):
        fmt_reqs.append(create_set_column_width_request(ws.id, c_i, c_i + 1, 95)) # Len to Win Area
    fmt_reqs.append(create_set_column_width_request(ws.id, 14, 15, 240)) # Window Specification
    fmt_reqs.append(create_set_column_width_request(ws.id, 15, 16, 95))  # U Window
    for c_i in range(16, 23):
        fmt_reqs.append(create_set_column_width_request(ws.id, c_i, c_i + 1, 95)) # Win Loss to Ceil Loss
    fmt_reqs.append(create_set_column_width_request(ws.id, 23, 24, 230)) # Base Construction
    fmt_reqs.append(create_set_column_width_request(ws.id, 24, 25, 220)) # Chimney / Flue
    fmt_reqs.append(create_set_column_width_request(ws.id, 25, 26, 230)) # Windows & Doors
    fmt_reqs.append(create_set_column_width_request(ws.id, 26, 27, 230)) # Floor Type
    fmt_reqs.append(create_set_column_width_request(ws.id, 27, 28, 230)) # Ceiling Boundary
    fmt_reqs.append(create_set_column_width_request(ws.id, 28, 29, 105)) # Calc ACH
    fmt_reqs.append(create_set_column_width_request(ws.id, 29, 30, 105)) # Vent Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 30, 31, 115)) # Room Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 31, 32, 105)) # Intensity
    fmt_reqs.append(create_set_column_width_request(ws.id, 32, 33, 120)) # Rad 45C
    fmt_reqs.append(create_set_column_width_request(ws.id, 33, 34, 120)) # Boiler Rad
    fmt_reqs.append(create_set_column_width_request(ws.id, 34, 35, 165)) # Recommended Emitter
    fmt_reqs.append(create_set_column_width_request(ws.id, 35, 36, 360)) # Notes

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
        (0, 4, THEME["SECONDARY_HEADER_BG"]),                    # Room ID (Cols A-D)
        (4, 10, THEME["SECTION_HEADER_BG"]),                    # Geometry (Cols E-J)
        (10, 23, {"red": 0.15, "green": 0.35, "blue": 0.45}),    # Fabric (Cols K-W)
        (23, 30, {"red": 0.10, "green": 0.32, "blue": 0.42}),   # Questionnaire & Vent (Cols X-AD)
        (30, 34, THEME["PRIMARY_HEADER_BG"]),                   # Totals & Emitters (Cols AE-AH)
        (34, 36, THEME["CARD_HEADER_BG"])                      # Notes (Cols AI-AJ)
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
        # Window Specification Dropdown (Col 14 / O)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 14, 15,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # U Window Formula (Col 15 / P)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 15, 16,
            align="RIGHT",
            font_size=9,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Questionnaire Dropdowns (Cols 23 to 27 / X to AB)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 23, 28,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # Calculated ACH (Col 28 / AC)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 28, 29,
            bg_color=THEME["CALC_BG"],
            bold=True,
            align="RIGHT",
            font_size=10,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Recommended Emitter & Notes (Cols 34, 35 / AI, AJ)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 34, 36,
            align="LEFT",
            font_size=9,
            font_color=THEME["MUTED_TEXT"]
        ))

    # Number formats
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 4, 5, number_format=FORMATS["TEMP_C"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 5, 9, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 10, 11, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 11, 12, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 12, 13, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 13, 14, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 15, 16, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 16, 17, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 17, 18, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 18, 19, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 19, 20, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 20, 21, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 21, 22, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 22, 23, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 28, 29, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 29, 31, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 31, 32, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, 27, 32, 34, number_format=FORMATS["INTEGER"]))

    # Total Row (Row 28, index 27)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 27, 28, 0, total_cols,
        bg_color=THEME["TOTAL_BG"],
        font_color=THEME["DARK_TEXT"],
        bold=True,
        font_size=10,
        align="RIGHT"
    ))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 0, 4, align="LEFT"))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 7, 8, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 12, 13, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 13, 14, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 16, 17, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 17, 18, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 19, 20, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 20, 21, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 22, 23, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 28, 29, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 29, 31, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 31, 32, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 27, 28, 32, 34, number_format=FORMATS["INTEGER"]))

    # Accent on Total Room Heat Loss (Col AE / 30)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 27, 28, 30, 31,
        bg_color=THEME["ACCENT_BG"],
        font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
        bold=True,
        font_size=11,
        align="RIGHT"
    ))

    # Add Native Google Sheets Dropdown Data Validations
    # 1. Window Specification (Col 14 / O)
    fmt_reqs.append(create_data_validation_request(
        sheet_id=ws.id,
        start_row=4,
        end_row=27,
        start_col=14,
        end_col=15,
        options=[opt["label"] for opt in WINDOW_SPECIFICATIONS],
        show_custom_ui=True,
        strict=False
    ))

    # 2. Infiltration Questionnaire Dropdowns (Cols 23 to 27 / X to AB)
    q_categories = [
        (23, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["base_construction"]]),
        (24, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["chimney_flue"]]),
        (25, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["windows_doors"]]),
        (26, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["floor_construction"]]),
        (27, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["ceiling_boundary"]]),
    ]
    for col_idx, options in q_categories:
        fmt_reqs.append(create_data_validation_request(
            sheet_id=ws.id,
            start_row=4,
            end_row=27,
            start_col=col_idx,
            end_col=col_idx + 1,
            options=options,
            show_custom_ui=True,
            strict=False
        ))

    return ws, fmt_reqs
