"""
Room-by-Room Heat Loss Assessment & Emitter Sizing Tab ('2_Room_Heat_Loss').
MCS MIS 3005-D / CIBSE DHDG / BS EN 12831-1:2017 compliant room-by-room schedule:
- 44-column canonical architecture managed via RoomCol schema
- Disaggregated fabric transmission (Walls, Windows, Doors, Floors, Ceilings)
- +10% Thermal Bridging Allowance per MCS MIS 3005-D
- BS EN 12831 6-factor infiltration assessment (Walls, Windows, Floors, Ceilings, Chimneys, Mechanical Extract)
- Planned Emitter Type selection (Type 22, Type 33, Type 21, Type 11, Underfloor Heating, Fan Convector)
- Low-temperature 45°C radiator sizing derated to BS EN 442 catalogue ratings (ΔT50)
- Design water flow rates (l/h) and pipe sizing recommendations
- Executive MCS Audit Header Block
"""
from typing import Tuple, List, Dict, Any
from pathlib import Path
import csv
import gspread
from ..config import (
    THEME,
    FORMATS,
    CHIMNEY_SPECIFICATIONS,
    MECHANICAL_VENTILATION_SPECIFICATIONS,
    EMITTER_SPECIFICATIONS,
    ROOM_TYPE_SPECIFICATIONS,
    CIBSE_DESIGN_TEMPS,
    WINDOW_SPECIFICATIONS,
    CEILING_SPECIFICATIONS,
    WALL_SPECIFICATIONS,
    FLOOR_SPECIFICATIONS,
    DOOR_SPECIFICATIONS
)
from ..schema import RoomCol
from ..formatting import (
    create_repeat_cell_request,
    create_set_column_width_request,
    create_freeze_pane_request,
    create_merge_cells_request,
    create_unmerge_cells_request,
    create_data_validation_request
)

VALID_CHIMNEY = {opt["label"] for opt in CHIMNEY_SPECIFICATIONS}
VALID_MECH = {opt["label"] for opt in MECHANICAL_VENTILATION_SPECIFICATIONS}
VALID_EMITTER = {opt["label"] for opt in EMITTER_SPECIFICATIONS}
VALID_ROOM_TYPE = {opt["label"] for opt in ROOM_TYPE_SPECIFICATIONS}

DEFAULT_ROOMS = [
    # Ground Floor (10 rooms)
    {
        "code": "GF-01", "name": "Entrance & Stair Hall (GF to FF)", "floor": "Ground Floor", "zone": "Georgian end",
        "room_type": "Hallway / Stairs / Circulation", "room_type": "Living Room / Sitting Room", "room_type": "Dining Room", "room_type": "Kitchen / Breakfast Room", "room_type": "Utility Room / Laundry", "room_type": "Playroom / Family Snug", "room_type": "Home Office / Study", "room_type": "Orangery / Conservatory / Garden Room", "room_type": "Cloakroom / Downstairs WC", "room_type": "Side Entrance Lobby / Porch", "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 5.60, "ext_wall": 10.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 6.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 2.0, "door_spec": "Solid Timber Door (Uninsulated / Historic Plank)", "u_door": 3.00,
        "fl_area": 28.1, "floor_spec": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "u_fl": 0.80,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Double-height open stair volume connecting GF to FF. Front entrance door."
    },
    {
        "code": "GF-02", "name": "Drawing Room / Main Living", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 21.0, "len": 8.50, "wid": 6.25, "ht": 2.80, "ext_wall": 14.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 9.2, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 0.0, "door_spec": "No External Door (Internal Boundary Only)", "u_door": 0.00,
        "fl_area": 53.1, "floor_spec": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "u_fl": 0.80,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "chimney": "Open Fireplace (Unsealed Flue)",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Solid Georgian brick wall. Open fireplace needs chimney balloon or flue damper to cut 0.6 ACH."
    },
    {
        "code": "GF-03", "name": "Dining Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 6.00, "wid": 5.50, "ht": 2.80, "ext_wall": 11.50,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 5.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 0.0, "door_spec": "No External Door (Internal Boundary Only)", "u_door": 0.00,
        "fl_area": 33.0, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.30,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Check floor void beneath timber boards. Good candidate for underfloor insulation."
    },
    {
        "code": "GF-04", "name": "Kitchen & Breakfast Area", "floor": "Ground Floor", "zone": "New build",
        "temp": 18.0, "len": 7.50, "wid": 6.50, "ht": 2.80, "ext_wall": 14.00,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 8.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "door_area": 2.0, "door_spec": "Modern High-Performance / Composite Insulated Door", "u_door": 1.20,
        "fl_area": 48.8, "floor_spec": "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)", "u_fl": 0.15,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.15,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "Intermittent Extract Fan (Kitchen / Utility - 30-60 l/s)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "New insulated cavity wall & insulated slab. High thermal and airtight performance."
    },
    {
        "code": "GF-05", "name": "Utility Room & Secondary Entrance", "floor": "Ground Floor", "zone": "New build",
        "temp": 18.0, "len": 4.00, "wid": 3.50, "ht": 2.80, "ext_wall": 7.50,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 2.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "door_area": 1.8, "door_spec": "Modern High-Performance / Composite Insulated Door", "u_door": 1.20,
        "fl_area": 14.0, "floor_spec": "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)", "u_fl": 0.15,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.15,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "Intermittent Extract Fan (Kitchen / Utility - 30-60 l/s)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Secondary entrance door. External boot wash access."
    },
    {
        "code": "GF-06", "name": "Family Snug / TV Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 5.00, "wid": 4.50, "ht": 2.60, "ext_wall": 9.50,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 3.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 0.0, "door_spec": "No External Door (Internal Boundary Only)", "u_door": 0.00,
        "fl_area": 22.5, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.30,
        "chimney": "Room-Sealed Stove with Ext Air",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Woodburner installed with external air supply (minimal infiltration penalty)."
    },
    {
        "code": "GF-07", "name": "Study / Home Office", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 21.0, "len": 4.50, "wid": 4.00, "ht": 2.80, "ext_wall": 8.50,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 4.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 0.0, "door_spec": "No External Door (Internal Boundary Only)", "u_door": 0.00,
        "fl_area": 18.0, "floor_spec": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "u_fl": 0.80,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Working from home occupancy. High priority for daytime heating comfort."
    },
    {
        "code": "GF-08", "name": "Orangery / Garden Room", "floor": "Ground Floor", "zone": "Orangery",
        "temp": 21.0, "len": 6.50, "wid": 4.50, "ht": 3.00, "ext_wall": 15.50,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 14.5, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "door_area": 3.6, "door_spec": "Modern French / Bi-fold Glazed Doors (Low-E Double)", "u_door": 1.40,
        "fl_area": 29.3, "floor_spec": "Solid Concrete: Moderate Insulation (50mm PIR / 1990s)", "u_fl": 0.35,
        "roof_area": 15.0, "ceil_spec": "Glazed Roof Lantern / Sloping Skylight (Orangery)", "u_roof": 1.50,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 22 (Double Convector)",
        "notes": "Large glazed perimeter & roof lantern. Consider underfloor heating or trench heaters."
    },
    {
        "code": "GF-09", "name": "Downstairs Cloakroom / WC", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 2.50, "wid": 1.80, "ht": 2.80, "ext_wall": 2.50,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 0.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 0.0, "door_spec": "No External Door (Internal Boundary Only)", "u_door": 0.00,
        "fl_area": 4.5, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "Intermittent Extract Fan (Bathroom / WC - 15 l/s)",
        "emitter_type": "Type 21 (Single Convector)",
        "notes": "Compact room. Small single panel radiator adequate."
    },
    {
        "code": "GF-10", "name": "Side Entrance Lobby / Porch", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 16.0, "len": 2.00, "wid": 2.00, "ht": 2.40, "ext_wall": 4.00,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 1.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "door_area": 1.8, "door_spec": "Part-Glazed External Door (Single Glazed)", "u_door": 3.60,
        "fl_area": 4.0, "floor_spec": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "u_fl": 1.10,
        "roof_area": 4.0, "ceil_spec": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "u_roof": 2.00,
        "chimney": "No Chimney / Permanently Sealed",
        "mech_vent": "None (Natural Infiltration Only)",
        "emitter_type": "Type 21 (Single Convector)",
        "notes": "Side entrance lobby (2.0m x 2.0m). External door with single glazing / draught seals."
    }
]

def safe_float(val: Any, default: float = 0.0) -> float:
    """Safely converts unformatted cell values to float."""
    if val is None or val == "":
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("°C", "").replace("m²", "").replace("m³", "").replace("m", "").replace(",", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return default

def safe_str(val: Any, default: str = "") -> str:
    """Safely converts cell values to non-empty string."""
    if val is None:
        return default
    s = str(val).strip()
    return s if s != "" else default

def map_u_to_wall_spec(u_val: float, zone: str = "") -> str:
    z = zone.lower()
    if u_val <= 0.05:
        return "Party Wall / Heated Boundary (No Heat Loss)"
    if u_val <= 0.15:
        return "High-Performance New Build (150mm+ PIR / Passivhaus)"
    if u_val <= 0.23:
        return "Modern Building Regs Cavity (100mm PIR / Full Fill)"
    if u_val <= 0.26:
        return "Solid Wall + 100mm External Wall Insulation (EWI)"
    if u_val <= 0.35:
        return "Solid Wall + 100mm Woodfibre / PIR IWI"
    if u_val <= 0.48:
        return "Partial-Fill Cavity Wall (1980s-1990s)"
    if u_val <= 0.65:
        if "cavity" in z:
            return "Retrofilled Cavity Wall (Blown Mineral / Bead)"
        return "Solid Wall + 50mm Breathable Woodfibre IWI"
    if u_val <= 1.05 and ("garage" in z or "unheated" in z or "cellar" in z or "store" in z):
        return "Internal Partition to Unheated Garage / Store / Cellar"
    if u_val <= 1.10:
        return "Cob / Earth Construction (500mm+)"
    if 1.30 <= u_val <= 1.45:
        return 'Solid Brick: 18" / 450mm (Georgian Facade)'
    if 1.46 <= u_val <= 1.60:
        return "Uninsulated Cavity Wall (Pre-1976)"
    if 1.61 <= u_val <= 1.75:
        return 'Solid Brick: 13.5" / 330mm (Uninsulated)'
    if 1.76 <= u_val <= 1.95:
        if "timber" in z:
            return "Historic Timber Frame (Wattle & Daub / Nogging)"
        return "Solid Stone: 450-500mm Sandstone / Limestone"
    if u_val > 1.95:
        if "stone" in z or "granite" in z:
            return "Solid Stone: 500-600mm Dense Granite / Whinstone"
        return 'Solid Brick: 9" / 225mm (Uninsulated)'
    return 'Solid Brick: 18" / 450mm (Georgian Facade)'

def map_u_to_floor_spec(u_val: float, floor_level: str = "", zone: str = "") -> str:
    fl = floor_level.lower()
    z = zone.lower()
    if "first" in fl or "ff" in fl or "upper" in fl:
        return "Intermediate Floor (Heated Space Below)"
    if u_val <= 0.05:
        return "Intermediate Floor (Heated Space Below)"
    if u_val <= 0.12:
        return "High-Performance Insulated Slab (150mm+ PIR / Passivhaus)"
    if u_val <= 0.22:
        if "lime" in z or "limecrete" in z:
            return "Historic Breathable Insulated Limecrete Floor (Foamed Glass / Cork Lime)"
        return "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)"
    if u_val <= 0.40:
        if "concrete" in z or "slab" in z:
            return "Solid Concrete: Moderate Insulation (50mm PIR / 1990s)"
        return "Suspended Timber: Insulated (50mm Quilt / Board)"
    if u_val <= 0.50:
        if "cellar" in z or "basement" in z or "void" in z:
            return "Suspended Timber / Floor over Unheated Cellar Void"
        return "Solid Concrete: Perimeter Insulation (1980s Standard)"
    if u_val <= 0.70:
        return "Suspended Timber: Carpet & Underlay over Uninsulated Void"
    if u_val <= 0.95:
        if "concrete" in z or "slab" in z:
            return "Solid Concrete Ground Slab (Uninsulated Pre-1976)"
        return "Suspended Timber: Bare Boards over Cold Uninsulated Void"
    return "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth"

def map_u_to_door_spec(u_val: float) -> str:
    if u_val <= 0.05:
        return "No External Door (Internal Boundary Only)"
    best_spec = DOOR_SPECIFICATIONS[0]["label"]
    best_diff = 999.0
    for item in DOOR_SPECIFICATIONS:
        if item["u_value"] > 0:
            diff = abs(item["u_value"] - u_val)
            if diff < best_diff:
                best_diff = diff
                best_spec = item["label"]
    return best_spec

def infer_room_type(name: str, floor: str = "", temp: float = 20.0) -> str:
    n = name.lower()
    if "drawing" in n or "sitting" in n or "lounge" in n:
        return "Living Room / Sitting Room"
    if "living" in n:
        return "Living Room / Sitting Room"
    if "dining" in n:
        return "Dining Room"
    if "kitchen" in n or "breakfast" in n:
        return "Kitchen / Breakfast Room"
    if "snug" in n or "playroom" in n or "tv" in n:
        return "Playroom / Family Snug"
    if "office" in n or "study" in n or "library" in n:
        return "Home Office / Study"
    if "orangery" in n or "conservatory" in n or "garden room" in n or "sun room" in n:
        return "Orangery / Conservatory / Garden Room"
    if "utility" in n or "laundry" in n or "boot" in n or "dog" in n or "scullery" in n:
        return "Utility Room / Laundry"
    if "plant" in n or "pantry" in n or "unheated" in n or "store" in n or "garage" in n or "cellar" in n:
        return "Unheated Space / Plant Room / Store"
    if "en-suite" in n or "ensuite" in n or "bath" in n or "shower" in n:
        return "Bathroom / Shower Room / En-suite"
    if "cloakroom" in n or "downstairs wc" in n or "wc" in n:
        return "Cloakroom / Downstairs WC"
    if "porch" in n or "lobby" in n or "vestibule" in n:
        return "Side Entrance Lobby / Porch"
    if "dressing" in n or "wardrobe" in n:
        return "Dressing Room / Walk-in Wardrobe"
    if "hall" in n or "corridor" in n or "landing" in n or "stair" in n or "circulation" in n:
        return "Hallway / Stairs / Circulation"
    if "bed" in n or "guest" in n:
        return "Bedroom (General / Master)"
    
    if temp >= 22.0:
        return "Bathroom / Shower Room / En-suite"
    elif temp >= 21.0:
        return "Living Room / Sitting Room"
    return "Bedroom (General / Master)"

def extract_existing_rooms(ws: gspread.Worksheet) -> List[Dict[str, Any]]:
    """
    Reads existing user inputs from 2_Room_Heat_Loss worksheet non-destructively.
    Supports 45-col, 44-col, 42-col, 39-col, and 38-col layouts seamlessly.
    """
    try:
        raw_values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except Exception:
        return []

    if len(raw_values) < 5:
        return []

    has_door_spec = False
    has_floor_spec = False
    has_wall_spec = False
    has_room_type_col = False
    if len(raw_values) > 3:
        h_row = [str(c).strip().lower() for c in raw_values[3]]
        for h in h_row:
            if "room type" in h:
                has_room_type_col = True
            if "door spec" in h or "door area" in h:
                has_door_spec = True
            if "floor spec" in h:
                has_floor_spec = True
            if "wall spec" in h:
                has_wall_spec = True

    csv_fallback: Dict[str, Dict[str, str]] = {}
    csv_path = Path(__file__).resolve().parent.parent.parent / "room_by_room_heat_loss_survey.csv"
    if csv_path.exists():
        try:
            with open(csv_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    c = r.get("Room Code", "").strip()
                    if c:
                        csv_fallback[c] = r
        except Exception:
            pass

    rooms = []
    for row in raw_values[4:]:
        if not row:
            continue
        first_col = str(row[0]).strip() if len(row) > 0 else ""
        second_col = str(row[1]).strip() if len(row) > 1 else ""

        if first_col.lower().startswith("total") or "whole house" in first_col.lower():
            break
        if first_col == "" and second_col == "":
            break

        csv_r = csv_fallback.get(first_col, {})
        floor_level = safe_str(row[2] if len(row) > 2 else "", csv_r.get("Floor Level", "Ground Floor"))
        zone = safe_str(row[3] if len(row) > 3 else "", csv_r.get("Zone / Wing", "Old House"))
        name = safe_str(row[1] if len(row) > 1 else "", csv_r.get("Room Name", "Room"))

        # Determine layout (45-col vs 44-col legacy vs older)
        is_45_col = has_room_type_col or (len(row) > 4 and safe_str(row[4], "") in VALID_ROOM_TYPE)

        if is_45_col:
            room_type = safe_str(row[RoomCol.idx("ROOM_TYPE")], "")
            if not room_type or room_type not in VALID_ROOM_TYPE:
                room_type = csv_r.get("Room Type", "") or infer_room_type(name, floor_level, 20.0)
            temp = safe_float(row[RoomCol.idx("TI")] if len(row) > RoomCol.idx("TI") else 20.0, 20.0)
            len_val = safe_float(row[RoomCol.idx("LENGTH")] if len(row) > RoomCol.idx("LENGTH") else 0.0, 0.0)
            wid_val = safe_float(row[RoomCol.idx("WIDTH")] if len(row) > RoomCol.idx("WIDTH") else 0.0, 0.0)
            ht_val = safe_float(row[RoomCol.idx("HEIGHT")] if len(row) > RoomCol.idx("HEIGHT") else 2.6, 2.6)
            ext_wall = safe_float(row[RoomCol.idx("EXT_WALL_L")] if len(row) > RoomCol.idx("EXT_WALL_L") else 0.0, 0.0)
            wall_spec = safe_str(row[RoomCol.idx("WALL_SPEC")] if len(row) > RoomCol.idx("WALL_SPEC") else "", "")
            u_wall = safe_float(row[RoomCol.idx("U_WALL")] if len(row) > RoomCol.idx("U_WALL") else 1.4, 1.4)
            win_area = safe_float(row[RoomCol.idx("WIN_AREA")] if len(row) > RoomCol.idx("WIN_AREA") else 0.0, 0.0)
            win_spec = safe_str(row[RoomCol.idx("WIN_SPEC")] if len(row) > RoomCol.idx("WIN_SPEC") else "", "")
            door_area = safe_float(row[RoomCol.idx("DOOR_AREA")] if len(row) > RoomCol.idx("DOOR_AREA") else 0.0, 0.0)
            door_spec = safe_str(row[RoomCol.idx("DOOR_SPEC")] if len(row) > RoomCol.idx("DOOR_SPEC") else "", "")
            u_door = safe_float(row[RoomCol.idx("U_DOOR")] if len(row) > RoomCol.idx("U_DOOR") else 0.0, 0.0)
            fl_area = safe_float(row[RoomCol.idx("FL_AREA")] if len(row) > RoomCol.idx("FL_AREA") else 0.0, 0.0)
            floor_spec = safe_str(row[RoomCol.idx("FLOOR_SPEC")] if len(row) > RoomCol.idx("FLOOR_SPEC") else "", "")
            u_fl = safe_float(row[RoomCol.idx("U_FLOOR")] if len(row) > RoomCol.idx("U_FLOOR") else 0.8, 0.8)
            roof_area = safe_float(row[RoomCol.idx("ROOF_AREA")] if len(row) > RoomCol.idx("ROOF_AREA") else 0.0, 0.0)
            ceil_spec = safe_str(row[RoomCol.idx("CEIL_SPEC")] if len(row) > RoomCol.idx("CEIL_SPEC") else "", "")
            chimney = safe_str(row[RoomCol.idx("CHIMNEY")] if len(row) > RoomCol.idx("CHIMNEY") else "", "")
            mech_vent = safe_str(row[RoomCol.idx("MECH_VENT")] if len(row) > RoomCol.idx("MECH_VENT") else "", "")
            emitter_type = safe_str(row[RoomCol.idx("EMITTER_TYPE")] if len(row) > RoomCol.idx("EMITTER_TYPE") else "", "")
            raw_notes = safe_str(row[RoomCol.idx("NOTES")] if len(row) > RoomCol.idx("NOTES") else "", "")

        elif has_door_spec:
            # 44-col legacy
            temp = safe_float(row[4] if len(row) > 4 else 20.0, 20.0)
            room_type = csv_r.get("Room Type", "")
            if not room_type or room_type not in VALID_ROOM_TYPE:
                room_type = infer_room_type(name, floor_level, temp)
            len_val = safe_float(row[5] if len(row) > 5 else 0.0, 0.0)
            wid_val = safe_float(row[6] if len(row) > 6 else 0.0, 0.0)
            ht_val = safe_float(row[8] if len(row) > 8 else 2.6, 2.6)
            ext_wall = safe_float(row[10] if len(row) > 10 else 0.0, 0.0)
            wall_spec = safe_str(row[11] if len(row) > 11 else "", "")
            u_wall = safe_float(row[12] if len(row) > 12 else 1.4, 1.4)
            win_area = safe_float(row[14] if len(row) > 14 else 0.0, 0.0)
            win_spec = safe_str(row[15] if len(row) > 15 else "", "")
            door_area = safe_float(row[18] if len(row) > 18 else 0.0, 0.0)
            door_spec = safe_str(row[19] if len(row) > 19 else "", "")
            u_door = safe_float(row[20] if len(row) > 20 else 0.0, 0.0)
            fl_area = safe_float(row[22] if len(row) > 22 else 0.0, 0.0)
            floor_spec = safe_str(row[23] if len(row) > 23 else "", "")
            u_fl = safe_float(row[24] if len(row) > 24 else 0.8, 0.8)
            roof_area = safe_float(row[26] if len(row) > 26 else 0.0, 0.0)
            ceil_spec = safe_str(row[27] if len(row) > 27 else "", "")
            chimney = safe_str(row[31] if len(row) > 31 else "", "")
            mech_vent = safe_str(row[32] if len(row) > 32 else "", "")
            emitter_type = safe_str(row[41] if len(row) > 41 else "", "")
            raw_notes = safe_str(row[43] if len(row) > 43 else "", "")

        elif has_floor_spec:
            # 39-col legacy
            temp = safe_float(row[4] if len(row) > 4 else 20.0, 20.0)
            room_type = infer_room_type(name, floor_level, temp)
            len_val = safe_float(row[5] if len(row) > 5 else 0.0, 0.0)
            wid_val = safe_float(row[6] if len(row) > 6 else 0.0, 0.0)
            ht_val = safe_float(row[8] if len(row) > 8 else 2.6, 2.6)
            ext_wall = safe_float(row[10] if len(row) > 10 else 0.0, 0.0)
            wall_spec = safe_str(row[11] if len(row) > 11 else "", "")
            u_wall = safe_float(row[12] if len(row) > 12 else 1.4, 1.4)
            win_area = safe_float(row[14] if len(row) > 14 else 0.0, 0.0)
            win_spec = safe_str(row[15] if len(row) > 15 else "", "")
            door_area = 0.0
            door_spec = "No External Door (Internal Boundary Only)"
            u_door = 0.0
            fl_area = safe_float(row[18] if len(row) > 18 else 0.0, 0.0)
            floor_spec = safe_str(row[19] if len(row) > 19 else "", "")
            u_fl = safe_float(row[20] if len(row) > 20 else 0.8, 0.8)
            roof_area = safe_float(row[22] if len(row) > 22 else 0.0, 0.0)
            ceil_spec = safe_str(row[23] if len(row) > 23 else "", "")
            chimney = safe_str(row[27] if len(row) > 27 else "", "")
            mech_vent = "None (Natural Infiltration Only)"
            emitter_type = "Type 22 (Double Convector)"
            raw_notes = safe_str(row[38] if len(row) > 38 else "", "")
        else:
            # 38 or older legacy
            temp = safe_float(row[4] if len(row) > 4 else 20.0, 20.0)
            room_type = infer_room_type(name, floor_level, temp)
            len_val = safe_float(row[5] if len(row) > 5 else 0.0, 0.0)
            wid_val = safe_float(row[6] if len(row) > 6 else 0.0, 0.0)
            ht_val = safe_float(row[8] if len(row) > 8 else 2.6, 2.6)
            ext_wall = safe_float(row[10] if len(row) > 10 else 0.0, 0.0)
            wall_spec = safe_str(row[11] if len(row) > 11 else "", "")
            u_wall = safe_float(row[12] if len(row) > 12 else 1.4, 1.4)
            win_area = safe_float(row[14] if len(row) > 14 else 0.0, 0.0)
            win_spec = safe_str(row[15] if len(row) > 15 else "", "")
            door_area = 0.0
            door_spec = "No External Door (Internal Boundary Only)"
            u_door = 0.0
            fl_area = safe_float(row[18] if len(row) > 18 else 0.0, 0.0)
            floor_spec = ""
            u_fl = safe_float(row[19] if len(row) > 19 else 0.8, 0.8)
            roof_area = safe_float(row[21] if len(row) > 21 else 0.0, 0.0)
            ceil_spec = safe_str(row[22] if len(row) > 22 else "", "")
            chimney = "No Chimney / Permanently Sealed"
            mech_vent = "None (Natural Infiltration Only)"
            emitter_type = "Type 22 (Double Convector)"
            raw_notes = safe_str(row[-1] if len(row) > 30 else "", "")

        # Fallbacks & spec mapping
        if not wall_spec or wall_spec.startswith("#") or wall_spec.startswith("="):
            wall_spec = csv_r.get("Wall Specification", "") or map_u_to_wall_spec(u_wall, zone)
        if not win_spec or win_spec.startswith("#") or win_spec.startswith("="):
            win_spec = csv_r.get("Window Specification", "Single Glazed (Historic Timber Sash / Casement)")
        if not door_spec or door_spec.startswith("#") or door_spec.startswith("="):
            door_spec = csv_r.get("Door Specification", "No External Door (Internal Boundary Only)")
        if not floor_spec or floor_spec.startswith("#") or floor_spec.startswith("="):
            floor_spec = csv_r.get("Floor Specification", "") or map_u_to_floor_spec(u_fl, floor_level, zone)
        if not ceil_spec or ceil_spec.startswith("#") or ceil_spec.startswith("="):
            ceil_spec = csv_r.get("Ceiling Specification", "") or ("Intermediate Floor (Heated Space Above)" if "ground" in floor_level.lower() else "Modern Building Regs Loft: 270-300mm (Mineral Wool)")
        if chimney not in VALID_CHIMNEY:
            csv_ch = csv_r.get("Chimney / Fireplace", csv_r.get("Chimney Flue", ""))
            chimney = csv_ch if csv_ch in VALID_CHIMNEY else "No Chimney / Permanently Sealed"
        if mech_vent not in VALID_MECH:
            csv_mv = csv_r.get("Mechanical Ventilation", "")
            if csv_mv in VALID_MECH:
                mech_vent = csv_mv
            else:
                n_lo = name.lower()
                if "bath" in n_lo or "wc" in n_lo or "shower" in n_lo or "en-suite" in n_lo:
                    mech_vent = "Intermittent Extract Fan (Bathroom / WC - 15 l/s)"
                elif "kitchen" in n_lo or "utility" in n_lo:
                    mech_vent = "Intermittent Extract Fan (Kitchen / Utility - 30-60 l/s)"
                else:
                    mech_vent = "None (Natural Infiltration Only)"
        if emitter_type not in VALID_EMITTER:
            csv_et = csv_r.get("Planned Emitter Type", "")
            emitter_type = csv_et if csv_et in VALID_EMITTER else "Type 22 (Double Convector)"

        if raw_notes and not raw_notes.startswith("=") and not raw_notes.isdigit():
            notes = raw_notes
        else:
            notes = csv_r.get("Notes", "") or csv_r.get("Notes / Survey Observations", "")

        room = {
            "code": first_col,
            "name": name,
            "floor": floor_level,
            "zone": zone,
            "room_type": room_type,
            "temp": temp,
            "len": len_val,
            "wid": wid_val,
            "ht": ht_val,
            "ext_wall": ext_wall,
            "wall_spec": wall_spec,
            "u_wall": u_wall,
            "win_area": win_area,
            "win_spec": win_spec,
            "door_area": door_area,
            "door_spec": door_spec,
            "u_door": u_door,
            "fl_area": fl_area,
            "floor_spec": floor_spec,
            "u_fl": u_fl,
            "roof_area": roof_area,
            "ceil_spec": ceil_spec,
            "chimney": chimney,
            "q_chimney": chimney,
            "mech_vent": mech_vent,
            "emitter_type": emitter_type,
            "notes": notes
        }
        rooms.append(room)

    return rooms

def update_named_ranges(ss: gspread.Spreadsheet, ws: gspread.Worksheet, total_row: int, num_rooms: int):
    named_ranges_spec = {
        "Room_Total_Loss_W": (total_row - 1, total_row, RoomCol.TOTAL_LOSS_IDX, RoomCol.TOTAL_LOSS_IDX + 1),
        "Room_Total_Area_m2": (total_row - 1, total_row, RoomCol.AREA_IDX, RoomCol.AREA_IDX + 1),
        "Room_Average_Ti": (total_row - 1, total_row, RoomCol.TI_IDX, RoomCol.TI_IDX + 1),
        "Room_Whole_House_Intensity": (total_row - 1, total_row, RoomCol.INTENSITY_IDX, RoomCol.INTENSITY_IDX + 1),
    }
    try:
        existing = {nr["name"]: nr for nr in ss.list_named_ranges()}
        requests = []
        for name in named_ranges_spec:
            if name in existing:
                requests.append({"deleteNamedRange": {"namedRangeId": existing[name]["namedRangeId"]}})
        for name, (sr, er, sc, ec) in named_ranges_spec.items():
            requests.append({
                "addNamedRange": {
                    "namedRange": {
                        "name": name,
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": sr,
                            "endRowIndex": er,
                            "startColumnIndex": sc,
                            "endColumnIndex": ec
                        }
                    }
                }
            })
        if requests:
            ss.batch_update({"requests": requests})
    except Exception as e:
        print(f"Note on updating named ranges: {e}")

def build_room_tab(ss: gspread.Spreadsheet) -> Tuple[gspread.Worksheet, List[Dict[str, Any]], int, int]:
    """
    Builds and formats the 2_Room_Heat_Loss master schedule worksheet.
    Preserves all user-edited non-formulaic cells from the live sheet before updating.
    Returns (worksheet, formatting_requests, num_rooms, total_row).
    """
    tab_name = "2_Room_Heat_Loss"
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab_name, rows=45, cols=RoomCol.total_cols)

    existing_rooms = extract_existing_rooms(ws)
    active_rooms = existing_rooms if len(existing_rooms) > 0 else DEFAULT_ROOMS

    num_rooms = len(active_rooms)
    last_room_row = 4 + num_rooms
    total_row = 5 + num_rooms
    total_rows = total_row + 2
    total_cols = RoomCol.total_cols

    if ws.row_count != total_rows or ws.col_count != total_cols:
        ws.resize(rows=total_rows, cols=total_cols)

    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1: Title Banner
    grid[0][0] = "2_Room_Heat_Loss: Room-by-Room Heat Loss Assessment & Low-Flow Radiator Sizing"

    # Row 2: MCS MIS 3005-D & BS EN 12831 Compliance Audit Header Block (5 Badges)
    grid[1][0] = "COMPLIANCE: BS EN 12831-1:2017 & MCS MIS 3005-D | Regime: Continuous 24/7 (0% Reheat Boost)"
    grid[1][6] = "CLIMATE: Outside Te = -4.0°C ('1_Inputs'!$C$5) | Sub-Floor Ground Tg = 10.0°C ('1_Inputs'!$C$6)"
    grid[1][11] = "HEAT PUMP FLOW: Design Flow = 45.0°C ('1_Inputs'!$C$175) | Return = 40.0°C ('1_Inputs'!$C$176) | System ΔT = 5.0 K ('1_Inputs'!$C$177)"
    grid[1][23] = "THERMAL BRIDGING: Allowance = +10.0% ('1_Inputs'!$C$178) applied to all fabric transmission losses"
    grid[1][32] = "EMITTER DESIGN: Mean Water ΔTm = 22.5 K | Exponent n = 1.30 ('1_Inputs'!$C$179) | Flow Rate = W / (1.163 × ΔT)"

    # Row 3: Category Group Headers
    grid[2][0] = "ROOM IDENTIFICATION"
    grid[2][6] = "ROOM GEOMETRY"
    grid[2][11] = "FABRIC TRANSMISSION LOSSES (W)"
    grid[2][32] = "INFILTRATION & VENTILATION (BS EN 12831)"
    grid[2][36] = "MCS 45°C EMITTER SIZING & FLOW RATES"
    grid[2][44] = "SURVEY NOTES"

    # Row 4: Column Headers from canonical schema
    for c_i, h in enumerate(RoomCol.headers):
        grid[3][c_i] = h

    # Populate active rooms with live formulas
    for r_idx, rm in enumerate(active_rooms):
        r = 5 + r_idx
        wall_spec = rm.get("wall_spec", map_u_to_wall_spec(rm.get("u_wall", 1.40), rm.get("zone", "")))
        floor_spec = rm.get("floor_spec", map_u_to_floor_spec(rm.get("u_fl", 0.80), rm.get("floor", "Ground Floor"), rm.get("zone", "")))
        door_spec = rm.get("door_spec", map_u_to_door_spec(rm.get("u_door", 0.00)))
        room_type = rm.get("room_type", infer_room_type(rm["name"], rm.get("floor", ""), rm.get("temp", 20.0)))

        # 6-lookup ACH formula (Walls, Windows, Floors, Ceilings, Chimneys, Mechanical Extract)
        ach_formula = (
            f"=VLOOKUP(M{r}, '1_Inputs'!$B$129:$D$145, 3, FALSE) + "
            f"VLOOKUP(Q{r}, '1_Inputs'!$B$98:$D$109, 3, FALSE) + "
            f"VLOOKUP(Y{r}, '1_Inputs'!$B$148:$D$163, 3, FALSE) + "
            f"VLOOKUP(AC{r}, '1_Inputs'!$B$113:$D$124, 3, FALSE) + "
            f"VLOOKUP(AG{r}, '1_Inputs'!$B$76:$C$79, 2, FALSE) + "
            f"VLOOKUP(AH{r}, '1_Inputs'!$B$81:$C$85, 2, FALSE)"
        )

        row_arr = [
            rm["code"],                                       # Col A (0)
            rm["name"],                                       # Col B (1)
            rm["floor"],                                      # Col C (2)
            rm["zone"],                                       # Col D (3)
            room_type,                                        # Col E (4) (Room Type Dropdown)
            f"=VLOOKUP(E{r}, '1_Inputs'!$B$185:$C$198, 2, FALSE)", # Col F (5) (Design Ti Formula)
            str(rm["len"]),                                   # Col G (6) (Length)
            str(rm["wid"]),                                   # Col H (7) (Width)
            f"=G{r}*H{r}",                                    # Col I (8) (Floor Area)
            str(rm["ht"]),                                    # Col J (9) (Height)
            f"=I{r}*J{r}",                                    # Col K (10) (Volume)
            str(rm["ext_wall"]),                              # Col L (11) (Ext Wall length)
            wall_spec,                                        # Col M (12) (Wall Specification Dropdown)
            f"=VLOOKUP(M{r}, '1_Inputs'!$B$129:$D$145, 2, FALSE)", # Col N (13) (U Wall Formula)
            f"=MAX(0, (L{r}*J{r}-P{r}-T{r})*N{r}*(F{r}-'1_Inputs'!$C$5))", # Col O (14) (Net Wall Loss W)
            str(rm["win_area"]),                              # Col P (15) (Window Area m²)
            rm["win_spec"],                                   # Col Q (16) (Window Specification Dropdown)
            f"=VLOOKUP(Q{r}, '1_Inputs'!$B$98:$D$109, 2, FALSE)", # Col R (17) (U Window Formula)
            f"=P{r}*R{r}*(F{r}-'1_Inputs'!$C$5)",             # Col S (18) (Window Loss W)
            str(rm.get("door_area", 0.0)),                    # Col T (19) (Door Area m²)
            door_spec,                                        # Col U (20) (Door Specification Dropdown)
            f"=VLOOKUP(U{r}, '1_Inputs'!$B$165:$C$171, 2, FALSE)", # Col V (21) (U Door Formula)
            f"=T{r}*V{r}*(F{r}-'1_Inputs'!$C$5)",             # Col W (22) (Door Loss W)
            f"=G{r}*H{r}",                                    # Col X (23) (Exposed Floor Area = Length * Width)
            floor_spec,                                       # Col Y (24) (Floor Specification Dropdown)
            f"=VLOOKUP(Y{r}, '1_Inputs'!$B$148:$D$163, 2, FALSE)", # Col Z (25) (U Floor Formula)
            f"=X{r}*Z{r}*(F{r}-'1_Inputs'!$C$6)",             # Col AA (26) (Floor Loss with Ground Temp C6)
            f"=G{r}*H{r}",                                    # Col AB (27) (Ceiling Area = Length * Width)
            rm.get("ceil_spec", "Intermediate Floor (Heated Space Above)"), # Col AC (28) (Ceiling Spec Dropdown)
            f"=VLOOKUP(AC{r}, '1_Inputs'!$B$113:$D$124, 2, FALSE)", # Col AD (29) (U Ceiling Formula)
            f"=AB{r}*AD{r}*(F{r}-'1_Inputs'!$C$5)",           # Col AE (30) (Ceiling Loss W)
            f"=(O{r}+S{r}+W{r}+AA{r}+AE{r})*'1_Inputs'!$C$178", # Col AF (31) (Thermal Bridging Loss W)
            rm.get("chimney", rm.get("q_chimney", "No Chimney / Permanently Sealed")), # Col AG (32) (Chimney Dropdown)
            rm.get("mech_vent", "None (Natural Infiltration Only)"), # Col AH (33) (Mechanical Ventilation Dropdown)
            ach_formula,                                      # Col AI (34) (Calculated ACH Formula)
            f"='1_Inputs'!$C$11*AI{r}*K{r}*(F{r}-'1_Inputs'!$C$5)", # Col AJ (35) (Vent Loss W)
            f"=SUM(O{r}, S{r}, W{r}, AA{r}, AE{r}, AF{r}, AJ{r})",   # Col AK (36) (Total Room Loss W)
            f"=AK{r}/I{r}",                                   # Col AL (37) (Intensity W/m²)
            f"=AK{r}",                                        # Col AM (38) (Req Emitter Output at 45°C Flow W)
            f"=ROUND(AK{r}/((MAX(5, (('1_Inputs'!$C$175+'1_Inputs'!$C$176)/2-F{r}))/50)^'1_Inputs'!$C$179), 0)", # Col AN (39) (Req Rad Catalogue at ΔT50 W)
            f"=ROUND(AK{r}/(1.163*'1_Inputs'!$C$177), 1)",    # Col AO (40) (Design Water Flow Rate l/h)
            f'=IF(AO{r}>450, "22mm copper", IF(AO{r}>150, "15mm copper", "10mm or 15mm"))', # Col AP (41) (Min Pipe Size)
            rm.get("emitter_type", "Type 22 (Double Convector)"), # Col AQ (42) (Planned Emitter Type Dropdown)
            f'=IF(AK{r}>2000, "Type 33 (e.g. 600x1600) or 2x Type 22", IF(AK{r}>1200, "Type 22 (e.g. 600x1200)", IF(AK{r}>600, "Type 22 (e.g. 600x800)", "Type 21 / Type 11 (e.g. 500x600)")))', # Col AR (43) (Recommended Sizing)
            rm.get("notes", "")                               # Col AS (44) (Survey Observations)
        ]
        grid[4 + r_idx] = row_arr

    # Summary Total Row
    grid[total_row - 1] = [
        "Total Whole House",
        f"{num_rooms} Assessed Rooms",
        "Ground & First",
        "All Wings",
        "-",
        f"=SUMPRODUCT(F5:F{last_room_row}, I5:I{last_room_row})/SUM(I5:I{last_room_row})",
        "-",
        "-",
        f"=SUM(I5:I{last_room_row})",
        "-",
        f"=SUM(K5:K{last_room_row})",
        "-",
        "-",
        "-",
        f"=SUM(O5:O{last_room_row})",
        f"=SUM(P5:P{last_room_row})",
        "-",
        "-",
        f"=SUM(S5:S{last_room_row})",
        f"=SUM(T5:T{last_room_row})",
        "-",
        "-",
        f"=SUM(W5:W{last_room_row})",
        f"=SUM(X5:X{last_room_row})",
        "-",
        "-",
        f"=SUM(AA5:AA{last_room_row})",
        f"=SUM(AB5:AB{last_room_row})",
        "-",
        "-",
        f"=SUM(AE5:AE{last_room_row})",
        f"=SUM(AF5:AF{last_room_row})",
        "-",
        "-",
        f"=AVERAGE(AI5:AI{last_room_row})",
        f"=SUM(AJ5:AJ{last_room_row})",
        f"=SUM(AK5:AK{last_room_row})",
        f"=AK{total_row}/I{total_row}",
        f"=SUM(AM5:AM{last_room_row})",
        f"=SUM(AN5:AN{last_room_row})",
        f"=SUM(AO5:AO{last_room_row})",
        "Distribution Header",
        "Whole House Emitters",
        "Whole House Sizing",
        "Survey & Sizing Complete"
    ]

    # Unmerge header rows first to allow multi-cell badge text writing
    try:
        ss.batch_update({"requests": [create_unmerge_cells_request(ws.id, max_rows=5, max_cols=total_cols)]})
    except Exception:
        pass

    # Write values into worksheet
    ws.update(values=grid, range_name=f"A1:AS{total_rows}", value_input_option="USER_ENTERED")

    # Formatting requests
    fmt_reqs: List[Dict[str, Any]] = []

    # Freeze top 4 rows
    fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=4, frozen_cols=0))

    # Column widths from schema
    for c_i, w in enumerate(RoomCol.widths):
        fmt_reqs.append(create_set_column_width_request(ws.id, c_i, c_i + 1, w))

    # Row 1: Title Banner
    fmt_reqs.append(create_merge_cells_request(ws.id, 0, 1, 0, total_cols))
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 0, 1, 0, total_cols,
        bg_color=THEME["PRIMARY_HEADER_BG"],
        font_color=THEME["HEADER_TEXT"],
        bold=True,
        font_size=12,
        align="LEFT"
    ))

    # Row 2: MCS Compliance Audit Header Block (5 Badges)
    audit_badges = [
        (0, 6, {"red": 0.10, "green": 0.25, "blue": 0.35}),
        (6, 11, {"red": 0.12, "green": 0.30, "blue": 0.40}),
        (11, 23, {"red": 0.15, "green": 0.35, "blue": 0.45}),
        (23, 32, {"red": 0.12, "green": 0.30, "blue": 0.40}),
        (32, total_cols, {"red": 0.10, "green": 0.25, "blue": 0.35})
    ]
    for start_c, end_c, bg in audit_badges:
        fmt_reqs.append(create_merge_cells_request(ws.id, 1, 2, start_c, end_c))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, 1, 2, start_c, end_c,
            bg_color=bg,
            font_color={"red": 0.95, "green": 0.98, "blue": 1.0},
            bold=True,
            font_size=9,
            align="CENTER"
        ))

    # Row 3: Category Group Headers
    group_spans = [
        (0, RoomCol.idx("LENGTH"), THEME["SECONDARY_HEADER_BG"]),
        (RoomCol.idx("LENGTH"), RoomCol.idx("EXT_WALL_L"), THEME["SECTION_HEADER_BG"]),
        (RoomCol.idx("EXT_WALL_L"), RoomCol.idx("CHIMNEY"), {"red": 0.15, "green": 0.35, "blue": 0.45}),
        (RoomCol.idx("CHIMNEY"), RoomCol.idx("TOTAL_LOSS"), {"red": 0.10, "green": 0.32, "blue": 0.42}),
        (RoomCol.idx("TOTAL_LOSS"), RoomCol.idx("REC_EMITTER") + 1, THEME["PRIMARY_HEADER_BG"]),
        (RoomCol.idx("REC_EMITTER") + 1, RoomCol.total_cols, THEME["CARD_HEADER_BG"])
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

    # Row 4: Column Headers
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, 3, 4, 0, total_cols,
        bg_color=THEME["CARD_HEADER_BG"],
        font_color=THEME["HEADER_TEXT"],
        bold=True,
        font_size=9,
        align="CENTER",
        wrap_strategy="WRAP"
    ))

    # Data Rows Formatting
    for r_i in range(4, last_room_row):
        bg = THEME["ZEBRA_BG"] if r_i % 2 == 1 else {"red": 1.0, "green": 1.0, "blue": 1.0}
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, total_cols,
            bg_color=bg,
            font_color=THEME["DARK_TEXT"],
            font_size=9,
            align="RIGHT"
        ))
        # Left align Room Code, Name, Floor, Zone
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, 0, 4, align="LEFT", font_size=9))

        # Room Type Dropdown in soft blue
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("ROOM_TYPE"), RoomCol.idx("ROOM_TYPE") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))

        # Dropdowns in soft blue:
        # Wall Spec Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("WALL_SPEC"), RoomCol.idx("WALL_SPEC") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Window Spec Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("WIN_SPEC"), RoomCol.idx("WIN_SPEC") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Door Spec Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("DOOR_SPEC"), RoomCol.idx("DOOR_SPEC") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Floor Spec Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("FLOOR_SPEC"), RoomCol.idx("FLOOR_SPEC") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Ceiling Spec Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("CEIL_SPEC"), RoomCol.idx("CEIL_SPEC") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Chimney Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("CHIMNEY"), RoomCol.idx("CHIMNEY") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Mechanical Ventilation Dropdown
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("MECH_VENT"), RoomCol.idx("MECH_VENT") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))

        # Calculated ACH in soft peach/salmon
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("ACH"), RoomCol.idx("ACH") + 1, bg_color=THEME["CALC_BG"], bold=True, align="RIGHT", font_size=10, number_format=FORMATS["DECIMAL_2"]))

        # Pipe size centered and bold
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("REC_PIPE"), RoomCol.idx("REC_PIPE") + 1, align="CENTER", bold=True, font_size=9))
        # Planned Emitter Type Dropdown in soft blue
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("EMITTER_TYPE"), RoomCol.idx("EMITTER_TYPE") + 1, bg_color=THEME["INPUT_BG"], font_color={"red": 0.05, "green": 0.20, "blue": 0.45}, align="LEFT", font_size=9))
        # Rec Emitter & Notes left aligned
        fmt_reqs.append(create_repeat_cell_request(ws.id, r_i, r_i + 1, RoomCol.idx("REC_EMITTER"), total_cols, align="LEFT", font_size=9, font_color=THEME["DARK_TEXT"]))

    # Number Formats across data rows
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("TI"), RoomCol.idx("TI") + 1, number_format=FORMATS["TEMP_C"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("LENGTH"), RoomCol.idx("AREA"), number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("AREA"), RoomCol.idx("AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("HEIGHT"), RoomCol.idx("HEIGHT") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("VOLUME"), RoomCol.idx("VOLUME") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("EXT_WALL_L"), RoomCol.idx("EXT_WALL_L") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("U_WALL"), RoomCol.idx("U_WALL") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("WALL_LOSS"), RoomCol.idx("WALL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("WIN_AREA"), RoomCol.idx("WIN_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("U_WIN"), RoomCol.idx("U_WIN") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("WIN_LOSS"), RoomCol.idx("WIN_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("DOOR_AREA"), RoomCol.idx("DOOR_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("U_DOOR"), RoomCol.idx("U_DOOR") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("DOOR_LOSS"), RoomCol.idx("DOOR_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("FL_AREA"), RoomCol.idx("FL_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("U_FLOOR"), RoomCol.idx("U_FLOOR") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("FLOOR_LOSS"), RoomCol.idx("FLOOR_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("ROOF_AREA"), RoomCol.idx("ROOF_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("U_CEIL"), RoomCol.idx("U_CEIL") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("CEIL_LOSS"), RoomCol.idx("CEIL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("TB_LOSS"), RoomCol.idx("TB_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("VENT_LOSS"), RoomCol.idx("VENT_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("TOTAL_LOSS"), RoomCol.idx("TOTAL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("INTENSITY"), RoomCol.idx("INTENSITY") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("RAD_45"), RoomCol.idx("RAD_45") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("RAD_DT50"), RoomCol.idx("RAD_DT50") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, RoomCol.idx("FLOW_RATE"), RoomCol.idx("FLOW_RATE") + 1, number_format=FORMATS["DECIMAL_1"]))

    # Total Row Formatting
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, last_room_row, total_row, 0, total_cols,
        bg_color=THEME["TOTAL_BG"],
        font_color=THEME["DARK_TEXT"],
        bold=True,
        font_size=10,
        align="RIGHT"
    ))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 0, 4, align="LEFT"))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("AREA"), RoomCol.idx("AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("VOLUME"), RoomCol.idx("VOLUME") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("WALL_LOSS"), RoomCol.idx("WALL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("WIN_AREA"), RoomCol.idx("WIN_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("WIN_LOSS"), RoomCol.idx("WIN_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("DOOR_AREA"), RoomCol.idx("DOOR_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("DOOR_LOSS"), RoomCol.idx("DOOR_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("FL_AREA"), RoomCol.idx("FL_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("FLOOR_LOSS"), RoomCol.idx("FLOOR_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("ROOF_AREA"), RoomCol.idx("ROOF_AREA") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("CEIL_LOSS"), RoomCol.idx("CEIL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("TB_LOSS"), RoomCol.idx("TB_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("ACH"), RoomCol.idx("ACH") + 1, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("VENT_LOSS"), RoomCol.idx("VENT_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("TOTAL_LOSS"), RoomCol.idx("TOTAL_LOSS") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("INTENSITY"), RoomCol.idx("INTENSITY") + 1, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("RAD_45"), RoomCol.idx("RAD_45") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("RAD_DT50"), RoomCol.idx("RAD_DT50") + 1, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, RoomCol.idx("FLOW_RATE"), RoomCol.idx("FLOW_RATE") + 1, number_format=FORMATS["DECIMAL_1"]))

    # Accent highlight on Total Room Heat Loss
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, last_room_row, total_row, RoomCol.idx("TOTAL_LOSS"), RoomCol.idx("TOTAL_LOSS") + 1,
        bg_color=THEME["ACCENT_BG"],
        font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
        bold=True,
        font_size=11,
        number_format=FORMATS["INTEGER"]
    ))

    # Data Validations
    wall_labels = [s["label"] for s in WALL_SPECIFICATIONS]
    win_labels = [s["label"] for s in WINDOW_SPECIFICATIONS]
    door_labels = [s["label"] for s in DOOR_SPECIFICATIONS]
    floor_labels = [s["label"] for s in FLOOR_SPECIFICATIONS]
    ceil_labels = [s["label"] for s in CEILING_SPECIFICATIONS]
    chimney_labels = [opt["label"] for opt in CHIMNEY_SPECIFICATIONS]
    mech_labels = [opt["label"] for opt in MECHANICAL_VENTILATION_SPECIFICATIONS]
    emitter_labels = [opt["label"] for opt in EMITTER_SPECIFICATIONS]
    room_type_labels = [s["label"] for s in ROOM_TYPE_SPECIFICATIONS]
    # Clear any residual data validations across data cells
    fmt_reqs.append({
        "setDataValidation": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 4,
                "endRowIndex": last_room_row,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols
            }
        }
    })

    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("ROOM_TYPE"), RoomCol.idx("ROOM_TYPE") + 1, room_type_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("WALL_SPEC"), RoomCol.idx("WALL_SPEC") + 1, wall_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("WIN_SPEC"), RoomCol.idx("WIN_SPEC") + 1, win_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("DOOR_SPEC"), RoomCol.idx("DOOR_SPEC") + 1, door_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("FLOOR_SPEC"), RoomCol.idx("FLOOR_SPEC") + 1, floor_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("CEIL_SPEC"), RoomCol.idx("CEIL_SPEC") + 1, ceil_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("CHIMNEY"), RoomCol.idx("CHIMNEY") + 1, chimney_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("MECH_VENT"), RoomCol.idx("MECH_VENT") + 1, mech_labels))
    fmt_reqs.append(create_data_validation_request(ws.id, 4, last_room_row, RoomCol.idx("EMITTER_TYPE"), RoomCol.idx("EMITTER_TYPE") + 1, emitter_labels))

    # Register Named Ranges
    update_named_ranges(ss, ws, total_row, num_rooms)

    return ws, fmt_reqs, num_rooms, total_row
