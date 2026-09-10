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
from ..config import THEME, FORMATS, INFILTRATION_QUESTIONNAIRE, WINDOW_SPECIFICATIONS, CEILING_SPECIFICATIONS, WALL_SPECIFICATIONS, FLOOR_SPECIFICATIONS
from ..schema import RoomCol
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
        "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 5.60, "ext_wall": 10.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 6.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 28.1, "floor_spec": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "u_fl": 0.80,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Double-height open stair volume connecting GF to FF. Check draught seal on front door."
    },
    {
        "code": "GF-02", "name": "Drawing Room / Main Living", "floor": "Ground Floor", "zone": "Georgian end",
        "temp": 21.0, "len": 8.50, "wid": 6.25, "ht": 2.80, "ext_wall": 14.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 9.2, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 53.1, "floor_spec": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "u_fl": 0.80,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "Open Fireplace (Unsealed Flue)",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Solid Georgian brick wall. Open fireplace needs chimney balloon or flue damper to cut 0.6 ACH."
    },
    {
        "code": "GF-03", "name": "Dining Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 6.00, "wid": 5.50, "ht": 2.80, "ext_wall": 11.50,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 5.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 33.0, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Check floor void beneath timber boards. Good candidate for underfloor insulation."
    },
    {
        "code": "GF-04", "name": "Kitchen & Breakfast Area", "floor": "Ground Floor", "zone": "New build",
        "temp": 18.0, "len": 7.50, "wid": 6.50, "ht": 2.80, "ext_wall": 14.00,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 8.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 48.8, "floor_spec": "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)", "u_fl": 0.15,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "New insulated cavity wall & insulated slab. High thermal and airtight performance."
    },
    {
        "code": "GF-05", "name": "Orangery / Garden Room", "floor": "Ground Floor", "zone": "Orangery",
        "temp": 20.0, "len": 10.00, "wid": 6.00, "ht": 2.80, "ext_wall": 22.00,
        "wall_spec": "Solid Wall + 100mm External Wall Insulation (EWI)", "u_wall": 0.25,
        "win_area": 35.0, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 60.0, "floor_spec": "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)", "u_fl": 0.15,
        "roof_area": 60.0, "ceil_spec": "Glazed Roof Lantern / Sloping Skylight (Orangery)", "u_roof": 1.50,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Large glazed roof lantern and double glazed perimeter. High heat loss density."
    },
    {
        "code": "GF-06", "name": "Snug / Family Sitting Room", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 21.0, "len": 6.00, "wid": 5.00, "ht": 2.80, "ext_wall": 11.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 4.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 30.0, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "Room-Sealed Stove with Ext Air",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Timber beam ceiling. Space available for Type 22 or Type 33 radiator."
    },
    {
        "code": "GF-07", "name": "Study / Home Office", "floor": "Ground Floor", "zone": "Thatched gable ended",
        "temp": 20.0, "len": 6.00, "wid": 4.00, "ht": 2.80, "ext_wall": 10.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 3.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 24.0, "floor_spec": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "u_fl": 0.60,
        "roof_area": 0.0, "ceil_spec": "Intermediate Floor (Heated Space Above)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Intermediate Floor (Heated Space Above)",
        "notes": "Continuous occupancy during workdays. Comfort requires 20°C."
    },
    {
        "code": "GF-08", "name": "Utility & Boot Room", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 18.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 7.50,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 2.2, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 14.0, "floor_spec": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "u_fl": 1.10,
        "roof_area": 14.0, "ceil_spec": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Solid uninsulated wall & sloping roof. Prime candidate for roof & wall insulation."
    },
    {
        "code": "GF-09", "name": "Ground Floor Cloakroom / WC", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 18.0, "len": 2.50, "wid": 4.00, "ht": 2.60, "ext_wall": 2.50,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 1.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 10.0, "floor_spec": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "u_fl": 1.10,
        "roof_area": 10.0, "ceil_spec": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Small heated towel rail or compact radiator needed."
    },
    {
        "code": "GF-10", "name": "Plant Room / Back Scullery", "floor": "Ground Floor", "zone": "Lean to North",
        "temp": 16.0, "len": 6.00, "wid": 3.00, "ht": 2.60, "ext_wall": 6.00,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 18.0, "floor_spec": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "u_fl": 1.10,
        "roof_area": 18.0, "ceil_spec": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Buffer cylinder & manifolds location. Unintentional heat gains from pipework."
    },
    {
        "code": "GF-11", "name": "Side Entrance Hall", "floor": "Ground Floor", "zone": "Lean to West",
        "temp": 18.0, "len": 2.00, "wid": 2.00, "ht": 2.60, "ext_wall": 4.00,
        "wall_spec": 'Solid Brick: 9" / 225mm (Uninsulated)', "u_wall": 2.10,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 4.0, "floor_spec": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "u_fl": 1.10,
        "roof_area": 4.0, "ceil_spec": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "u_roof": 2.00,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Suspended Timber (Unsealed Boards over Cold Void)",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Side entrance lobby (2.0m x 2.0m). External door with single glazing / draught seals."
    },

    # First Floor (13 rooms)
    {
        "code": "FF-01", "name": "Upper Stair Landing & Main Gallery", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 6.25, "wid": 4.50, "ht": 2.80, "ext_wall": 10.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 6.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 28.1, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 28.1, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Connects with GF-01 hall. Ceiling below insulated attic."
    },
    {
        "code": "FF-02", "name": "First Floor East Corridor", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 12.00, "wid": 1.80, "ht": 2.60, "ext_wall": 12.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 3.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 21.6, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 21.6, "ceil_spec": "Historic Thatched Roof (Straw / Reed Thatch)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Internal corridor with external end wall. Long pipe run access in floor void."
    },
    {
        "code": "FF-03", "name": "Master Bedroom", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 6.25, "wid": 5.50, "ht": 2.80, "ext_wall": 11.75,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 5.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 34.4, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 34.4, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Retrofitted Brush / Pile Draught Seals", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Dual aspect windows. Large Type 22 radiator recommended."
    },
    {
        "code": "FF-04", "name": "Master En-Suite Bathroom", "floor": "First Floor", "zone": "Georgian end",
        "temp": 22.0, "len": 4.50, "wid": 3.00, "ht": 2.80, "ext_wall": 4.50,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 13.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 13.5, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Retrofitted Brush / Pile Draught Seals", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Attic Downlights / Unsealed Loft Hatch",
        "notes": "CIBSE 22°C design temp. Underfloor heating mat or high-output towel radiator."
    },
    {
        "code": "FF-05", "name": "Bedroom 2 (Guest Suite)", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 5.50, "wid": 5.00, "ht": 2.60, "ext_wall": 10.50,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 4.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 27.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 27.5, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "New build high insulation section. Modest emitter sizing needed."
    },
    {
        "code": "FF-06", "name": "Bedroom 2 En-Suite", "floor": "First Floor", "zone": "New build",
        "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 1.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 7.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 7.5, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Warm 22°C design requirement. Extract fan with backdraught shutter."
    },
    {
        "code": "FF-07", "name": "Bedroom 3", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 5.50, "wid": 4.50, "ht": 2.60, "ext_wall": 10.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 3.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 24.8, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 24.8, "ceil_spec": "Historic Thatched Roof (Straw / Reed Thatch)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Dormer window and sloping ceiling. Ensure loft insulation reaches eaves."
    },
    {
        "code": "FF-08", "name": "Bedroom 4", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 18.0, "len": 5.00, "wid": 4.50, "ht": 2.60, "ext_wall": 9.50,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 3.5, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 22.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 22.5, "ceil_spec": "Historic Thatched Roof (Straw / Reed Thatch)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "Gable end wall exposed. Internal woodfibre insulation would reduce loss by 60%."
    },
    {
        "code": "FF-09", "name": "Bedroom 5 / Nursery", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 4.50, "wid": 4.00, "ht": 2.60, "ext_wall": 8.50,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 3.2, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 18.0, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 18.0, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Modern insulated timber/cavity construction."
    },
    {
        "code": "FF-10", "name": "Family Bathroom", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 22.0, "len": 4.00, "wid": 3.50, "ht": 2.60, "ext_wall": 4.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 2.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 14.0, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 14.0, "ceil_spec": "Historic Thatched Roof (Straw / Reed Thatch)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "High target temp (22°C). Needs dedicated radiator plus heated towel rail."
    },
    {
        "code": "FF-11", "name": "Laundry & Linen Room", "floor": "First Floor", "zone": "New build",
        "temp": 18.0, "len": 3.50, "wid": 3.00, "ht": 2.60, "ext_wall": 3.50,
        "wall_spec": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "u_wall": 0.18,
        "win_area": 1.5, "win_spec": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "u_win": 1.40,
        "fl_area": 10.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 10.5, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.15,
        "q_base": "New Build (Cavity / Insulated)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Modern High-Performance (Compression Gaskets)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Washing/drying equipment. Internal heat gains help offset space heating."
    },
    {
        "code": "FF-12", "name": "Dressing Room / Walk-in Robe", "floor": "First Floor", "zone": "Georgian end",
        "temp": 18.0, "len": 4.00, "wid": 3.00, "ht": 2.80, "ext_wall": 4.00,
        "wall_spec": 'Solid Brick: 18" / 450mm (Georgian Facade)', "u_wall": 1.40,
        "win_area": 1.8, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 12.0, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 12.0, "ceil_spec": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "u_roof": 0.18,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Insulated Loft (Sealed Plaster & Sealed Hatch)",
        "notes": "Fitted wardrobes on external wall create thermal shadowing. Keep background heat."
    },
    {
        "code": "FF-13", "name": "First Floor Shower Room / WC", "floor": "First Floor", "zone": "Thatched gable ended",
        "temp": 22.0, "len": 3.00, "wid": 2.50, "ht": 2.60, "ext_wall": 3.00,
        "wall_spec": "Solid Stone: 450-500mm Sandstone / Limestone", "u_wall": 1.80,
        "win_area": 1.0, "win_spec": "Single Glazed (Historic Timber Sash / Casement)", "u_win": 4.80,
        "fl_area": 7.5, "floor_spec": "Intermediate Floor (Heated Space Below)", "u_fl": 0.00,
        "roof_area": 7.5, "ceil_spec": "Historic Thatched Roof (Straw / Reed Thatch)", "u_roof": 0.30,
        "q_base": "Standard Historic (Solid Masonry)", "q_chimney": "No Chimney / Permanently Sealed",
        "q_win": "Original Loose Sash / Casement (Undraughted)", "q_floor": "Solid Concrete Slab / Insulated Floor",
        "q_ceil": "Sloping Roof / Exposed Eaves / Thatched Ridge",
        "notes": "22°C bathroom comfort criteria. Compact high-efficiency emitter."
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
    """Maps numeric U-value and optional zone name to the best-matching BS EN 12831 wall archetype."""
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
    """Maps numeric floor U-value, floor level, and zone to the best-matching BS EN 12831 floor archetype."""
    fl = floor_level.lower()
    z = zone.lower()
    if "first" in fl or "ff" in fl or "upper" in fl:
        return "Intermediate Floor (Heated Space Below)"
    if u_val <= 0.05:
        return "Intermediate Floor (Heated Space Below)"
    if u_val <= 0.12:
        return "High-Performance Insulated Slab (150mm+ PIR / Passivhaus)"
    if u_val <= 0.22:
        return "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)"
    if u_val <= 0.40:
        if "concrete" in z or "slab" in z:
            return "Solid Concrete: Moderate Insulation (50mm PIR / 1990s)"
        return "Suspended Timber: Insulated (50mm Quilt / Board)"
    if u_val <= 0.50:
        return "Solid Concrete: Perimeter Insulation (1980s Standard)"
    if u_val <= 0.70:
        return "Suspended Timber: Carpet & Underlay over Uninsulated Void"
    if u_val <= 0.95:
        if "concrete" in z or "slab" in z:
            return "Solid Concrete Ground Slab (Uninsulated Pre-1976)"
        return "Suspended Timber: Bare Boards over Cold Uninsulated Void"
    return "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth"

def extract_existing_rooms(ws: gspread.Worksheet) -> List[Dict[str, Any]]:
    """
    Reads existing user inputs from 2_Room_Heat_Loss worksheet non-destructively.
    Uses unformatted values to prevent truncation/rounding of dimensions and U-values.
    Supports 37-column (legacy), 38-column (with Wall Specification dropdown),
    and 39-column (with both Wall and Floor Specification dropdowns) layouts.
    """
    try:
        raw_values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except Exception:
        return []

    if len(raw_values) < 5:
        return []

    # Detect layout from row 4 (index 3) headers
    has_wall_spec = False
    has_floor_spec = False
    if len(raw_values) > 3:
        h_row = [str(c).strip().lower() for c in raw_values[3]]
        if len(h_row) > 11 and ("wall spec" in h_row[11] or "specification" in h_row[11]):
            has_wall_spec = True
        if len(h_row) > 19 and ("floor spec" in h_row[19] or "specification" in h_row[19]):
            has_floor_spec = True

    rooms = []
    # Room data rows start at row index 4 (Row 5 in spreadsheet)
    for row in raw_values[4:]:
        if not row:
            continue
        first_col = str(row[0]).strip() if len(row) > 0 else ""
        second_col = str(row[1]).strip() if len(row) > 1 else ""

        # Stop when encountering the Total row or a completely blank row
        if first_col.lower().startswith("total") or "whole house" in first_col.lower():
            break
        if first_col == "" and second_col == "":
            break

        floor_level = safe_str(row[2] if len(row) > 2 else "", "Ground Floor")
        zone = safe_str(row[3] if len(row) > 3 else "", "Old House")

        if has_floor_spec:
            # 39-column layout
            wall_spec = safe_str(row[11] if len(row) > 11 else "", "")
            u_wall = safe_float(row[12] if len(row) > 12 else 1.4, 1.4)
            if not wall_spec:
                wall_spec = map_u_to_wall_spec(u_wall, zone)
            win_area = safe_float(row[14] if len(row) > 14 else 0.0, 0.0)
            win_spec = safe_str(row[15] if len(row) > 15 else "", "Single Glazed (Historic Timber Sash / Casement)")
            fl_area = safe_float(row[18] if len(row) > 18 else 0.0, 0.0)
            floor_spec = safe_str(row[19] if len(row) > 19 else "", "")
            u_fl = safe_float(row[20] if len(row) > 20 else 0.8, 0.8)
            if not floor_spec:
                floor_spec = map_u_to_floor_spec(u_fl, floor_level, zone)
            roof_area = safe_float(row[22] if len(row) > 22 else 0.0, 0.0)
            ceil_spec = safe_str(row[23] if len(row) > 23 else "", "Intermediate Floor (Heated Space Above)")
            q_base = safe_str(row[26] if len(row) > 26 else "", "Standard Historic (Solid Masonry)")
            q_chimney = safe_str(row[27] if len(row) > 27 else "", "No Chimney / Permanently Sealed")
            q_win = safe_str(row[28] if len(row) > 28 else "", "Original Loose Sash / Casement (Undraughted)")
            q_floor = safe_str(row[29] if len(row) > 29 else "", "Solid Concrete Slab / Insulated Floor")
            q_ceil = safe_str(row[30] if len(row) > 30 else "", "Intermediate Floor (Heated Space Above)")
            notes = safe_str(row[38] if len(row) > 38 else "", "")
        elif has_wall_spec:
            # 38-column layout (current live sheet)
            wall_spec = safe_str(row[11] if len(row) > 11 else "", "")
            u_wall = safe_float(row[12] if len(row) > 12 else 1.4, 1.4)
            if not wall_spec:
                wall_spec = map_u_to_wall_spec(u_wall, zone)
            win_area = safe_float(row[14] if len(row) > 14 else 0.0, 0.0)
            win_spec = safe_str(row[15] if len(row) > 15 else "", "Single Glazed (Historic Timber Sash / Casement)")
            fl_area = safe_float(row[18] if len(row) > 18 else 0.0, 0.0)
            u_fl = safe_float(row[19] if len(row) > 19 else 0.8, 0.8)
            floor_spec = map_u_to_floor_spec(u_fl, floor_level, zone)
            roof_area = safe_float(row[21] if len(row) > 21 else 0.0, 0.0)
            ceil_spec = safe_str(row[22] if len(row) > 22 else "", "Intermediate Floor (Heated Space Above)")
            q_base = safe_str(row[25] if len(row) > 25 else "", "Standard Historic (Solid Masonry)")
            q_chimney = safe_str(row[26] if len(row) > 26 else "", "No Chimney / Permanently Sealed")
            q_win = safe_str(row[27] if len(row) > 27 else "", "Original Loose Sash / Casement (Undraughted)")
            q_floor = safe_str(row[28] if len(row) > 28 else "", "Solid Concrete Slab / Insulated Floor")
            q_ceil = safe_str(row[29] if len(row) > 29 else "", "Intermediate Floor (Heated Space Above)")
            notes = safe_str(row[37] if len(row) > 37 else "", "")
        else:
            # 37-column layout (legacy)
            u_wall = safe_float(row[11] if len(row) > 11 else 1.4, 1.4)
            wall_spec = map_u_to_wall_spec(u_wall, zone)
            win_area = safe_float(row[13] if len(row) > 13 else 0.0, 0.0)
            win_spec = safe_str(row[14] if len(row) > 14 else "", "Single Glazed (Historic Timber Sash / Casement)")
            fl_area = safe_float(row[17] if len(row) > 17 else 0.0, 0.0)
            u_fl = safe_float(row[18] if len(row) > 18 else 0.8, 0.8)
            floor_spec = map_u_to_floor_spec(u_fl, floor_level, zone)
            roof_area = safe_float(row[20] if len(row) > 20 else 0.0, 0.0)
            ceil_spec = safe_str(row[21] if len(row) > 21 else "", "Intermediate Floor (Heated Space Above)")
            q_base = safe_str(row[24] if len(row) > 24 else "", "Standard Historic (Solid Masonry)")
            q_chimney = safe_str(row[25] if len(row) > 25 else "", "No Chimney / Permanently Sealed")
            q_win = safe_str(row[26] if len(row) > 26 else "", "Original Loose Sash / Casement (Undraughted)")
            q_floor = safe_str(row[27] if len(row) > 27 else "", "Solid Concrete Slab / Insulated Floor")
            q_ceil = safe_str(row[28] if len(row) > 28 else "", "Intermediate Floor (Heated Space Above)")
            notes = safe_str(row[36] if len(row) > 36 else "", "")

        room = {
            "code": first_col,
            "name": safe_str(row[1] if len(row) > 1 else "", "Room"),
            "floor": floor_level,
            "zone": zone,
            "temp": safe_float(row[4] if len(row) > 4 else 20.0, 20.0),
            "len": safe_float(row[5] if len(row) > 5 else 0.0, 0.0),
            "wid": safe_float(row[6] if len(row) > 6 else 0.0, 0.0),
            "ht": safe_float(row[8] if len(row) > 8 else 2.6, 2.6),
            "ext_wall": safe_float(row[10] if len(row) > 10 else 0.0, 0.0),
            "wall_spec": wall_spec,
            "u_wall": u_wall,
            "win_area": win_area,
            "win_spec": win_spec,
            "fl_area": fl_area,
            "floor_spec": floor_spec,
            "u_fl": u_fl,
            "roof_area": roof_area,
            "ceil_spec": ceil_spec,
            "q_base": q_base,
            "q_chimney": q_chimney,
            "q_win": q_win,
            "q_floor": q_floor,
            "q_ceil": q_ceil,
            "notes": notes
        }
        rooms.append(room)

    return rooms

def update_named_ranges(ss: gspread.Spreadsheet, ws: gspread.Worksheet, total_row: int, num_rooms: int):
    """Registers / updates Google Sheets Named Ranges for dynamic downstream linking."""
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
                requests.append({
                    "deleteNamedRange": {
                        "namedRangeId": existing[name]["namedRangeId"]
                    }
                })
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
        ws = ss.add_worksheet(title=tab_name, rows=45, cols=39)

    # Step 1: Read existing rooms from sheet to preserve all user edits
    existing_rooms = extract_existing_rooms(ws)
    active_rooms = existing_rooms if len(existing_rooms) > 0 else DEFAULT_ROOMS

    num_rooms = len(active_rooms)
    last_room_row = 4 + num_rooms
    total_row = 5 + num_rooms  # 1-indexed row number of summary total row (e.g. 29 for 24 rooms)
    total_rows = total_row + 2
    total_cols = 39

    if ws.row_count < total_rows or ws.col_count < total_cols:
        ws.resize(rows=max(ws.row_count, total_rows), cols=max(ws.col_count, total_cols))

    grid: List[List[str]] = [["" for _ in range(total_cols)] for _ in range(total_rows)]

    # Row 1 & 2: Banner
    grid[0][0] = "2_Room_Heat_Loss: Room-by-Room Heat Loss Assessment & Low-Flow Radiator Sizing"
    grid[1][0] = f"Master MCS / CIBSE BS EN 12831 Room Schedule: {num_rooms} Assessed Rooms with Dynamic Wall, Floor & Infiltration Questionnaires."

    # Row 3: Category Groups
    grid[2][0] = "ROOM IDENTIFICATION"                   # Cols A-D (0-3)
    grid[2][4] = "DESIGN & GEOMETRY"                    # Cols E-J (4-9)
    grid[2][10] = "FABRIC LOSS CALCULATIONS (W)"         # Cols K-Z (10-25)
    grid[2][26] = "INFILTRATION & VENTILATION QUESTIONNAIRE" # Cols AA-AG (26-32)
    grid[2][33] = "TOTAL HEAT LOSS & EMITTERS"          # Cols AH-AK (33-36)
    grid[2][37] = "SPECIFICATION & NOTES"               # Cols AL-AM (37-38)

    # Row 4: Column Headers from canonical schema
    for c_i, h in enumerate(RoomCol.headers):
        grid[3][c_i] = h

    # Populate active rooms with live formulas
    for r_idx, rm in enumerate(active_rooms):
        r = 5 + r_idx  # 1-indexed row in spreadsheet
        wall_spec = rm.get("wall_spec", map_u_to_wall_spec(rm.get("u_wall", 1.40), rm.get("zone", "")))
        floor_spec = rm.get("floor_spec", map_u_to_floor_spec(rm.get("u_fl", 0.80), rm.get("floor", "Ground Floor"), rm.get("zone", "")))
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
            wall_spec,                                        # Col L (11) (Wall Specification Dropdown)
            f"=VLOOKUP(L{r}, '1_Inputs'!$B$129:$C$144, 2, FALSE)", # Col M (12) (U Wall Formula)
            f"=MAX(0, (K{r}*I{r}-O{r})*M{r}*(E{r}-'1_Inputs'!$C$5))", # Col N (13) (Wall Loss W)
            str(rm["win_area"]),                              # Col O (14)
            rm["win_spec"],                                   # Col P (15) (Window Specification Dropdown)
            f"=VLOOKUP(P{r}, '1_Inputs'!$B$98:$C$108, 2, FALSE)", # Col Q (16) (U Window Formula)
            f"=O{r}*Q{r}*(E{r}-'1_Inputs'!$C$5)",             # Col R (17) (Window Loss W)
            str(rm["fl_area"]),                               # Col S (18)
            floor_spec,                                       # Col T (19) (Floor Specification Dropdown)
            f"=VLOOKUP(T{r}, '1_Inputs'!$B$148:$C$161, 2, FALSE)", # Col U (20) (U Floor Formula)
            f"=S{r}*U{r}*(E{r}-'1_Inputs'!$C$6)",             # Col V (21) (Floor Loss with Ground Temp C6!)
            str(rm["roof_area"]),                             # Col W (22)
            rm.get("ceil_spec", "Intermediate Floor (Heated Space Above)"), # Col X (23) (Ceiling Specification)
            f"=VLOOKUP(X{r}, '1_Inputs'!$B$113:$C$124, 2, FALSE)", # Col Y (24) (U Ceiling Formula)
            f"=W{r}*Y{r}*(E{r}-'1_Inputs'!$C$5)",             # Col Z (25) (Ceiling Loss W)
            rm["q_base"],                                     # Col AA (26) - Dropdown Q1
            rm["q_chimney"],                                  # Col AB (27) - Dropdown Q2
            rm["q_win"],                                      # Col AC (28) - Dropdown Q3
            rm["q_floor"],                                    # Col AD (29) - Dropdown Q4
            rm["q_ceil"],                                     # Col AE (30) - Dropdown Q5
            f"=VLOOKUP(AA{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AB{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AC{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AD{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE) + VLOOKUP(AE{r}, '1_Inputs'!$B$76:$C$94, 2, FALSE)", # Col AF (31)
            f"='1_Inputs'!$C$11*AF{r}*J{r}*(E{r}-'1_Inputs'!$C$5)", # Col AG (32) (Vent Loss W)
            f"=SUM(N{r}, R{r}, V{r}, Z{r}, AG{r})",           # Col AH (33) (Total Room Loss W)
            f"=AH{r}/H{r}",                                   # Col AI (34) (W/m²)
            f"=AH{r}",                                        # Col AJ (35) (Req Rad at ΔT30)
            f"=ROUND(AH{r}*1.89, 0)",                         # Col AK (36) (Boiler equivalent at ΔT50)
            f"=IF(AI{r}>100, \"Type 33 or 2x Type 22\", IF(AI{r}>65, \"Type 22 High-Output\", \"Type 21 / Underfloor\"))", # Col AL (37)
            rm["notes"]                                       # Col AM (38)
        ]
        grid[4 + r_idx] = row_arr

    # Summary Total Row
    grid[total_row - 1] = [
        "Total Whole House",
        f"{num_rooms} Assessed Rooms",
        "Ground & First",
        "All Wings",
        f"=SUMPRODUCT(E5:E{last_room_row}, H5:H{last_room_row})/SUM(H5:H{last_room_row})",
        "-",
        "-",
        f"=SUM(H5:H{last_room_row})",
        "-",
        f"=SUM(J5:J{last_room_row})",
        "-",
        "-",
        "-",
        f"=SUM(N5:N{last_room_row})",
        f"=SUM(O5:O{last_room_row})",
        "-",
        "-",
        f"=SUM(R5:R{last_room_row})",
        f"=SUM(S5:S{last_room_row})",
        "-",
        "-",
        f"=SUM(V5:V{last_room_row})",
        f"=SUM(W5:W{last_room_row})",
        "-",
        "-",
        f"=SUM(Z5:Z{last_room_row})",
        "-",
        "-",
        "-",
        "-",
        "-",
        f"=AVERAGE(AF5:AF{last_room_row})",
        f"=SUM(AG5:AG{last_room_row})",
        f"=SUM(AH5:AH{last_room_row})",
        f"=AH{total_row}/H{total_row}",
        f"=SUM(AJ5:AJ{last_room_row})",
        f"=SUM(AK5:AK{last_room_row})",
        "Whole House Emitters",
        "Survey Complete"
    ]

    # Write values into worksheet
    ws.update(values=grid, range_name=f"A1:AM{total_rows}", value_input_option="USER_ENTERED")

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
    for c_i in range(5, 11):
        fmt_reqs.append(create_set_column_width_request(ws.id, c_i, c_i + 1, 95)) # Len to Ext Wall
    fmt_reqs.append(create_set_column_width_request(ws.id, 11, 12, 240)) # Wall Specification
    fmt_reqs.append(create_set_column_width_request(ws.id, 12, 13, 95))  # U Wall
    fmt_reqs.append(create_set_column_width_request(ws.id, 13, 14, 95))  # Wall Loss
    fmt_reqs.append(create_set_column_width_request(ws.id, 14, 15, 95))  # Window Area
    fmt_reqs.append(create_set_column_width_request(ws.id, 15, 16, 240)) # Window Specification
    fmt_reqs.append(create_set_column_width_request(ws.id, 16, 17, 95))  # U Window
    fmt_reqs.append(create_set_column_width_request(ws.id, 17, 18, 95))  # Window Loss
    fmt_reqs.append(create_set_column_width_request(ws.id, 18, 19, 95))  # Exp Floor
    fmt_reqs.append(create_set_column_width_request(ws.id, 19, 20, 240)) # Floor Specification
    fmt_reqs.append(create_set_column_width_request(ws.id, 20, 21, 95))  # U Floor
    fmt_reqs.append(create_set_column_width_request(ws.id, 21, 22, 95))  # Floor Loss
    fmt_reqs.append(create_set_column_width_request(ws.id, 22, 23, 95))  # Ceiling Area
    fmt_reqs.append(create_set_column_width_request(ws.id, 23, 24, 240)) # Ceiling Specification
    fmt_reqs.append(create_set_column_width_request(ws.id, 24, 25, 95))  # U Ceiling
    fmt_reqs.append(create_set_column_width_request(ws.id, 25, 26, 95))  # Ceiling Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 26, 27, 230)) # Base Construction
    fmt_reqs.append(create_set_column_width_request(ws.id, 27, 28, 220)) # Chimney / Flue
    fmt_reqs.append(create_set_column_width_request(ws.id, 28, 29, 230)) # Windows & Doors
    fmt_reqs.append(create_set_column_width_request(ws.id, 29, 30, 230)) # Floor Type
    fmt_reqs.append(create_set_column_width_request(ws.id, 30, 31, 230)) # Ceiling Boundary
    fmt_reqs.append(create_set_column_width_request(ws.id, 31, 32, 105)) # Calc ACH
    fmt_reqs.append(create_set_column_width_request(ws.id, 32, 33, 105)) # Vent Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 33, 34, 115)) # Room Loss W
    fmt_reqs.append(create_set_column_width_request(ws.id, 34, 35, 105)) # Intensity
    fmt_reqs.append(create_set_column_width_request(ws.id, 35, 36, 120)) # Rad 45C
    fmt_reqs.append(create_set_column_width_request(ws.id, 36, 37, 120)) # Boiler Rad
    fmt_reqs.append(create_set_column_width_request(ws.id, 37, 38, 165)) # Recommended Emitter
    fmt_reqs.append(create_set_column_width_request(ws.id, 38, 39, 360)) # Notes

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

    # Category Group Headers (Row 3) dynamically derived from schema
    group_spans = [
        (0, RoomCol.idx("LENGTH"), THEME["SECONDARY_HEADER_BG"]),
        (RoomCol.idx("LENGTH"), RoomCol.idx("EXT_WALL_L"), THEME["SECTION_HEADER_BG"]),
        (RoomCol.idx("EXT_WALL_L"), RoomCol.idx("Q_BASE"), {"red": 0.15, "green": 0.35, "blue": 0.45}),
        (RoomCol.idx("Q_BASE"), RoomCol.idx("TOTAL_LOSS"), {"red": 0.10, "green": 0.32, "blue": 0.42}),
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

    # Data Rows (Rows 5 to last_room_row)
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
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 0, 4,
            align="LEFT",
            font_size=9
        ))
        # Wall Specification Dropdown (Col 11 / L)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 11, 12,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # U Wall Formula (Col 12 / M)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 12, 13,
            align="RIGHT",
            font_size=9,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Window Specification Dropdown (Col 15 / P)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 15, 16,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # U Window Formula (Col 16 / Q)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 16, 17,
            align="RIGHT",
            font_size=9,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Floor Specification Dropdown (Col 19 / T)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 19, 20,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # U Floor Formula (Col 20 / U)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 20, 21,
            align="RIGHT",
            font_size=9,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Ceiling Specification Dropdown (Col 23 / X)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 23, 24,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # U Ceiling Formula (Col 24 / Y)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 24, 25,
            align="RIGHT",
            font_size=9,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Questionnaire Dropdowns (Cols 26 to 30 / AA to AE)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 26, 31,
            bg_color=THEME["INPUT_BG"],
            font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
            align="LEFT",
            font_size=9
        ))
        # Calculated ACH (Col 31 / AF)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 31, 32,
            bg_color=THEME["CALC_BG"],
            bold=True,
            align="RIGHT",
            font_size=10,
            number_format=FORMATS["DECIMAL_2"]
        ))
        # Recommended Emitter & Notes (Cols 37, 38 / AL, AM)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, r_i, r_i + 1, 37, 39,
            align="LEFT",
            font_size=9,
            font_color=THEME["MUTED_TEXT"]
        ))

    # Number formats
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 4, 5, number_format=FORMATS["TEMP_C"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 5, 9, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 10, 11, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 12, 13, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 13, 14, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 14, 15, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 16, 17, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 17, 18, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 18, 19, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 20, 21, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 21, 22, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 22, 23, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 24, 25, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 25, 26, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 31, 32, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 32, 34, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 34, 35, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, 4, last_room_row, 35, 37, number_format=FORMATS["INTEGER"]))

    # Total Row (Row total_row, index total_row - 1 = last_room_row)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, last_room_row, total_row, 0, total_cols,
        bg_color=THEME["TOTAL_BG"],
        font_color=THEME["DARK_TEXT"],
        bold=True,
        font_size=10,
        align="RIGHT"
    ))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 0, 4, align="LEFT"))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 7, 8, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 9, 10, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 13, 14, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 14, 15, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 17, 18, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 18, 19, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 21, 22, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 22, 23, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 25, 26, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 31, 32, number_format=FORMATS["DECIMAL_2"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 32, 34, number_format=FORMATS["INTEGER"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 34, 35, number_format=FORMATS["DECIMAL_1"]))
    fmt_reqs.append(create_repeat_cell_request(ws.id, last_room_row, total_row, 35, 37, number_format=FORMATS["INTEGER"]))

    # Accent on Total Room Heat Loss (Col AH / 33)
    fmt_reqs.append(create_repeat_cell_request(
        ws.id, last_room_row, total_row, 33, 34,
        bg_color=THEME["ACCENT_BG"],
        font_color={"red": 0.70, "green": 0.20, "blue": 0.05},
        bold=True,
        font_size=11,
        align="RIGHT"
    ))

    # Add Native Google Sheets Dropdown Data Validations using RoomCol
    # 1. Wall Specification
    fmt_reqs.append(create_data_validation_request(
        sheet_id=ws.id,
        start_row=4,
        end_row=last_room_row,
        start_col=RoomCol.WALL_SPEC_IDX,
        end_col=RoomCol.WALL_SPEC_IDX + 1,
        options=[opt["label"] for opt in WALL_SPECIFICATIONS],
        show_custom_ui=True,
        strict=False
    ))

    # 2. Window Specification
    fmt_reqs.append(create_data_validation_request(
        sheet_id=ws.id,
        start_row=4,
        end_row=last_room_row,
        start_col=RoomCol.WIN_SPEC_IDX,
        end_col=RoomCol.WIN_SPEC_IDX + 1,
        options=[opt["label"] for opt in WINDOW_SPECIFICATIONS],
        show_custom_ui=True,
        strict=False
    ))

    # 3. Floor Specification
    fmt_reqs.append(create_data_validation_request(
        sheet_id=ws.id,
        start_row=4,
        end_row=last_room_row,
        start_col=RoomCol.FLOOR_SPEC_IDX,
        end_col=RoomCol.FLOOR_SPEC_IDX + 1,
        options=[opt["label"] for opt in FLOOR_SPECIFICATIONS],
        show_custom_ui=True,
        strict=False
    ))

    # 4. Ceiling Specification
    fmt_reqs.append(create_data_validation_request(
        sheet_id=ws.id,
        start_row=4,
        end_row=last_room_row,
        start_col=RoomCol.CEIL_SPEC_IDX,
        end_col=RoomCol.CEIL_SPEC_IDX + 1,
        options=[opt["label"] for opt in CEILING_SPECIFICATIONS],
        show_custom_ui=True,
        strict=False
    ))

    # 5. Infiltration Questionnaire Dropdowns
    q_categories = [
        (RoomCol.Q_BASE_IDX, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["base_construction"]]),
        (RoomCol.Q_CHIMNEY_IDX, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["chimney_flue"]]),
        (RoomCol.Q_WIN_IDX, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["windows_doors"]]),
        (RoomCol.Q_FLOOR_IDX, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["floor_construction"]]),
        (RoomCol.Q_CEIL_IDX, [opt["label"] for opt in INFILTRATION_QUESTIONNAIRE["ceiling_boundary"]]),
    ]
    for col_idx, options in q_categories:
        fmt_reqs.append(create_data_validation_request(
            sheet_id=ws.id,
            start_row=4,
            end_row=last_room_row,
            start_col=col_idx,
            end_col=col_idx + 1,
            options=options,
            show_custom_ui=True,
            strict=False
        ))

    # Update Named Ranges so downstream tabs and formulas are completely dynamic
    update_named_ranges(ss, ws, total_row, num_rooms)

    return ws, fmt_reqs, num_rooms, total_row
