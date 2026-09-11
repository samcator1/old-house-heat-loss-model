"""
Canonical Column Schema for 2_Room_Heat_Loss Sheet.
Defines column positions, Excel letters, headers, and formatting spans.
Eliminates hardcoded column letters across the codebase.
"""

from typing import List, Dict, Tuple, Any

def col_idx_to_letter(idx: int) -> str:
    """Converts a 0-indexed column integer (0 -> 'A', 25 -> 'Z', 26 -> 'AA', 33 -> 'AH') to Excel column letter."""
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

COLUMN_DEFINITIONS: List[Tuple[str, str, int, str]] = [
    ("CODE", "Room Code", 95, "ROOM"),
    ("NAME", "Room Name", 220, "ROOM"),
    ("FLOOR", "Floor Level", 100, "ROOM"),
    ("ZONE", "Zone / Wing", 150, "ROOM"),
    ("TI", "Design Ti (°C)", 105, "ROOM"),
    ("LENGTH", "Length (m)", 85, "DIMENSIONS"),
    ("WIDTH", "Width (m)", 85, "DIMENSIONS"),
    ("AREA", "Floor Area (m²)", 95, "DIMENSIONS"),
    ("HEIGHT", "Height (m)", 85, "DIMENSIONS"),
    ("VOLUME", "Volume (m³)", 95, "DIMENSIONS"),
    ("EXT_WALL_L", "Ext Wall (m)", 95, "FABRIC"),
    ("WALL_SPEC", "Wall Specification", 260, "FABRIC"),
    ("U_WALL", "U Wall (W/m²K)", 105, "FABRIC"),
    ("WALL_LOSS", "Wall Loss (W)", 100, "FABRIC"),
    ("WIN_AREA", "Window Area (m²)", 110, "FABRIC"),
    ("WIN_SPEC", "Window Specification", 260, "FABRIC"),
    ("U_WIN", "U Window (W/m²K)", 115, "FABRIC"),
    ("WIN_LOSS", "Window Loss (W)", 110, "FABRIC"),
    ("DOOR_AREA", "Door Area (m²)", 105, "FABRIC"),
    ("DOOR_SPEC", "Door Specification", 260, "FABRIC"),
    ("U_DOOR", "U Door (W/m²K)", 110, "FABRIC"),
    ("DOOR_LOSS", "Door Loss (W)", 105, "FABRIC"),
    ("FL_AREA", "Exposed Floor Area (m²)", 125, "FABRIC"),
    ("FLOOR_SPEC", "Floor Specification", 260, "FABRIC"),
    ("U_FLOOR", "U Floor (W/m²K)", 110, "FABRIC"),
    ("FLOOR_LOSS", "Floor Loss (W)", 105, "FABRIC"),
    ("ROOF_AREA", "Ceiling Area (m²)", 110, "FABRIC"),
    ("CEIL_SPEC", "Ceiling Specification", 260, "FABRIC"),
    ("U_CEIL", "U Ceiling (W/m²K)", 115, "FABRIC"),
    ("CEIL_LOSS", "Ceiling Loss (W)", 110, "FABRIC"),
    ("TB_LOSS", "Thermal Bridge Loss (W)", 125, "FABRIC"),
    ("CHIMNEY", "Chimney / Fireplace", 200, "VENTILATION"),
    ("MECH_VENT", "Mechanical Ventilation", 220, "VENTILATION"),
    ("ACH", "Calculated ACH", 105, "VENTILATION"),
    ("VENT_LOSS", "Vent Loss (W)", 105, "VENTILATION"),
    ("TOTAL_LOSS", "Room Heat Loss (W)", 135, "TOTALS"),
    ("INTENSITY", "Heat Loss Intensity (W/m²)", 140, "TOTALS"),
    ("RAD_45", "Req Emitter Output (45°C Flow) (W)", 165, "TOTALS"),
    ("RAD_DT50", "Catalogue Rating Req (ΔT50) (W)", 165, "TOTALS"),
    ("FLOW_RATE", "Design Flow Rate (l/h)", 125, "TOTALS"),
    ("REC_PIPE", "Min Pipe Size", 130, "TOTALS"),
    ("EMITTER_TYPE", "Planned Emitter Type", 210, "TOTALS"),
    ("REC_EMITTER", "Recommended Sizing", 200, "TOTALS"),
    ("NOTES", "Notes / Survey Observations", 260, "NOTES"),
]

class _RoomColumnRegistry:
    def __init__(self, defs: List[Tuple[str, str, int, str]]):
        self._defs = defs
        self._key_to_idx: Dict[str, int] = {}
        self._key_to_letter: Dict[str, str] = {}
        self._headers: List[str] = []
        self._widths: List[int] = []

        for idx, (key, header, width, category) in enumerate(defs):
            letter = col_idx_to_letter(idx)
            self._key_to_idx[key] = idx
            self._key_to_letter[key] = letter
            self._headers.append(header)
            self._widths.append(width)

            setattr(self, f"{key}_IDX", idx)
            setattr(self, f"{key}_LETTER", letter)
            setattr(self, key, letter)

            # Backwards compatibility alias for Q_CHIMNEY
            if key == "CHIMNEY":
                setattr(self, "Q_CHIMNEY_IDX", idx)
                setattr(self, "Q_CHIMNEY_LETTER", letter)
                setattr(self, "Q_CHIMNEY", letter)

    @property
    def total_cols(self) -> int:
        return len(self._defs)

    @property
    def headers(self) -> List[str]:
        return list(self._headers)

    @property
    def widths(self) -> List[int]:
        return list(self._widths)

    def idx(self, key: str) -> int:
        return self._key_to_idx[key]

    def letter(self, key: str) -> str:
        return self._key_to_letter[key]

    @property
    def fabric_span(self) -> Tuple[int, int]:
        return (self.idx("EXT_WALL_L"), self.idx("TB_LOSS") + 1)

    @property
    def vent_span(self) -> Tuple[int, int]:
        return (self.idx("CHIMNEY"), self.idx("VENT_LOSS") + 1)

    @property
    def totals_span(self) -> Tuple[int, int]:
        return (self.idx("TOTAL_LOSS"), self.idx("REC_EMITTER") + 1)

    @property
    def notes_span(self) -> Tuple[int, int]:
        return (self.idx("NOTES"), self.idx("NOTES") + 1)

RoomCol = _RoomColumnRegistry(COLUMN_DEFINITIONS)
