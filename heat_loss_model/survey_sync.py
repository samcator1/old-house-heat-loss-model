"""
Survey Snapshot & Export Module.
Exports live room data and U-value specifications from Google Sheets to:
- room_by_room_heat_loss_survey.csv (CSV backup snapshot)
"""

import os
import sys
import csv
import re
from pathlib import Path
from typing import List, Dict, Any, Optional

import gspread
from .config import (
    WINDOW_SPECIFICATIONS,
    CEILING_SPECIFICATIONS,
    WALL_SPECIFICATIONS,
    FLOOR_SPECIFICATIONS,
    DOOR_SPECIFICATIONS,
    MECHANICAL_VENTILATION_SPECIFICATIONS,
    EMITTER_SPECIFICATIONS
)
from .tabs.room_tab import extract_existing_rooms
from .schema import RoomCol

CSV_HEADERS = [
    "Room Code",
    "Room Name",
    "Floor Level",
    "Zone / Wing",
    "Design Ti (°C)",
    "Length (m)",
    "Width (m)",
    "Height (m)",
    "Ext Wall (m)",
    "Wall Specification",
    "U-Wall",
    "Window Area (m2)",
    "Window Specification",
    "Door Area (m2)",
    "Door Specification",
    "U-Door",
    "Floor Area (m2)",
    "Floor Specification",
    "U-Floor",
    "Ceiling Area (m2)",
    "Ceiling Specification",
    "Chimney / Fireplace",
    "Mechanical Ventilation",
    "Planned Emitter Type",
    "Wall Type",
    "Floor Type",
    "Radiator Type",
    "Rad Size",
    "Pipework",
    "Notes"
]

def get_wall_u(spec: str) -> float:
    for item in WALL_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 1.40))
    return 1.40

def get_window_u(spec: str) -> float:
    for item in WINDOW_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 4.80))
    return 4.80

def get_door_u(spec: str) -> float:
    for item in DOOR_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 0.00))
    return 0.00

def get_floor_u(spec: str) -> float:
    for item in FLOOR_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 0.80))
    return 0.80

def get_ceiling_u(spec: str) -> float:
    for item in CEILING_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 0.18))
    return 0.18

def parse_room_notes(raw_notes: str) -> Dict[str, str]:
    """Parses optional pipe-delimited metadata from notes if present."""
    result = {
        "wall_type": "",
        "floor_type": "",
        "rad_type": "",
        "rad_size": "",
        "pipe_size": "",
        "clean_notes": raw_notes
    }
    if not raw_notes:
        return result

    parts = [p.strip() for p in raw_notes.split("|")]
    clean_parts = []
    for part in parts:
        if part.lower().startswith("construction:"):
            result["wall_type"] = part.split(":", 1)[1].strip()
        elif part.lower().startswith("rad:"):
            rad_part = part.split(":", 1)[1].strip()
            m = re.search(r"\((.*?)\)", rad_part)
            if m:
                result["pipe_size"] = m.group(1).strip()
                rad_part = rad_part[:m.start()].strip()
            rad_tokens = rad_part.split()
            if len(rad_tokens) >= 2 and ("x" in rad_tokens[-1] or "X" in rad_tokens[-1]):
                result["rad_size"] = rad_tokens[-1]
                result["rad_type"] = " ".join(rad_tokens[:-1])
            else:
                result["rad_type"] = rad_part
        else:
            if part not in clean_parts:
                clean_parts.append(part)
    result["clean_notes"] = " | ".join(clean_parts) if clean_parts else raw_notes
    return result

def export_rooms_to_csv(rooms: List[Dict[str, Any]], csv_path: Path) -> None:
    """Exports room list to CSV compatible with import_survey_csv.py."""
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        for rm in rooms:
            parsed = parse_room_notes(rm.get("notes", ""))
            wall_spec = rm.get("wall_spec", 'Solid Brick: 18" / 450mm (Georgian Facade)')
            floor_spec = rm.get("floor_spec", "Suspended Timber: Bare Boards Over Void")
            row = [
                rm.get("code", ""),
                rm.get("name", ""),
                rm.get("floor", ""),
                rm.get("zone", ""),
                rm.get("temp", 20.0),
                rm.get("len", 0.0),
                rm.get("wid", 0.0),
                rm.get("ht", 2.6),
                rm.get("ext_wall", 0.0),
                wall_spec,
                rm.get("u_wall", get_wall_u(wall_spec)),
                rm.get("win_area", 0.0),
                rm.get("win_spec", "Single Glazed (Historic Timber Sash / Casement)"),
                rm.get("door_area", 0.0),
                rm.get("door_spec", "No External Door (Internal Boundary Only)"),
                rm.get("u_door", get_door_u(rm.get("door_spec", ""))),
                rm.get("fl_area", 0.0),
                floor_spec,
                rm.get("u_fl", get_floor_u(floor_spec)),
                rm.get("roof_area", 0.0),
                rm.get("ceil_spec", "Intermediate Floor (Heated Space Above)"),
                rm.get("chimney", rm.get("q_chimney", "No Chimney / Permanently Sealed")),
                rm.get("mech_vent", "None (Natural Infiltration Only)"),
                rm.get("emitter_type", "Type 22 (Double Convector)"),
                rm.get("wall_type", parsed["wall_type"] or "Solid Masonry"),
                rm.get("floor_type", parsed["floor_type"] or "Standard"),
                rm.get("rad_type", parsed["rad_type"] or "Type 22"),
                rm.get("rad_size", parsed["rad_size"] or "600x1000"),
                rm.get("pipe_size", parsed["pipe_size"] or "15mm copper"),
                rm.get("notes", parsed["clean_notes"])
            ]
            writer.writerow(row)

def sync_survey_artifacts(ss: gspread.Spreadsheet, repo_root: Path) -> Dict[str, Any]:
    """
    Exports a clean CSV snapshot from Google Sheets 2_Room_Heat_Loss.
    """
    ws = ss.worksheet("2_Room_Heat_Loss")
    rooms = extract_existing_rooms(ws)
    if not rooms:
        return {"status": "error", "message": "No rooms found on 2_Room_Heat_Loss worksheet."}

    csv_path = repo_root / "room_by_room_heat_loss_survey.csv"
    export_rooms_to_csv(rooms, csv_path)

    return {
        "status": "success",
        "num_rooms": len(rooms),
        "csv_path": str(csv_path)
    }
