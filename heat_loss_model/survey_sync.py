"""
Survey Synchronization Module.
Coordinates exporting live room data and U-value specifications from Google Sheets to:
1. room_by_room_heat_loss_survey.csv (CSV snapshot)
2. index.html (preloaded JavaScript rooms array and live U-value archetypes from 1_Inputs)
3. GitHub Pages git commit & push (optional)
"""

import os
import sys
import csv
import json
import re
import subprocess
from pathlib import Path
from typing import List, Dict, Any, Optional

import gspread
from .tabs.room_tab import (
    extract_existing_rooms,
    WINDOW_SPECIFICATIONS,
    CEILING_SPECIFICATIONS,
    WALL_SPECIFICATIONS,
    FLOOR_SPECIFICATIONS
)
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
    "Floor Area (m2)",
    "Floor Specification",
    "U-Floor",
    "Ceiling Area (m2)",
    "Ceiling Specification",
    "Base Construction",
    "Chimney Flue",
    "Windows Doors",
    "Floor Boundary",
    "Ceiling Boundary",
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
            clean_parts.append(part)
    result["clean_notes"] = " | ".join(clean_parts) if clean_parts else raw_notes
    return result

def extract_specs_from_inputs(ss: gspread.Spreadsheet) -> Dict[str, List[Dict[str, Any]]]:
    """Reads live archetypes and U-values directly from 1_Inputs in Google Sheets."""
    try:
        ws_inputs = ss.worksheet("1_Inputs")
        vals = ws_inputs.get_values("B96:C165")
    except Exception as e:
        return {
            "walls": WALL_SPECIFICATIONS,
            "windows": WINDOW_SPECIFICATIONS,
            "ceilings": CEILING_SPECIFICATIONS,
            "floors": FLOOR_SPECIFICATIONS
        }

    def parse_section(start_row: int, end_row: int, default_fallback: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        items = []
        for r_num in range(start_row, end_row + 1):
            idx = r_num - 96
            if idx < len(vals) and len(vals[idx]) >= 2 and vals[idx][0].strip():
                label = vals[idx][0].strip()
                try:
                    u_val = float(vals[idx][1].strip())
                except ValueError:
                    u_val = 0.0
                items.append({"label": label, "u_val": u_val})
        return items if len(items) > 0 else default_fallback

    return {
        "windows": parse_section(98, 108, WINDOW_SPECIFICATIONS),
        "ceilings": parse_section(113, 124, CEILING_SPECIFICATIONS),
        "walls": parse_section(129, 144, WALL_SPECIFICATIONS),
        "floors": parse_section(148, 161, FLOOR_SPECIFICATIONS),
    }

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
                rm.get("fl_area", 0.0),
                floor_spec,
                rm.get("u_fl", get_floor_u(floor_spec)),
                rm.get("roof_area", 0.0),
                rm.get("ceil_spec", "Intermediate Floor (Heated Space Above)"),
                rm.get("q_base", "Standard Historic (Solid Masonry)"),
                rm.get("q_chimney", "No Chimney / Permanently Sealed"),
                rm.get("q_win", "Original Loose Sash / Casement (Undraughted)"),
                rm.get("q_floor", "Solid Concrete Slab / Insulated Floor"),
                rm.get("q_ceil", "Intermediate Floor (Heated Space Above)"),
                rm.get("wall_type", parsed["wall_type"] or "Solid Masonry"),
                rm.get("floor_type", parsed["floor_type"] or "Standard"),
                rm.get("rad_type", parsed["rad_type"] or "Type 22"),
                rm.get("rad_size", parsed["rad_size"] or "600x1000"),
                rm.get("pipe_size", parsed["pipe_size"] or "15mm copper"),
                rm.get("notes", parsed["clean_notes"])
            ]
            writer.writerow(row)

def update_html_specs_and_rooms(rooms: List[Dict[str, Any]], specs: Dict[str, List[Dict[str, Any]]], html_path: Path) -> None:
    """Updates pre-loaded specs and rooms array in index.html."""
    if not html_path.exists():
        return

    content = html_path.read_text(encoding="utf-8")

    # Helper to serialize specs array
    def format_spec_js(spec_list: List[Dict[str, Any]]) -> str:
        lines = []
        for item in spec_list:
            u = item.get("u_val", item.get("u_value", 0.0))
            lbl = item["label"].replace('"', '\\"')
            lines.append(f'      {{ label: "{lbl}", u_val: {u:.2f} }}')
        body = ",\n".join(lines)
        return f"[\n{body}\n    ];"

    # 1. Update Specs
    if "walls" in specs and specs["walls"]:
        pattern = re.compile(r"const WALL_SPECIFICATIONS = \[\s*\{.*?\n    \];", re.DOTALL)
        content = pattern.sub(f"const WALL_SPECIFICATIONS = {format_spec_js(specs['walls'])}", content, count=1)

    if "windows" in specs and specs["windows"]:
        pattern = re.compile(r"const WINDOW_SPECIFICATIONS = \[\s*\{.*?\n    \];", re.DOTALL)
        content = pattern.sub(f"const WINDOW_SPECIFICATIONS = {format_spec_js(specs['windows'])}", content, count=1)

    if "ceilings" in specs and specs["ceilings"]:
        pattern = re.compile(r"const CEILING_SPECIFICATIONS = \[\s*\{.*?\n    \];", re.DOTALL)
        content = pattern.sub(f"const CEILING_SPECIFICATIONS = {format_spec_js(specs['ceilings'])}", content, count=1)

    if "floors" in specs and specs["floors"]:
        pattern = re.compile(r"const FLOOR_SPECIFICATIONS = \[\s*\{.*?\n    \];", re.DOTALL)
        content = pattern.sub(f"const FLOOR_SPECIFICATIONS = {format_spec_js(specs['floors'])}", content, count=1)

    # 2. Update Rooms Array
    js_objects = []
    for rm in rooms:
        parsed = parse_room_notes(rm.get("notes", ""))
        wall_spec = rm.get("wall_spec", 'Solid Brick: 18" / 450mm (Georgian Facade)')
        floor_spec = rm.get("floor_spec", "Suspended Timber: Bare Boards Over Void")
        win_spec = rm.get("win_spec", "Single Glazed (Historic Timber Sash / Casement)")
        ceil_spec = rm.get("ceil_spec", "Intermediate Floor (Heated Space Above)")

        js_obj = {
            "code": rm.get("code", ""),
            "name": rm.get("name", ""),
            "floor": rm.get("floor", "Ground Floor"),
            "zone": rm.get("zone", "Old House"),
            "temp": rm.get("temp", 20.0),
            "len": rm.get("len", 0.0),
            "wid": rm.get("wid", 0.0),
            "ht": rm.get("ht", 2.6),
            "extWall": rm.get("ext_wall", 0.0),
            "wallSpec": wall_spec,
            "uWall": get_wall_u(wall_spec),
            "winArea": rm.get("win_area", 0.0),
            "winSpec": win_spec,
            "uWin": get_window_u(win_spec),
            "flArea": rm.get("fl_area", 0.0),
            "floorSpec": floor_spec,
            "uFl": rm.get("u_fl", get_floor_u(floor_spec)),
            "roofArea": rm.get("roof_area", 0.0),
            "ceilSpec": ceil_spec,
            "uRoof": get_ceiling_u(ceil_spec),
            "qBase": rm.get("q_base", "Standard Historic (Solid Masonry)"),
            "qChimney": rm.get("q_chimney", "No Chimney / Permanently Sealed"),
            "qWin": rm.get("q_win", "Original Loose Sash / Casement (Undraughted)"),
            "qFloor": rm.get("q_floor", "Solid Concrete Slab / Insulated Floor"),
            "qCeil": rm.get("q_ceil", "Intermediate Floor (Heated Space Above)"),
            "wallType": rm.get("wall_type", parsed["wall_type"] or "Solid Masonry"),
            "floorType": rm.get("floor_type", parsed["floor_type"] or "Standard"),
            "radType": rm.get("rad_type", parsed["rad_type"] or "Type 22"),
            "radSize": rm.get("rad_size", parsed["rad_size"] or "600x1000"),
            "pipeSize": rm.get("pipe_size", parsed["pipe_size"] or "15mm copper"),
            "notes": rm.get("notes", parsed["clean_notes"])
        }
        js_line = "      " + json.dumps(js_obj, ensure_ascii=False)
        js_objects.append(js_line)

    js_array_body = ",\n".join(js_objects)
    new_block = f"    let rooms = [\n{js_array_body}\n    ];"
    pattern = re.compile(r"    let rooms = \[\s*\{.*?\n    \];", re.DOTALL)

    if pattern.search(content):
        content = pattern.sub(new_block, content, count=1)
    else:
        print(f"Warning: regex pattern did not match 'let rooms = [...]' in {html_path.name}")

    html_path.write_text(content, encoding="utf-8")

def sync_survey_artifacts(ss: gspread.Spreadsheet, repo_root: Path, push_pages: bool = False) -> Dict[str, Any]:
    """
    Exports room data and live U-value specifications from Google Sheets to CSV and index.html,
    and pushes to GitHub Pages if requested.
    """
    ws = ss.worksheet("2_Room_Heat_Loss")
    rooms = extract_existing_rooms(ws)
    if not rooms:
        return {"status": "error", "message": "No rooms found on 2_Room_Heat_Loss worksheet."}

    csv_path = repo_root / "room_by_room_heat_loss_survey.csv"
    export_rooms_to_csv(rooms, csv_path)

    specs = extract_specs_from_inputs(ss)
    index_html = repo_root / "index.html"
    update_html_specs_and_rooms(rooms, specs, index_html)

    git_status = "skipped"
    if push_pages:
        try:
            files_to_add = [
                str(csv_path.name),
                "index.html",
                "room_survey_tool.html"
            ]
            subprocess.run(["git", "add"] + files_to_add, cwd=repo_root, check=True)
            msg = f"Auto-sync room survey data ({len(rooms)} rooms) and live specs from 1_Inputs"
            diff_res = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=repo_root)
            if diff_res.returncode != 0:
                subprocess.run(["git", "commit", "-m", msg], cwd=repo_root, check=True)
                subprocess.run(["git", "push", "origin", "main"], cwd=repo_root, check=True)
                git_status = "pushed"
            else:
                git_status = "no_changes"
        except Exception as e:
            git_status = f"error: {e}"

    return {
        "status": "success",
        "num_rooms": len(rooms),
        "csv_path": str(csv_path),
        "html_updated": ["index.html"],
        "git_push": git_status
    }
