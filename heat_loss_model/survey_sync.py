"""
Survey Synchronization Module.
Coordinates exporting live room data from 2_Room_Heat_Loss to:
1. room_by_room_heat_loss_survey.csv (CSV snapshot)
2. index.html and room_survey_tool.html (preloaded JavaScript rooms array)
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
from .tabs.room_tab import extract_existing_rooms, WINDOW_SPECIFICATIONS, CEILING_SPECIFICATIONS

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
    "U-Wall",
    "Window Area (m2)",
    "Window Specification",
    "Floor Area (m2)",
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

def get_window_u(spec: str) -> float:
    for item in WINDOW_SPECIFICATIONS:
        if item["label"] == spec:
            return item.get("u_value", item.get("u_val", 4.80))
    return 4.80

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

def export_rooms_to_csv(rooms: List[Dict[str, Any]], csv_path: Path) -> None:
    """Exports room list to CSV compatible with import_survey_csv.py."""
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        for rm in rooms:
            parsed = parse_room_notes(rm.get("notes", ""))
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
                rm.get("u_wall", 1.4),
                rm.get("win_area", 0.0),
                rm.get("win_spec", "Single Glazed (Historic Timber Sash / Casement)"),
                rm.get("fl_area", 0.0),
                rm.get("u_fl", 0.8),
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

def update_html_rooms(rooms: List[Dict[str, Any]], html_paths: List[Path]) -> None:
    """Updates the pre-loaded let rooms = [...] array in web survey HTML files."""
    js_objects = []
    for rm in rooms:
        parsed = parse_room_notes(rm.get("notes", ""))
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
            "uWall": rm.get("u_wall", 1.4),
            "winArea": rm.get("win_area", 0.0),
            "winSpec": win_spec,
            "uWin": get_window_u(win_spec),
            "flArea": rm.get("fl_area", 0.0),
            "uFl": rm.get("u_fl", 0.8),
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

    for p in html_paths:
        if not p.exists():
            continue
        content = p.read_text(encoding="utf-8")
        if pattern.search(content):
            updated_content = pattern.sub(new_block, content, count=1)
            p.write_text(updated_content, encoding="utf-8")
        else:
            print(f"Warning: regex pattern did not match 'let rooms = [...]' in {p.name}")

def sync_survey_artifacts(ss: gspread.Spreadsheet, repo_root: Path, push_pages: bool = False) -> Dict[str, Any]:
    """
    Coordinates exporting room data from Google Sheets to CSV and HTML survey tools,
    and pushes to GitHub Pages if requested.
    """
    ws = ss.worksheet("2_Room_Heat_Loss")
    rooms = extract_existing_rooms(ws)
    if not rooms:
        return {"status": "error", "message": "No rooms found on 2_Room_Heat_Loss worksheet."}

    csv_path = repo_root / "room_by_room_heat_loss_survey.csv"
    export_rooms_to_csv(rooms, csv_path)

    html_files = [
        repo_root / "index.html",
        repo_root / "room_survey_tool.html"
    ]
    update_html_rooms(rooms, html_files)

    git_status = "skipped"
    if push_pages:
        try:
            files_to_add = [str(csv_path.name)] + [str(h.name) for h in html_files if h.exists()]
            subprocess.run(["git", "add"] + files_to_add, cwd=repo_root, check=True)
            msg = f"Auto-sync room survey data ({len(rooms)} rooms) from live model"
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
        "html_updated": [str(h.name) for h in html_files if h.exists()],
        "git_push": git_status
    }
