#!/usr/bin/env python3
"""
Imports room survey CSV data exported from the on-site survey tool
and synchronizes it directly into the 5_Room_Heat_Loss Google Sheets tab.
Usage:
    python import_survey_csv.py [path_to_csv]
"""

import os
import sys
import csv
import argparse
from pathlib import Path
from dotenv import load_dotenv

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from rich.console import Console
from rich.panel import Panel

sys.path.insert(0, str(Path(__file__).resolve().parent))
from heat_loss_model.auth import get_gspread_client
from heat_loss_model.schema import RoomCol

console = Console(force_terminal=False, no_color=False)

def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="Import room survey CSV to Google Sheets.")
    parser.add_argument("csv_path", nargs="?", default="room_by_room_heat_loss_survey.csv", help="Path to exported survey CSV")
    parser.add_argument("--sheet-id", default=os.getenv("SPREADSHEET_ID"), help="Spreadsheet ID")
    args = parser.parse_args()

    csv_file = Path(args.csv_path)
    if not csv_file.exists():
        console.print(f"[bold red]Error:[/bold red] Survey file '{csv_file}' not found.")
        sys.exit(1)

    console.print(Panel.fit(f"[bold cyan]Importing Survey Data from {csv_file.name}...[/bold cyan]"))

    gc = get_gspread_client()
    ss = gc.open_by_key(args.sheet_id)
    tab_name = "2_Room_Heat_Loss" if "2_Room_Heat_Loss" in [w.title for w in ss.worksheets()] else "5_Room_Heat_Loss"
    ws = ss.worksheet(tab_name)

    f_g_data = []
    i_data = []
    k_l_data = []
    o_data = []
    p_data = []
    s_t_data = []
    w_data = []
    x_data = []
    q_dropdown_data = []
    notes_data = []

    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            r_len = row.get("Length (m)", "")
            r_wid = row.get("Width (m)", "")
            r_ht = row.get("Height (m)", "")
            ext_wall = row.get("Ext Wall (m)", "")
            wall_spec = row.get("Wall Specification", "")
            if not wall_spec:
                u_wall_str = row.get("U-Wall", "")
                if u_wall_str:
                    try:
                        u_flt = float(u_wall_str)
                        from heat_loss_model.tabs.room_tab import map_u_to_wall_spec
                        wall_spec = map_u_to_wall_spec(u_flt)
                    except Exception:
                        wall_spec = 'Solid Brick: 18" / 450mm (Georgian Facade)'
                else:
                    wall_spec = 'Solid Brick: 18" / 450mm (Georgian Facade)'

            win_area = row.get("Window Area (m2)", "")
            win_spec = row.get("Window Specification", "")
            if not win_spec:
                win_spec = "Single Glazed (Historic Timber Sash / Casement)"
            fl_area = row.get("Floor Area (m2)", "")
            floor_spec = row.get("Floor Specification", "")
            if not floor_spec:
                u_fl_str = row.get("U-Floor", "")
                floor_lvl = row.get("Floor Level", row.get("Floor", ""))
                zone_name = row.get("Zone / Wing", row.get("Zone", ""))
                if u_fl_str:
                    try:
                        u_flt = float(u_fl_str)
                        from heat_loss_model.tabs.room_tab import map_u_to_floor_spec
                        floor_spec = map_u_to_floor_spec(u_flt, floor_lvl, zone_name)
                    except Exception:
                        floor_spec = "Suspended Timber: Bare Boards Over Void"
                else:
                    floor_spec = "Suspended Timber: Bare Boards Over Void"

            roof_area = row.get("Ceiling Area (m2)", "")
            ceil_spec = row.get("Ceiling Specification", "")
            if not ceil_spec:
                ceil_spec = "Intermediate Floor (Heated Space Above)"
            
            q_base = row.get("Base Construction", "")
            q_chimney = row.get("Chimney Flue", "")
            q_win = row.get("Windows Doors", "")
            q_floor = row.get("Floor Boundary", "")
            q_ceil = row.get("Ceiling Boundary", "")

            notes = row.get("Notes", "")
            wall_type = row.get("Wall Type", "")
            rad_type = row.get("Radiator Type", "")
            rad_size = row.get("Rad Size", "")
            pipe_size = row.get("Pipework", "")
            
            combined_notes = f"{notes} | Construction: {wall_type} | Rad: {rad_type} {rad_size} ({pipe_size})"

            f_g_data.append([r_len, r_wid])
            i_data.append([r_ht])
            k_l_data.append([ext_wall, wall_spec])
            o_data.append([win_area])
            p_data.append([win_spec])
            s_t_data.append([fl_area, floor_spec])
            w_data.append([roof_area])
            x_data.append([ceil_spec])
            if q_base:
                q_dropdown_data.append([q_base, q_chimney, q_win, q_floor, q_ceil])
            notes_data.append([combined_notes])

    num_rooms = len(f_g_data)
    if num_rooms == 0:
        console.print("[bold yellow]No rooms found in CSV file.[/bold yellow]")
        return

    end_row = 4 + num_rooms
    batch_updates = [
        {"range": f"{RoomCol.LENGTH}5:{RoomCol.WIDTH}{end_row}", "values": f_g_data},
        {"range": f"{RoomCol.HEIGHT}5:{RoomCol.HEIGHT}{end_row}", "values": i_data},
        {"range": f"{RoomCol.EXT_WALL_L}5:{RoomCol.WALL_SPEC}{end_row}", "values": k_l_data},
        {"range": f"{RoomCol.WIN_AREA}5:{RoomCol.WIN_AREA}{end_row}", "values": o_data},
        {"range": f"{RoomCol.WIN_SPEC}5:{RoomCol.WIN_SPEC}{end_row}", "values": p_data},
        {"range": f"{RoomCol.FL_AREA}5:{RoomCol.FLOOR_SPEC}{end_row}", "values": s_t_data},
        {"range": f"{RoomCol.ROOF_AREA}5:{RoomCol.ROOF_AREA}{end_row}", "values": w_data},
        {"range": f"{RoomCol.CEIL_SPEC}5:{RoomCol.CEIL_SPEC}{end_row}", "values": x_data},
        {"range": f"{RoomCol.NOTES}5:{RoomCol.NOTES}{end_row}", "values": notes_data}
    ]
    if len(q_dropdown_data) == num_rooms:
        batch_updates.append({"range": f"{RoomCol.Q_BASE}5:{RoomCol.Q_CEIL}{end_row}", "values": q_dropdown_data})

    ws.batch_update(batch_updates)
    console.print(f"[bold green]✓ Successfully updated {num_rooms} rooms in {tab_name} via batch API![/bold green]")

if __name__ == "__main__":
    main()
