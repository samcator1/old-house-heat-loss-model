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

    room_type_data = []
    f_g_data = []
    i_data = []
    k_l_data = []
    o_data = []
    p_data = []
    door_data = []
    floor_spec_data = []
    ceil_spec_data = []
    chimney_data = []
    mech_vent_data = []
    emitter_type_data = []
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
            
            door_area = row.get("Door Area (m2)", "")
            door_spec = row.get("Door Specification", "")
            if not door_spec:
                door_spec = "No External Door (Internal Boundary Only)"

            chimney = row.get("Chimney / Fireplace", row.get("Chimney Flue", "No Chimney / Permanently Sealed"))
            mech_vent = row.get("Mechanical Ventilation", "None (Natural Infiltration Only)")
            emitter_type = row.get("Planned Emitter Type", "Type 22 (Double Convector)")

            notes = row.get("Notes", "")
            wall_type = row.get("Wall Type", "")
            rad_type = row.get("Radiator Type", "")
            rad_size = row.get("Rad Size", "")
            pipe_size = row.get("Pipework", "")
            
            combined_notes = f"{notes} | Construction: {wall_type} | Rad: {rad_type} {rad_size} ({pipe_size})"

            r_type = row.get("Room Type", "")
            room_type_data.append([r_type])
            f_g_data.append([r_len, r_wid])
            i_data.append([r_ht])
            k_l_data.append([ext_wall, wall_spec])
            o_data.append([win_area])
            p_data.append([win_spec])
            door_data.append([door_area, door_spec])
            floor_spec_data.append([floor_spec])
            ceil_spec_data.append([ceil_spec])
            chimney_data.append([chimney])
            mech_vent_data.append([mech_vent])
            emitter_type_data.append([emitter_type])
            notes_data.append([combined_notes])

    num_rooms = len(f_g_data)
    if num_rooms == 0:
        console.print("[bold yellow]No rooms found in CSV file.[/bold yellow]")
        return

    end_row = 4 + num_rooms
    batch_updates = []
    if any(r[0] for r in room_type_data):
        batch_updates.append({"range": f"{RoomCol.ROOM_TYPE}5:{RoomCol.ROOM_TYPE}{end_row}", "values": room_type_data})
    batch_updates.extend([
        {"range": f"{RoomCol.LENGTH}5:{RoomCol.WIDTH}{end_row}", "values": f_g_data},
        {"range": f"{RoomCol.HEIGHT}5:{RoomCol.HEIGHT}{end_row}", "values": i_data},
        {"range": f"{RoomCol.EXT_WALL_L}5:{RoomCol.WALL_SPEC}{end_row}", "values": k_l_data},
        {"range": f"{RoomCol.WIN_AREA}5:{RoomCol.WIN_AREA}{end_row}", "values": o_data},
        {"range": f"{RoomCol.WIN_SPEC}5:{RoomCol.WIN_SPEC}{end_row}", "values": p_data},
        {"range": f"{RoomCol.DOOR_AREA}5:{RoomCol.DOOR_SPEC}{end_row}", "values": door_data},
        {"range": f"{RoomCol.FLOOR_SPEC}5:{RoomCol.FLOOR_SPEC}{end_row}", "values": floor_spec_data},
        {"range": f"{RoomCol.CEIL_SPEC}5:{RoomCol.CEIL_SPEC}{end_row}", "values": ceil_spec_data},
        {"range": f"{RoomCol.CHIMNEY}5:{RoomCol.CHIMNEY}{end_row}", "values": chimney_data},
        {"range": f"{RoomCol.MECH_VENT}5:{RoomCol.MECH_VENT}{end_row}", "values": mech_vent_data},
        {"range": f"{RoomCol.EMITTER_TYPE}5:{RoomCol.EMITTER_TYPE}{end_row}", "values": emitter_type_data},
        {"range": f"{RoomCol.NOTES}5:{RoomCol.NOTES}{end_row}", "values": notes_data}
    ])

    ws.batch_update(batch_updates)
    console.print(f"[bold green]✓ Successfully updated {num_rooms} rooms in {tab_name} via batch API![/bold green]")

if __name__ == "__main__":
    main()
