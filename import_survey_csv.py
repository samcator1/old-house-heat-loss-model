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
    n_o_data = []
    q_r_data = []
    t_u_data = []
    q_dropdown_data = []
    notes_data = []

    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            r_len = row.get("Length (m)", "")
            r_wid = row.get("Width (m)", "")
            r_ht = row.get("Height (m)", "")
            ext_wall = row.get("Ext Wall (m)", "")
            u_wall = row.get("U-Wall", "")
            win_area = row.get("Window Area (m2)", "")
            u_win = row.get("U-Win", "")
            fl_area = row.get("Floor Area (m2)", "")
            u_fl = row.get("U-Floor", "")
            roof_area = row.get("Ceiling Area (m2)", "")
            u_roof = row.get("U-Ceiling", "")
            
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
            k_l_data.append([ext_wall, u_wall])
            n_o_data.append([win_area, u_win])
            q_r_data.append([fl_area, u_fl])
            t_u_data.append([roof_area, u_roof])
            if q_base:
                q_dropdown_data.append([q_base, q_chimney, q_win, q_floor, q_ceil])
            notes_data.append([combined_notes])

    num_rooms = len(f_g_data)
    if num_rooms == 0:
        console.print("[bold yellow]No rooms found in CSV file.[/bold yellow]")
        return

    end_row = 4 + num_rooms
    batch_updates = [
        {"range": f"F5:G{end_row}", "values": f_g_data},
        {"range": f"I5:I{end_row}", "values": i_data},
        {"range": f"K5:L{end_row}", "values": k_l_data},
        {"range": f"N5:O{end_row}", "values": n_o_data},
        {"range": f"Q5:R{end_row}", "values": q_r_data},
        {"range": f"T5:U{end_row}", "values": t_u_data},
        {"range": f"AI5:AI{end_row}", "values": notes_data}
    ]
    if len(q_dropdown_data) == num_rooms:
        batch_updates.append({"range": f"W5:AA{end_row}", "values": q_dropdown_data})

    ws.batch_update(batch_updates)
    console.print(f"[bold green]✓ Successfully updated {num_rooms} rooms in {tab_name} via batch API![/bold green]")

if __name__ == "__main__":
    main()
