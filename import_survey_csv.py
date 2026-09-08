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

    rows_to_update = []
    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r_idx, row in enumerate(reader):
            sheet_row = 5 + r_idx
            # Update Dimensions, U-values, and Notes while keeping live Google Sheet formulas intact
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
            ach = row.get("ACH", "")
            notes = row.get("Notes", "")
            wall_type = row.get("Wall Type", "")
            rad_type = row.get("Radiator Type", "")
            rad_size = row.get("Rad Size", "")
            pipe_size = row.get("Pipework", "")
            
            combined_notes = f"{notes} | Construction: {wall_type} | Rad: {rad_type} {rad_size} ({pipe_size})"

            # Row columns:
            # Col E: Ti, Col F: Len, Col G: Wid, Col H: Area (formula), Col I: Ht, Col J: Vol (formula)
            # Col K: Ext Wall, Col L: U-Wall, Col M: Wall Loss (formula), Col N: Win Area, Col O: U-Win
            # Col P: Win Loss (formula), Col Q: Exp Floor, Col R: U-Floor, Col S: Fl Loss (formula)
            # Col T: Ceiling Area, Col U: U-Roof, Col V: Ceiling Loss (formula), Col W: ACH
            # Col X: Vent Loss (formula), Col Y: Total Loss (formula), Col Z: W/m2 (formula)
            # Col AA: Rad 45C (formula), Col AB: Boiler Rad (formula), Col AC: Emitter Type (formula)
            # Col AD: Notes
            ws.update(range_name=f"F{sheet_row}:G{sheet_row}", values=[[r_len, r_wid]])
            ws.update(range_name=f"I{sheet_row}:I{sheet_row}", values=[[r_ht]])
            ws.update(range_name=f"K{sheet_row}:L{sheet_row}", values=[[ext_wall, u_wall]])
            ws.update(range_name=f"N{sheet_row}:O{sheet_row}", values=[[win_area, u_win]])
            ws.update(range_name=f"Q{sheet_row}:R{sheet_row}", values=[[fl_area, u_fl]])
            ws.update(range_name=f"T{sheet_row}:U{sheet_row}", values=[[roof_area, u_roof]])
            ws.update(range_name=f"W{sheet_row}:W{sheet_row}", values=[[ach]])
            ws.update(range_name=f"AD{sheet_row}:AD{sheet_row}", values=[[combined_notes]])

    console.print(f"[bold green]✓ Successfully updated {tab_name} from survey CSV![/bold green]")

if __name__ == "__main__":
    main()
