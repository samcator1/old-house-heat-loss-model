#!/usr/bin/env python3
"""
Inspect an existing Google Sheets model.
Usage:
    python inspect_sheet.py
    python inspect_sheet.py --sheet-id <ID_OR_URL>
    python inspect_sheet.py --dump-json
"""

import os
import sys
import json
import argparse
from pathlib import Path
from dotenv import load_dotenv

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from rich.console import Console
from rich.table import Table
from rich.panel import Panel

sys.path.insert(0, str(Path(__file__).resolve().parent))
from heat_loss_model.auth import get_gspread_client

console = Console(force_terminal=False, no_color=False)

def parse_args():
    parser = argparse.ArgumentParser(description="Inspect Google Sheet structure and formulas.")
    parser.add_argument(
        "--sheet-id",
        type=str,
        default=os.getenv("SPREADSHEET_ID"),
        help="Google Spreadsheet ID or URL (defaults to SPREADSHEET_ID in .env)"
    )
    parser.add_argument(
        "--dump-json",
        action="store_true",
        help="Save raw data and formulas to a local JSON file for detailed offline analysis."
    )
    return parser.parse_args()

def inspect():
    load_dotenv()
    args = parse_args()

    sheet_id = args.sheet_id
    if not sheet_id:
        console.print(
            "[bold red]Error:[/bold red] No spreadsheet ID provided.\n"
            "Please provide `--sheet-id <ID_OR_URL>` or set `SPREADSHEET_ID` in `.env`."
        )
        sys.exit(1)

    console.print(f"[cyan]Connecting with service account credentials...[/cyan]")
    gc = get_gspread_client()

    console.print(f"[cyan]Opening spreadsheet:[/cyan] {sheet_id}")
    if sheet_id.startswith("http"):
        ss = gc.open_by_url(sheet_id)
    else:
        ss = gc.open_by_key(sheet_id)

    console.print(Panel.fit(
        f"[bold green]Title:[/bold green] {ss.title}\n"
        f"[bold green]URL:[/bold green] {ss.url}\n"
        f"[bold green]Worksheet Count:[/bold green] {len(ss.worksheets())}",
        title="Spreadsheet Overview",
        border_style="green"
    ))

    dump_data = {
        "title": ss.title,
        "url": ss.url,
        "id": ss.id,
        "sheets": {}
    }

    for ws in ss.worksheets():
        console.print(f"\n[bold yellow]=== Worksheet: {ws.title} (Rows: {ws.row_count}, Cols: {ws.col_count}) ===[/bold yellow]")
        
        # Get values and formulas
        values = ws.get_all_values()
        formulas = ws.get(value_render_option="FORMULA")
        
        dump_data["sheets"][ws.title] = {
            "values": values,
            "formulas": formulas
        }

        if not values:
            console.print("[dim]Empty worksheet[/dim]")
            continue

        # Display preview table (up to first 25 rows, 10 cols)
        max_rows = min(len(values), 25)
        max_cols = min(max(len(r) for r in values[:max_rows]), 10)

        table = Table(title=f"Preview: {ws.title} (Top {max_rows} rows)", border_style="dim")
        table.add_column("Row", justify="right", style="dim")
        for col_idx in range(max_cols):
            table.add_column(f"Col {col_idx+1}", style="white", max_width=30, overflow="fold")

        for r_idx in range(max_rows):
            row = values[r_idx]
            row_cells = [str(r_idx + 1)]
            for c_idx in range(max_cols):
                val = row[c_idx] if c_idx < len(row) else ""
                row_cells.append(str(val))
            table.add_row(*row_cells)

        console.print(table)

        # Formula summary
        formula_cells = []
        for r_idx, row in enumerate(formulas):
            for c_idx, cell in enumerate(row):
                if str(cell).startswith("="):
                    col_letter = chr(65 + c_idx) if c_idx < 26 else f"{chr(65 + c_idx // 26 - 1)}{chr(65 + c_idx % 26)}"
                    formula_cells.append(f"{col_letter}{r_idx+1}: {cell}")

        if formula_cells:
            console.print(f"[bold cyan]Found {len(formula_cells)} formulas in '{ws.title}'[/bold cyan]")
            console.print("[dim]Sample formulas:[/dim] " + ", ".join(formula_cells[:5]))
        else:
            console.print(f"[dim]No formulas found in '{ws.title}' (all static values)[/dim]")

    if args.dump_json:
        out_path = Path("sheet_dump.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(dump_data, f, indent=2)
        console.print(f"\n[bold green]Dumped full sheet content and formulas to {out_path}[/bold green]")

if __name__ == "__main__":
    inspect()
