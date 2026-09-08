#!/usr/bin/env python3
"""
CLI Tool for synchronizing and vibe-coding the Old House Heat Loss & Energy Model in Google Sheets.
Usage:
    python sync_model.py
    python sync_model.py --sheet-id <ID_OR_URL>
    python sync_model.py --reset-inputs
"""

import os
import sys
import argparse
from pathlib import Path
from dotenv import load_dotenv

# Ensure safe UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn

# Ensure heat_loss_model package is in python path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from heat_loss_model.auth import get_gspread_client
from heat_loss_model.builder import HeatLossModelBuilder

console = Console(force_terminal=False, no_color=False)

def parse_args():
    parser = argparse.ArgumentParser(
        description="Synchronize the Old House Heat Loss Model to Google Sheets with dynamic formulas."
    )
    parser.add_argument(
        "--sheet-id",
        type=str,
        default=os.getenv("SPREADSHEET_ID"),
        help="Google Spreadsheet ID or URL (defaults to SPREADSHEET_ID from .env)"
    )
    parser.add_argument(
        "--reset-inputs",
        action="store_true",
        help="Force reset all inputs on 1_Inputs to baseline defaults instead of preserving custom edits."
    )
    parser.add_argument(
        "--service-account",
        type=str,
        default=os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE"),
        help="Path to service account JSON credentials file."
    )
    return parser.parse_args()

def main():
    load_dotenv()
    args = parse_args()

    console.print(Panel.fit(
        "[bold cyan]Old House Heat Loss & Energy Model Engine[/bold cyan]\n"
        "[dim]Dynamic Google Sheets Formula Generator & Vibe-Coding Sync Tool[/dim]",
        border_style="cyan"
    ))

    sheet_id = args.sheet_id
    if not sheet_id:
        console.print(
            "[bold red]Error:[/bold red] No spreadsheet ID provided.\n"
            "Please set [green]SPREADSHEET_ID[/green] in your [yellow].env[/yellow] file "
            "or pass it via [yellow]--sheet-id <ID_OR_URL>[/yellow]."
        )
        sys.exit(1)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        transient=True
    ) as progress:
        task1 = progress.add_task("[cyan]Connecting to Google Sheets API...", total=None)
        try:
            gc = get_gspread_client(service_account_path=args.service_account)
            progress.update(task1, description="[green]Authenticated successfully!")
        except Exception as e:
            console.print(f"[bold red]Authentication Error:[/bold red] {e}")
            sys.exit(1)

        task2 = progress.add_task("[cyan]Syncing spreadsheet model & formulas...", total=None)
        try:
            builder = HeatLossModelBuilder(gc, sheet_id)
            preserve_inputs = not args.reset_inputs
            result = builder.sync(preserve_inputs=preserve_inputs)
            progress.update(task2, description="[green]Model synced successfully!")
        except Exception as e:
            err_msg = str(e)
            if hasattr(e, "__cause__") and e.__cause__:
                err_msg = str(e.__cause__)
            console.print(f"[bold red]Sync Error:[/bold red] {err_msg}")
            sys.exit(1)

    # Output Success Summary
    console.print(f"\n[bold green]✓ Successfully updated spreadsheet:[/bold green] [bold white]{result['spreadsheet_title']}[/bold white]")
    console.print(f"URL: {result['spreadsheet_url']}\n")

    table = Table(title="Synchronized Sheets & Formula Architecture", border_style="dim")
    table.add_column("Tab Name", style="bold cyan")
    table.add_column("Type", style="green")
    table.add_column("Description")

    table.add_row("0_Executive_Dashboard", "Formula View", "KPI cards, heating options comparison, zone breakdown & solar balance")
    table.add_row(
        "1_Inputs", 
        "User Control" if result["inputs_preserved"] else "Reset to Default", 
        f"Central parameters & assumptions ({'Preserved user edits' if result['inputs_preserved'] else 'Reset to baseline defaults'})"
    )
    table.add_row("2_Building_Heat_Loss", "Dynamic Formula", "Zone geometry, transparent element-by-element fabric losses (corrected ground ΔT) & infiltration")
    table.add_row("3_DHW_and_Pool", "Dynamic Formula", "800L DHW storage, recharge rate, secondary circulation, Legionella cycle & pool thermal demand")
    table.add_row("4_Heating_and_Renewables", "Dynamic Formula", "GSHP vs ASHP vs Oil Boiler, Solar PV & battery load shifting, and smart tariff economics")
    table.add_row("5_Room_Heat_Loss", "Dynamic Formula", "23-room schedule (10 GF + 13 FF), fabric & vent loss, and 45°C low-flow radiator sizing")
    table.add_row("_Legacy_Heating", "Protected Archive", "Original user sheet preserved untouched as a historical reference")

    console.print(table)
    console.print("\n[dim]All calculations are live Google Sheets formulas. You can edit any soft blue cell in '1_Inputs' directly in Google Sheets.[/dim]\n")

if __name__ == "__main__":
    main()
