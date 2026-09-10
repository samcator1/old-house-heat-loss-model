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
from heat_loss_model.survey_sync import sync_survey_artifacts

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
        "--push-pages",
        action="store_true",
        help="Automatically commit and push updated room survey tool artifacts to GitHub Pages."
    )
    parser.add_argument(
        "--skip-survey-sync",
        action="store_true",
        help="Skip updating room_by_room_heat_loss_survey.csv and web tool HTML files."
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

    survey_res = None
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

        if not args.skip_survey_sync:
            task3 = progress.add_task("[cyan]Exporting survey CSV & updating mobile web tool...", total=None)
            try:
                repo_root = Path(__file__).resolve().parent
                survey_res = sync_survey_artifacts(builder.ss, repo_root, push_pages=args.push_pages)
                progress.update(task3, description="[green]Survey artifacts updated successfully!")
            except Exception as e:
                console.print(f"[bold yellow]Survey Tool Sync Note:[/bold yellow] {e}")

    # Output Success Summary
    console.print(f"\n[bold green]✓ Successfully updated spreadsheet:[/bold green] [bold white]{result['spreadsheet_title']}[/bold white]")
    console.print(f"URL: {result['spreadsheet_url']}\n")

    num_rooms = result.get("num_rooms", 24)
    table = Table(title="Synchronized Sheets & Formula Architecture", border_style="dim")
    table.add_column("Tab Name", style="bold cyan")
    table.add_column("Type", style="green")
    table.add_column("Description")

    table.add_row("0_Executive_Dashboard", "Formula View", f"KPI cards, heating options comparison, bottom-up wing breakdown ({num_rooms} rooms) & solar balance")
    table.add_row(
        "1_Inputs", 
        "User Control" if result["inputs_preserved"] else "Reset to Default", 
        f"Central parameters & assumptions ({'Preserved user edits' if result['inputs_preserved'] else 'Reset to baseline defaults'})"
    )
    table.add_row("2_Room_Heat_Loss", "Master Dynamic Formula", f"{num_rooms}-room schedule (with live user inputs preserved), fabric & vent loss, 45°C low-flow radiator sizing")
    table.add_row("3_DHW_and_Pool", "Dynamic Formula", "800L DHW storage, recharge rate, secondary circulation, Legionella cycle & pool thermal demand")
    table.add_row("4_Heating_and_Renewables", "Dynamic Formula", "GSHP vs ASHP vs Oil Boiler, Solar PV & battery load shifting, and smart tariff economics")
    table.add_row("_Archive_Wing_Heat_Loss", "Protected Archive", "Preserved macro 6-zone approximation with corrected ground ΔT (superseded by room schedule)")
    table.add_row("_Legacy_Heating", "Protected Archive", "Original user sheet preserved untouched as a historical reference")

    console.print(table)

    if survey_res and survey_res.get("status") == "success":
        console.print(f"[bold green]✓ Mobile Survey Tool Sync:[/bold green] Exported {survey_res['num_rooms']} rooms to [cyan]room_by_room_heat_loss_survey.csv[/cyan] and updated [cyan]index.html[/cyan] / [cyan]room_survey_tool.html[/cyan].")
        if survey_res.get("git_push") == "pushed":
            console.print("[bold green]✓ GitHub Pages:[/bold green] Changes committed and pushed to [link=https://samcator1.github.io/old-house-heat-loss-model/]GitHub Pages[/link]!")
        elif survey_res.get("git_push") == "no_changes":
            console.print("[dim]GitHub Pages: Already up to date (no diff).[/dim]")

    console.print("\n[dim]All calculations are live Google Sheets formulas. You can edit any soft blue cell in '1_Inputs' or room dimensions/dropdowns in '2_Room_Heat_Loss' directly in Google Sheets without losing your data on future syncs.[/dim]\n")

if __name__ == "__main__":
    main()
