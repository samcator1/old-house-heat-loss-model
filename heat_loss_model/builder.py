"""
Main Heat Loss Model Builder & Orchestrator.
Coordinates input preservation, calculation tab generation, complete format clearing,
clean batch styling, and safe legacy sheet preservation.
"""

from typing import Optional, Dict, Any, List
import gspread
from .inputs_manager import InputsManager
from .tabs.dashboard_tab import build_dashboard_tab
from .tabs.fabric_tab import build_fabric_tab
from .tabs.dhw_pool_tab import build_dhw_pool_tab
from .tabs.systems_tab import build_systems_tab
from .tabs.room_tab import build_room_tab
from .formatting import create_clear_formatting_request, create_unmerge_cells_request

class HeatLossModelBuilder:
    """Orchestrates building and syncing the complete heat loss model in Google Sheets."""

    def __init__(self, gc: gspread.Client, spreadsheet_id: str):
        self.gc = gc
        self.spreadsheet_id = spreadsheet_id
        self.ss: Optional[gspread.Spreadsheet] = None

    def open_spreadsheet(self) -> gspread.Spreadsheet:
        """Opens the spreadsheet by ID or URL."""
        if self.spreadsheet_id.startswith("http"):
            self.ss = self.gc.open_by_url(self.spreadsheet_id)
        else:
            self.ss = self.gc.open_by_key(self.spreadsheet_id)
        return self.ss

    def migrate_sheet_titles(self) -> None:
        """Safely renames and migrates sheet titles to preserve history and set up clean structure."""
        if not self.ss:
            self.open_spreadsheet()
        ss = self.ss

        existing = {w.title: w for w in ss.worksheets()}

        # 1. Protect original 'Heating'
        if "Heating" in existing and "_Legacy_Heating" not in existing:
            try:
                existing["Heating"].update_title("_Legacy_Heating")
                print("✓ Preserved original 'Heating' tab as '_Legacy_Heating'")
            except Exception as e:
                print(f"Note on legacy sheet rename: {e}")

        # 2. Rename 2_Building_Heat_Loss -> _Archive_Wing_Heat_Loss
        if "2_Building_Heat_Loss" in existing and "_Archive_Wing_Heat_Loss" not in existing:
            try:
                existing["2_Building_Heat_Loss"].update_title("_Archive_Wing_Heat_Loss")
                print("✓ Archived '2_Building_Heat_Loss' to '_Archive_Wing_Heat_Loss'")
            except Exception as e:
                print(f"Note on archive wing rename: {e}")

        # 3. Rename 5_Room_Heat_Loss -> 2_Room_Heat_Loss
        if "5_Room_Heat_Loss" in existing and "2_Room_Heat_Loss" not in existing:
            try:
                existing["5_Room_Heat_Loss"].update_title("2_Room_Heat_Loss")
                print("✓ Promoted '5_Room_Heat_Loss' to master driver '2_Room_Heat_Loss'")
            except Exception as e:
                print(f"Note on room tab rename: {e}")

    def sync(self, preserve_inputs: bool = True) -> Dict[str, Any]:
        """
        Synchronizes the upgraded modular model to Google Sheets:
        1. Migrates sheet titles safely (preserves legacy, archives wing calc, promotes room schedule).
        2. Reads existing user modifications on '1_Inputs' non-destructively.
        3. Updates all calculation grids with 100% native Google Sheets formulas.
        4. Clears previous formatting and applies crisp modern theme styling in batch.
        5. Reorders tabs logically.
        """
        if not self.ss:
            self.open_spreadsheet()

        ss = self.ss

        # Step 1: Migrate sheet titles
        self.migrate_sheet_titles()

        # Step 2: Read existing inputs & merge
        inputs_mgr = InputsManager(ss)
        merged_inputs = inputs_mgr.get_merged_inputs(preserve_inputs=preserve_inputs)

        # Step 3: Render Tabs in dependency order
        ws_inputs, fmt_inputs = inputs_mgr.render_inputs_tab(merged_inputs)
        ws_room, fmt_room = build_room_tab(ss)
        ws_dhw, fmt_dhw = build_dhw_pool_tab(ss)
        ws_systems, fmt_systems = build_systems_tab(ss)
        ws_dashboard, fmt_dashboard = build_dashboard_tab(ss)
        ws_archive_wing, fmt_archive_wing = build_fabric_tab(ss)

        active_worksheets = [ws_dashboard, ws_inputs, ws_room, ws_dhw, ws_systems, ws_archive_wing]

        # Step 4: Clear leftover formatting and unmerge on active tabs
        all_formatting_requests: List[Dict[str, Any]] = []
        for ws in active_worksheets:
            all_formatting_requests.append(create_unmerge_cells_request(ws.id, max_rows=100, max_cols=38))
            all_formatting_requests.append(create_clear_formatting_request(ws.id, max_rows=100, max_cols=38))

        # Step 5: Add targeted theme formatting
        all_formatting_requests.extend(fmt_dashboard)
        all_formatting_requests.extend(fmt_inputs)
        all_formatting_requests.extend(fmt_room)
        all_formatting_requests.extend(fmt_dhw)
        all_formatting_requests.extend(fmt_systems)
        all_formatting_requests.extend(fmt_archive_wing)

        # Step 6: Execute batch formatting in 1 API call
        if all_formatting_requests:
            try:
                ss.batch_update({"requests": all_formatting_requests})
            except Exception as e:
                print(f"Note on batch styling: {e}")

        # Step 7: Reorder tabs logically
        desired_order = [
            "0_Executive_Dashboard",
            "1_Inputs",
            "2_Room_Heat_Loss",
            "3_DHW_and_Pool",
            "4_Heating_and_Renewables",
            "_Archive_Wing_Heat_Loss",
            "_Legacy_Heating"
        ]

        try:
            current_worksheets = ss.worksheets()
            ws_by_title = {w.title: w for w in current_worksheets}
            reorder_requests = []
            current_idx = 0

            for title in desired_order:
                if title in ws_by_title:
                    w = ws_by_title[title]
                    reorder_requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": w.id,
                                "index": current_idx
                            },
                            "fields": "index"
                        }
                    })
                    current_idx += 1

            # Any remaining user-created sheets placed after
            for w in current_worksheets:
                if w.title not in desired_order:
                    reorder_requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": w.id,
                                "index": current_idx
                            },
                            "fields": "index"
                        }
                    })
                    current_idx += 1

            if reorder_requests:
                ss.batch_update({"requests": reorder_requests})
        except Exception as e:
            print(f"Note on sheet reordering: {e}")

        return {
            "spreadsheet_title": ss.title,
            "spreadsheet_url": ss.url,
            "inputs_preserved": preserve_inputs
        }
