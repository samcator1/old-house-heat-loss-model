"""
Helper functions for batch spreadsheet styling, formatting, and Google Sheets API v4 requests.
"""

from typing import Dict, Any, List, Optional

def create_clear_formatting_request(
    sheet_id: int,
    max_rows: int = 120,
    max_cols: int = 35
) -> Dict[str, Any]:
    """Resets cell formatting across the entire worksheet."""
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": max_rows,
                "startColumnIndex": 0,
                "endColumnIndex": max_cols
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "textFormat": {
                        "foregroundColor": {"red": 0.09, "green": 0.16, "blue": 0.24},
                        "bold": False,
                        "italic": False,
                        "fontSize": 10
                    },
                    "horizontalAlignment": "LEFT",
                    "verticalAlignment": "MIDDLE",
                    "wrapStrategy": "OVERFLOW_CELL"
                }
            },
            "fields": (
                "userEnteredFormat.backgroundColor,"
                "userEnteredFormat.textFormat,"
                "userEnteredFormat.horizontalAlignment,"
                "userEnteredFormat.verticalAlignment,"
                "userEnteredFormat.wrapStrategy"
            )
        }
    }

def create_unmerge_cells_request(
    sheet_id: int,
    max_rows: int = 120,
    max_cols: int = 35
) -> Dict[str, Any]:
    """Unmerges all cells in the given range."""
    return {
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": max_rows,
                "startColumnIndex": 0,
                "endColumnIndex": max_cols
            }
        }
    }

def create_repeat_cell_request(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    bg_color: Optional[Dict[str, float]] = None,
    font_color: Optional[Dict[str, float]] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_size: Optional[int] = None,
    number_format: Optional[str] = None,
    align: Optional[str] = None,
    valign: Optional[str] = "MIDDLE",
    wrap_strategy: Optional[str] = None
) -> Dict[str, Any]:
    """Creates a Google Sheets API v4 repeatCell request dictionary."""
    fields = []
    cell_format: Dict[str, Any] = {}
    
    if bg_color is not None:
        cell_format["backgroundColor"] = bg_color
        fields.append("userEnteredFormat.backgroundColor")
        
    text_format: Dict[str, Any] = {}
    if font_color is not None:
        text_format["foregroundColor"] = font_color
    if bold is not None:
        text_format["bold"] = bold
    if italic is not None:
        text_format["italic"] = italic
    if font_size is not None:
        text_format["fontSize"] = font_size
    if text_format:
        cell_format["textFormat"] = text_format
        fields.append("userEnteredFormat.textFormat")
        
    if number_format is not None:
        if "£" in number_format:
            fmt_type = "CURRENCY"
        elif "%" in number_format:
            fmt_type = "PERCENT"
        else:
            fmt_type = "NUMBER"
        cell_format["numberFormat"] = {
            "type": fmt_type,
            "pattern": number_format
        }
        fields.append("userEnteredFormat.numberFormat")
        
    if align is not None:
        cell_format["horizontalAlignment"] = align
        fields.append("userEnteredFormat.horizontalAlignment")

    if valign is not None:
        cell_format["verticalAlignment"] = valign
        fields.append("userEnteredFormat.verticalAlignment")

    if wrap_strategy is not None:
        cell_format["wrapStrategy"] = wrap_strategy
        fields.append("userEnteredFormat.wrapStrategy")

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "cell": {
                "userEnteredFormat": cell_format
            },
            "fields": ",".join(fields)
        }
    }

def create_freeze_pane_request(sheet_id: int, frozen_rows: int = 3, frozen_cols: int = 0) -> Dict[str, Any]:
    """Freezes header rows or columns."""
    grid_properties: Dict[str, Any] = {}
    fields = []
    if frozen_rows > 0:
        grid_properties["frozenRowCount"] = frozen_rows
        fields.append("gridProperties.frozenRowCount")
    if frozen_cols > 0:
        grid_properties["frozenColumnCount"] = frozen_cols
        fields.append("gridProperties.frozenColumnCount")

    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": grid_properties
            },
            "fields": ",".join(fields)
        }
    }

def create_set_column_width_request(sheet_id: int, start_col: int, end_col: int, pixel_size: int) -> Dict[str, Any]:
    """Sets explicit column width."""
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col
            },
            "properties": {
                "pixelSize": pixel_size
            },
            "fields": "pixelSize"
        }
    }

def create_merge_cells_request(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    merge_type: str = "MERGE_ALL"
) -> Dict[str, Any]:
    """Merges cells in range."""
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "mergeType": merge_type
        }
    }

def create_auto_resize_request(sheet_id: int, num_columns: int) -> Dict[str, Any]:
    """Auto-resizes columns to fit content."""
    return {
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": num_columns
            }
        }
    }

def create_data_validation_request(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    options: List[str],
    show_custom_ui: bool = True,
    strict: bool = False
) -> Dict[str, Any]:
    """Creates a setDataValidation request to render in-cell dropdown chips in Google Sheets."""
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": opt} for opt in options]
                },
                "showCustomUi": show_custom_ui,
                "strict": strict
            }
        }
    }

