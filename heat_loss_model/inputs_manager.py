"""
Inputs Manager for Old House Heat Loss Model.
Reads existing user inputs from Google Sheets non-destructively, merges with defaults,
and renders the '1_Inputs' control center worksheet with clean formatting.
"""

from typing import Dict, Any, List, Tuple
import gspread
from .config import THEME, FORMATS
from .formatting import (
    create_repeat_cell_request,
    create_set_column_width_request,
    create_freeze_pane_request,
    create_merge_cells_request
)

# Canonical parameter definitions with exact row numbers on '1_Inputs'
INPUT_DEFINITIONS: List[Dict[str, Any]] = [
    # Section 1: Weather & Design Temperatures
    {"row": 4, "key": "design_inside_temp", "section": "1. Weather & Design Temperatures", "label": "Indoor Design Temperature", "default": 20.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "CIBSE Guide A comfort design benchmark"},
    {"row": 5, "key": "design_outside_temp", "section": "1. Weather & Design Temperatures", "label": "Outdoor Design Temperature", "default": -4.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "External 99% winter peak design temp"},
    {"row": 6, "key": "ground_temp", "section": "1. Weather & Design Temperatures", "label": "Ground Temperature", "default": 10.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "Sub-floor ground temp (~10°C year-round)"},
    {"row": 7, "key": "heating_degree_days", "section": "1. Weather & Design Temperatures", "label": "Heating Degree Days (HDD)", "default": 1400, "unit": "K·days", "fmt": FORMATS["INTEGER"], "notes": "Base temp degree-days for space heating estimation"},
    {"row": 8, "key": "calibration_factor", "section": "1. Weather & Design Temperatures", "label": "Heating Usage Factor (f_usage)", "default": 0.80, "unit": "fraction", "fmt": FORMATS["DECIMAL_2"], "notes": "Accounts for internal gains, solar gains, setback & diversity"},

    # Section 2: Ventilation & Infiltration Rates
    {"row": 11, "key": "air_constant", "section": "2. Ventilation & Infiltration Rates", "label": "Air Volumetric Heat Capacity", "default": 0.33, "unit": "W/m³K", "fmt": FORMATS["DECIMAL_2"], "notes": "Standard air constant (density x specific heat / 3600)"},
    {"row": 12, "key": "ach_peak_old", "section": "2. Ventilation & Infiltration Rates", "label": "Peak Infiltration Rate - Old House", "default": 1.80, "unit": "ACH", "fmt": FORMATS["ACH"], "notes": "Windward peak ACH for emitter/radiator sizing"},
    {"row": 13, "key": "ach_avg_old", "section": "2. Ventilation & Infiltration Rates", "label": "Average Infiltration Rate - Old House", "default": 1.40, "unit": "ACH", "fmt": FORMATS["ACH"], "notes": "Whole-house diversified ACH for plant/heat pump sizing"},
    {"row": 14, "key": "ach_peak_new", "section": "2. Ventilation & Infiltration Rates", "label": "Peak Infiltration Rate - New Build", "default": 0.50, "unit": "ACH", "fmt": FORMATS["ACH"], "notes": "Modern air permeability standard (~5 m³/h·m²)"},
    {"row": 15, "key": "ach_avg_new", "section": "2. Ventilation & Infiltration Rates", "label": "Average Infiltration Rate - New Build", "default": 0.35, "unit": "ACH", "fmt": FORMATS["ACH"], "notes": "Whole-house average ACH for new extension"},
    {"row": 16, "key": "intermittent_boost_hp", "section": "2. Ventilation & Infiltration Rates", "label": "Intermittent Boost Factor - Heat Pump", "default": 0.00, "unit": "%", "fmt": FORMATS["PERCENT"], "notes": "0% for continuous 24/7 low-temp weather compensated heating"},
    {"row": 17, "key": "intermittent_boost_boiler", "section": "2. Ventilation & Infiltration Rates", "label": "Intermittent Boost Factor - Oil Boiler", "default": 0.20, "unit": "%", "fmt": FORMATS["PERCENT"], "notes": "20% standard margin for morning/evening intermittent setback"},

    # Section 3: Domestic Hot Water (DHW) System
    {"row": 20, "key": "dhw_occupants", "section": "3. Domestic Hot Water (DHW) System", "label": "Number of Occupants", "default": 3, "unit": "people", "fmt": FORMATS["INTEGER"], "notes": "Design building occupancy"},
    {"row": 21, "key": "dhw_litres_per_person", "section": "3. Domestic Hot Water (DHW) System", "label": "Hot Water Allowance per Person", "default": 50.0, "unit": "L/person/day", "fmt": FORMATS["DECIMAL_1"], "notes": "Standard UK domestic consumption"},
    {"row": 22, "key": "dhw_stored_temp", "section": "3. Domestic Hot Water (DHW) System", "label": "DHW Storage Temperature", "default": 55.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "Optimal heat pump storage temperature"},
    {"row": 23, "key": "dhw_cold_temp", "section": "3. Domestic Hot Water (DHW) System", "label": "Mains Cold Feed Temperature", "default": 10.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "UK average annual mains cold water"},
    {"row": 24, "key": "dhw_vessel_litres", "section": "3. Domestic Hot Water (DHW) System", "label": "DHW Cylinder / Vessel Capacity", "default": 800, "unit": "Litres", "fmt": FORMATS["INTEGER"], "notes": "Baystar recommendation to ensure smooth HP run cycles"},
    {"row": 25, "key": "dhw_standing_loss_pct", "section": "3. Domestic Hot Water (DHW) System", "label": "Cylinder Standing Loss Allowance", "default": 0.12, "unit": "%", "fmt": FORMATS["PERCENT"], "notes": "Insulated vessel daily standing loss allowance"},
    {"row": 26, "key": "dhw_secondary_loop", "section": "3. Domestic Hot Water (DHW) System", "label": "Secondary Circulation Loop Present?", "default": "Yes", "unit": "Yes/No", "fmt": None, "notes": "Pumped loop for instant hot water in large building"},
    {"row": 27, "key": "dhw_secondary_loss_kwh_day", "section": "3. Domestic Hot Water (DHW) System", "label": "Secondary Loop Standing Loss", "default": 3.5, "unit": "kWh/day", "fmt": FORMATS["DECIMAL_1"], "notes": "Distribution pipework heat losses (~3.5 kWh/day)"},
    {"row": 28, "key": "dhw_legionella_boost_kwh_yr", "section": "3. Domestic Hot Water (DHW) System", "label": "Legionella Weekly Pasteurization", "default": 250, "unit": "kWh/yr", "fmt": FORMATS["INTEGER"], "notes": "Weekly 60-65°C immersion pasteurization cycle"},
    {"row": 29, "key": "dhw_recharge_hours", "section": "3. Domestic Hot Water (DHW) System", "label": "Target Vessel Recharge Time", "default": 1.5, "unit": "Hours", "fmt": FORMATS["DECIMAL_1"], "notes": "Time allowed to reheat full cylinder from cold"},

    # Section 4: Swimming Pool Thermal Parameters
    {"row": 32, "key": "pool_length", "section": "4. Swimming Pool Thermal Parameters", "label": "Pool Length", "default": 10.0, "unit": "m", "fmt": FORMATS["DECIMAL_1"], "notes": "Pool length"},
    {"row": 33, "key": "pool_width", "section": "4. Swimming Pool Thermal Parameters", "label": "Pool Width", "default": 5.0, "unit": "m", "fmt": FORMATS["DECIMAL_1"], "notes": "Pool width"},
    {"row": 34, "key": "pool_depth", "section": "4. Swimming Pool Thermal Parameters", "label": "Average Pool Depth", "default": 3.0, "unit": "m", "fmt": FORMATS["DECIMAL_1"], "notes": "Pool depth"},
    {"row": 35, "key": "pool_temp", "section": "4. Swimming Pool Thermal Parameters", "label": "Target Pool Water Temperature", "default": 24.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "Standard outdoor heated pool setpoint"},
    {"row": 36, "key": "pool_season_days", "section": "4. Swimming Pool Thermal Parameters", "label": "Pool Heating Season Length", "default": 90, "unit": "Days", "fmt": FORMATS["INTEGER"], "notes": "May to September heating window (adjustable)"},
    {"row": 37, "key": "pool_covered_loss_rate", "section": "4. Swimming Pool Thermal Parameters", "label": "Covered Surface Heat Loss Rate", "default": 0.060, "unit": "kW/m²", "fmt": FORMATS["DECIMAL_2"], "notes": "5-7 kW total heat loss with thermal pool cover"},
    {"row": 38, "key": "pool_uncovered_loss_rate", "section": "4. Swimming Pool Thermal Parameters", "label": "Uncovered Surface Heat Loss Rate", "default": 0.150, "unit": "kW/m²", "fmt": FORMATS["DECIMAL_2"], "notes": "Peak uncovered evaporation and surface loss rate"},
    {"row": 39, "key": "pool_covered_hours", "section": "4. Swimming Pool Thermal Parameters", "label": "Daily Covered Hours", "default": 20.0, "unit": "Hours/day", "fmt": FORMATS["DECIMAL_1"], "notes": "Hours pool cover remains on per day"},

    # Section 5: Heating System Efficiencies & SCOPs
    {"row": 42, "key": "boiler_efficiency", "section": "5. Heating System Efficiencies & SCOPs", "label": "Oil Boiler Seasonal Efficiency", "default": 0.85, "unit": "fraction", "fmt": FORMATS["PERCENT"], "notes": "Modern condensing oil boiler benchmark"},
    {"row": 43, "key": "kerosene_kwh_per_litre", "section": "5. Heating System Efficiencies & SCOPs", "label": "Kerosene Energy Content", "default": 10.0, "unit": "kWh/Litre", "fmt": FORMATS["DECIMAL_1"], "notes": "Standard net calorific value"},
    {"row": 44, "key": "gshp_scop_space", "section": "5. Heating System Efficiencies & SCOPs", "label": "GSHP Seasonal COP - Space (45°C flow)", "default": 3.60, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "Ground source heat pump at medium flow"},
    {"row": 45, "key": "gshp_scop_dhw", "section": "5. Heating System Efficiencies & SCOPs", "label": "GSHP Seasonal COP - DHW (55°C flow)", "default": 2.80, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "High-temp hot water generation efficiency"},
    {"row": 46, "key": "gshp_scop_pool", "section": "5. Heating System Efficiencies & SCOPs", "label": "GSHP Seasonal COP - Pool (35°C flow)", "default": 4.20, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "Low-temp pool heat exchanger efficiency"},
    {"row": 47, "key": "ashp_scop_space", "section": "5. Heating System Efficiencies & SCOPs", "label": "ASHP Seasonal COP - Space (45°C flow)", "default": 3.00, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "Air source heat pump annual average benchmark"},
    {"row": 48, "key": "ashp_scop_dhw", "section": "5. Heating System Efficiencies & SCOPs", "label": "ASHP Seasonal COP - DHW (55°C flow)", "default": 2.40, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "Air source heat pump DHW efficiency"},
    {"row": 49, "key": "ashp_scop_pool", "section": "5. Heating System Efficiencies & SCOPs", "label": "ASHP Seasonal COP - Pool (35°C flow)", "default": 3.50, "unit": "SCOP", "fmt": FORMATS["DECIMAL_2"], "notes": "Air source pool heating efficiency"},

    # Section 6: Energy Tariffs, Solar PV & Battery Storage
    {"row": 52, "key": "price_heating_oil", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Price of Heating Oil", "default": 0.65, "unit": "£/Litre", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "Current kerosene bulk delivery price"},
    {"row": 53, "key": "tariff_flat_electric", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Flat Electricity Tariff", "default": 0.25, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "Standard variable single-rate electricity"},
    {"row": 54, "key": "tariff_smart_offpeak", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Smart Tariff - Off-Peak Rate", "default": 0.15, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "Cheapest overnight & mid-day slots (8 hrs/day)"},
    {"row": 55, "key": "tariff_smart_day", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Smart Tariff - Day / Standard Rate", "default": 0.27, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "Daytime rate (13 hrs/day)"},
    {"row": 56, "key": "tariff_smart_peak", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Smart Tariff - Evening Peak Rate", "default": 0.40, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "4pm - 7pm peak demand window (3 hrs/day)"},
    {"row": 57, "key": "tariff_export_seg", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Smart Export Guarantee (SEG)", "default": 0.12, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP_DEC"], "notes": "Feed-in export price for surplus solar generation"},
    {"row": 58, "key": "solar_pv_kwp", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Solar PV Installed Capacity", "default": 37.0, "unit": "kWp", "fmt": FORMATS["DECIMAL_1"], "notes": "Installed solar photovoltaic peak power"},
    {"row": 59, "key": "solar_specific_yield", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Solar Specific Yield", "default": 900, "unit": "kWh/kWp/yr", "fmt": FORMATS["INTEGER"], "notes": "UK East Anglia annual solar yield"},
    {"row": 60, "key": "battery_capacity_kwh", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Battery Storage Usable Capacity", "default": 20.0, "unit": "kWh", "fmt": FORMATS["DECIMAL_1"], "notes": "Home battery storage capacity for load shifting"},
    {"row": 61, "key": "domestic_base_electric", "section": "6. Energy Tariffs, Solar PV & Battery Storage", "label": "Domestic Non-Heating Electricity Load", "default": 8000, "unit": "kWh/yr", "fmt": FORMATS["INTEGER"], "notes": "Baseline household appliance/lighting electricity"},

    # Section 7: Capital Equipment Cost Rates & Carbon Factors
    {"row": 64, "key": "capex_gshp_per_kw", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "GSHP System Turnkey Cost Rate", "default": 2444, "unit": "£/kW", "fmt": FORMATS["CURRENCY_GBP"], "notes": "Includes ground loops / boreholes and plantroom"},
    {"row": 65, "key": "capex_ashp_per_kw", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "ASHP System Turnkey Cost Rate", "default": 1400, "unit": "£/kW", "fmt": FORMATS["CURRENCY_GBP"], "notes": "Turnkey air source heat pump installation"},
    {"row": 66, "key": "capex_boiler_per_kw", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "Oil Boiler Replacement Cost Rate", "default": 455, "unit": "£/kW", "fmt": FORMATS["CURRENCY_GBP"], "notes": "High-output commercial/domestic boiler replacement"},
    {"row": 67, "key": "capex_solar_per_kwp", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "Solar PV Cost Rate", "default": 1150, "unit": "£/kWp", "fmt": FORMATS["CURRENCY_GBP"], "notes": "Turnkey commercial roof-mount solar PV"},
    {"row": 68, "key": "capex_battery_per_kwh", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "Battery Storage Cost Rate", "default": 500, "unit": "£/kWh", "fmt": FORMATS["CURRENCY_GBP"], "notes": "Installed modular battery storage"},
    {"row": 69, "key": "carbon_grid_electricity", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "Grid Electricity Carbon Intensity", "default": 0.150, "unit": "kg CO2/kWh", "fmt": FORMATS["DECIMAL_2"], "notes": "UK grid average operational emissions"},
    {"row": 70, "key": "carbon_heating_oil", "section": "7. Capital Equipment Cost Rates & Carbon Factors", "label": "Heating Oil Carbon Intensity", "default": 0.298, "unit": "kg CO2/kWh", "fmt": FORMATS["DECIMAL_2"], "notes": "Kerosene burning direct carbon emissions"},

    # Section 8: Infiltration Questionnaire Scoring (CIBSE Guide A & BS EN 12831)
    {"row": 76, "key": "ach_base_new", "section": "8. Infiltration Questionnaire Scoring", "label": "New Build (Cavity / Insulated)", "default": 0.30, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Modern insulated cavity wall and airtight membrane"},
    {"row": 77, "key": "ach_base_reno", "section": "8. Infiltration Questionnaire Scoring", "label": "Renovated Historic (Draught-proofed)", "default": 0.60, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Solid masonry with retrofitted draught-proofing and sealed penetrations"},
    {"row": 78, "key": "ach_base_solid", "section": "8. Infiltration Questionnaire Scoring", "label": "Standard Historic (Solid Masonry)", "default": 0.80, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional uninsulated solid brick / stone without airtightness layer"},
    {"row": 79, "key": "ach_base_exposed", "section": "8. Infiltration Questionnaire Scoring", "label": "Exposed Historic (High Wind / Exposed Site)", "default": 1.00, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Hilltop or coastal windward facade with high stack pressure"},
    {"row": 80, "key": "ach_chimney_none", "section": "8. Infiltration Questionnaire Scoring", "label": "No Chimney / Permanently Sealed", "default": 0.00, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "No open flue or chimney has been permanently capped/sealed"},
    {"row": 81, "key": "ach_chimney_stove", "section": "8. Infiltration Questionnaire Scoring", "label": "Room-Sealed Stove with Ext Air", "default": 0.05, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Closed woodburner with dedicated direct external combustion air duct"},
    {"row": 82, "key": "ach_chimney_balloon", "section": "8. Infiltration Questionnaire Scoring", "label": "Flue Damper / Chimney Balloon Fitted", "default": 0.15, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Open chimney fitted with inflatable throat balloon or mechanical damper"},
    {"row": 83, "key": "ach_chimney_open", "section": "8. Infiltration Questionnaire Scoring", "label": "Open Fireplace (Unsealed Flue)", "default": 0.60, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Unrestricted open chimney flue driving continuous warm air stack extraction"},
    {"row": 84, "key": "ach_win_modern", "section": "8. Infiltration Questionnaire Scoring", "label": "Modern High-Performance (Compression Gaskets)", "default": 0.00, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Double/triple glazed casement with intact rubber compression gaskets"},
    {"row": 85, "key": "ach_win_sec", "section": "8. Infiltration Questionnaire Scoring", "label": "Secondary Glazing Fitted", "default": 0.10, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal secondary glazing panes providing sealed air buffer"},
    {"row": 86, "key": "ach_win_brush", "section": "8. Infiltration Questionnaire Scoring", "label": "Retrofitted Brush / Pile Draught Seals", "default": 0.15, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Timber sash or casement retrofitted with routed brush pile carrier strips"},
    {"row": 87, "key": "ach_win_loose", "section": "8. Infiltration Questionnaire Scoring", "label": "Original Loose Sash / Casement (Undraughted)", "default": 0.40, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional loose-fitting timber sash rattles in track with air gaps"},
    {"row": 88, "key": "ach_floor_slab", "section": "8. Infiltration Questionnaire Scoring", "label": "Solid Concrete Slab / Insulated Floor", "default": 0.00, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Continuous concrete slab or insulated screed without air leakage"},
    {"row": 89, "key": "ach_floor_sealed", "section": "8. Infiltration Questionnaire Scoring", "label": "Suspended Timber (Sealed / Carpet & Underlay)", "default": 0.10, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Suspended joists with tongue-and-groove boards and heavy underlay"},
    {"row": 90, "key": "ach_floor_draughty", "section": "8. Infiltration Questionnaire Scoring", "label": "Suspended Timber (Unsealed Boards over Cold Void)", "default": 0.35, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Air bricks beneath unsealed floorboards with noticeable crawlspace draughts"},
    {"row": 91, "key": "ach_ceil_inter", "section": "8. Infiltration Questionnaire Scoring", "label": "Intermediate Floor (Heated Space Above)", "default": 0.00, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal ceiling with actively heated room above"},
    {"row": 92, "key": "ach_ceil_sealed", "section": "8. Infiltration Questionnaire Scoring", "label": "Insulated Loft (Sealed Plaster & Sealed Hatch)", "default": 0.05, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Airtight plaster ceiling with draught-stripped insulated loft trapdoor"},
    {"row": 93, "key": "ach_ceil_downlights", "section": "8. Infiltration Questionnaire Scoring", "label": "Attic Downlights / Unsealed Loft Hatch", "default": 0.20, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Recessed downlight penetrations or loose loft hatch venting into cold roof"},
    {"row": 94, "key": "ach_ceil_sloping", "section": "8. Infiltration Questionnaire Scoring", "label": "Sloping Roof / Exposed Eaves / Thatched Ridge", "default": 0.30, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic uncounter-battened thatched ridge or open soffit eaves"}
]

SECTION_HEADERS = [
    {"row": 3, "title": "1. WEATHER & DESIGN TEMPERATURES"},
    {"row": 10, "title": "2. VENTILATION & INFILTRATION RATES"},
    {"row": 19, "title": "3. DOMESTIC HOT WATER (DHW) SYSTEM"},
    {"row": 31, "title": "4. SWIMMING POOL THERMAL PARAMETERS"},
    {"row": 41, "title": "5. HEATING SYSTEM EFFICIENCIES & SCOPS"},
    {"row": 51, "title": "6. ENERGY TARIFFS, SOLAR PV & BATTERY STORAGE"},
    {"row": 63, "title": "7. CAPITAL EQUIPMENT COST RATES & CARBON FACTORS"},
    {"row": 74, "title": "8. INFILTRATION QUESTIONNAIRE SCORING & AIR CHANGE RATE (ACH) PENALTIES"}
]

class InputsManager:
    """Manages reading, merging, and rendering the 1_Inputs tab."""

    def __init__(self, ss: gspread.Spreadsheet):
        self.ss = ss

    def get_merged_inputs(self, preserve_inputs: bool = True) -> Dict[str, Any]:
        """
        Reads existing inputs from 1_Inputs worksheet if present.
        If preserve_inputs is True, user-entered values in Col C overwrite defaults.
        """
        merged = {item["key"]: item["default"] for item in INPUT_DEFINITIONS}
        if not preserve_inputs:
            return merged

        try:
            ws = self.ss.worksheet("1_Inputs")
            raw_values = ws.get_all_values()
        except gspread.WorksheetNotFound:
            return merged
        except Exception:
            return merged

        row_map = {item["row"]: item for item in INPUT_DEFINITIONS}

        for r_idx, row in enumerate(raw_values):
            row_num = r_idx + 1
            if row_num in row_map and len(row) >= 3:
                val = row[2].strip()  # Column C (0-indexed: 2)
                item = row_map[row_num]
                if val != "":
                    try:
                        clean_val = val.replace("£", "").replace("%", "").replace("°C", "").replace("W/m³K", "").replace("ACH", "").replace(",", "").strip()
                        if isinstance(item["default"], int):
                            merged[item["key"]] = int(float(clean_val))
                        elif isinstance(item["default"], float):
                            flt = float(clean_val)
                            if item.get("fmt") in [FORMATS["PERCENT"], FORMATS["PERCENT_INT"]] or "%" in val:
                                if flt > 1.0:
                                    flt = flt / 100.0
                                    # Handle case if it was compounded
                                    while flt > 1.0:
                                        flt = flt / 100.0
                            merged[item["key"]] = flt
                        else:
                            merged[item["key"]] = clean_val
                    except Exception:
                        pass
        return merged

    def render_inputs_tab(self, merged_inputs: Dict[str, Any]) -> Tuple[gspread.Worksheet, List[Dict[str, Any]]]:
        """
        Writes data into 1_Inputs and generates clean batch formatting requests.
        """
        try:
            ws = self.ss.worksheet("1_Inputs")
        except gspread.WorksheetNotFound:
            ws = self.ss.add_worksheet(title="1_Inputs", rows=105, cols=6)

        total_rows = 96
        grid: List[List[str]] = [["" for _ in range(5)] for _ in range(total_rows)]

        # Banner & Column Headers
        grid[0][0] = "1_Inputs: Central Parameters & Assumptions Control Center"
        grid[1][0] = "Blue cells are user-adjustable inputs. All calculations dynamically update across the workbook."
        grid[2][0] = "Section"
        grid[2][1] = "Parameter Description"
        grid[2][2] = "Value"
        grid[2][3] = "Unit"
        grid[2][4] = "Notes & Technical Guidance"

        # Section Headers
        for sh in SECTION_HEADERS:
            r_idx = sh["row"] - 1
            grid[r_idx][0] = sh["title"]

        # Fill Parameters
        for item in INPUT_DEFINITIONS:
            r_idx = item["row"] - 1
            grid[r_idx][0] = item["section"]
            grid[r_idx][1] = item["label"]
            val = merged_inputs.get(item["key"], item["default"])
            grid[r_idx][2] = str(val)
            grid[r_idx][3] = item["unit"]
            grid[r_idx][4] = item["notes"]

        # Write in single batch
        ws.update(values=grid, range_name=f"A1:E{total_rows}", value_input_option="USER_ENTERED")

        # Prepare formatting requests
        fmt_reqs: List[Dict[str, Any]] = []

        # Freeze header rows
        fmt_reqs.append(create_freeze_pane_request(ws.id, frozen_rows=3, frozen_cols=0))

        # Column widths
        fmt_reqs.append(create_set_column_width_request(ws.id, 0, 1, 260))  # Section
        fmt_reqs.append(create_set_column_width_request(ws.id, 1, 2, 280))  # Description
        fmt_reqs.append(create_set_column_width_request(ws.id, 2, 3, 110))  # Value (Input)
        fmt_reqs.append(create_set_column_width_request(ws.id, 3, 4, 100))  # Unit
        fmt_reqs.append(create_set_column_width_request(ws.id, 4, 5, 450))  # Notes

        # Banner styling
        fmt_reqs.append(create_merge_cells_request(ws.id, 0, 1, 0, 5))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, 0, 1, 0, 5,
            bg_color=THEME["PRIMARY_HEADER_BG"],
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=12,
            align="LEFT"
        ))

        fmt_reqs.append(create_merge_cells_request(ws.id, 1, 2, 0, 5))
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, 1, 2, 0, 5,
            bg_color=THEME["ZEBRA_BG"],
            font_color=THEME["MUTED_TEXT"],
            italic=True,
            font_size=9,
            align="LEFT"
        ))

        # Table Header (Row 3)
        fmt_reqs.append(create_repeat_cell_request(
            ws.id, 2, 3, 0, 5,
            bg_color=THEME["SECONDARY_HEADER_BG"],
            font_color=THEME["HEADER_TEXT"],
            bold=True,
            font_size=10,
            align="CENTER"
        ))

        # Section Headers styling
        for sh in SECTION_HEADERS:
            r_idx = sh["row"] - 1
            fmt_reqs.append(create_merge_cells_request(ws.id, r_idx, r_idx + 1, 0, 5))
            fmt_reqs.append(create_repeat_cell_request(
                ws.id, r_idx, r_idx + 1, 0, 5,
                bg_color=THEME["SECTION_HEADER_BG"],
                font_color=THEME["HEADER_TEXT"],
                bold=True,
                font_size=10,
                align="LEFT"
            ))

        # Format input rows
        for item in INPUT_DEFINITIONS:
            r_idx = item["row"] - 1
            # Section & Description
            fmt_reqs.append(create_repeat_cell_request(
                ws.id, r_idx, r_idx + 1, 0, 2,
                bg_color=THEME["ZEBRA_BG"],
                font_color=THEME["DARK_TEXT"],
                font_size=9,
                align="LEFT"
            ))
            # Editable Value in Column C (Soft Blue)
            fmt_reqs.append(create_repeat_cell_request(
                ws.id, r_idx, r_idx + 1, 2, 3,
                bg_color=THEME["INPUT_BG"],
                font_color={"red": 0.05, "green": 0.20, "blue": 0.45},
                bold=True,
                font_size=10,
                align="RIGHT",
                number_format=item.get("fmt")
            ))
            # Unit & Notes
            fmt_reqs.append(create_repeat_cell_request(
                ws.id, r_idx, r_idx + 1, 3, 4,
                bg_color=THEME["ZEBRA_BG"],
                font_color=THEME["MUTED_TEXT"],
                font_size=9,
                align="CENTER"
            ))
            fmt_reqs.append(create_repeat_cell_request(
                ws.id, r_idx, r_idx + 1, 4, 5,
                bg_color={"red": 1.0, "green": 1.0, "blue": 1.0},
                font_color=THEME["MUTED_TEXT"],
                italic=True,
                font_size=9,
                align="LEFT"
            ))

        return ws, fmt_reqs
