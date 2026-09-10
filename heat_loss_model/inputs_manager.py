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
    {"row": 94, "key": "ach_ceil_sloping", "section": "8. Infiltration Questionnaire Scoring", "label": "Sloping Roof / Exposed Eaves / Thatched Ridge", "default": 0.30, "unit": "ACH", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic uncounter-battened thatched ridge or open soffit eaves"},

    # Section 9: Window & Glazing U-Value Specifications (CIBSE Guide A & Historic England)
    {"row": 98, "key": "u_win_single_timber", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Single Glazed (Historic Timber Sash / Casement)", "default": 4.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Uninsulated traditional timber sash, 3-4mm float glass (CIBSE Guide A Table 3.29)"},
    {"row": 99, "key": "u_win_single_metal", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Single Glazed (Metal / Crittall / Stone Mullion)", "default": 5.60, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Single glazed in uninsulated steel frame or direct to stone mullion"},
    {"row": 100, "key": "u_win_single_shutters", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Single Glazed + Heavy Thermal Curtains / Working Shutters", "default": 3.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional solid interior wooden shutters closed or heavy thermal interlined curtains"},
    {"row": 101, "key": "u_win_single_sec_std", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Single Glazed + Standard Secondary Glazing", "default": 2.60, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal secondary glazing pane with >20mm sealed cavity, clear float glass"},
    {"row": 102, "key": "u_win_single_sec_lowe", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Single Glazed + Low-E Secondary Glazing", "default": 1.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Secondary glazing pane with pyrolytic Low-E coating, draught-sealed perimeter"},
    {"row": 103, "key": "u_win_double_slim", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Heritage Slimline Double Glazing (Timber Frame)", "default": 1.90, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Ultra-slim vacuum or Krypton units (6-8mm) fitting traditional historic glazing rebates"},
    {"row": 104, "key": "u_win_double_early", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Early Standard Double Glazing (Pre-2002, Air Gap)", "default": 2.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "1st-generation aluminium spacer double glazing without Low-E coating"},
    {"row": 105, "key": "u_win_double_modern", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Modern Double Glazing (Argon, Low-E, Warm Edge)", "default": 1.40, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Current UK Building Regulations Part L standard domestic double glazing"},
    {"row": 106, "key": "u_win_triple_modern", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Modern High-Performance Triple Glazing", "default": 0.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Passivhaus standard triple glazing with two Low-E coats and warm edge spacer"},
    {"row": 107, "key": "u_win_roof_lantern", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "Glazed Roof Lantern / Sloping Skylight", "default": 1.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Sloping or horizontal glazed orientation with increased internal convective transfer"},
    {"row": 108, "key": "u_win_none", "section": "9. Window & Glazing Thermal Transmittance (U-Values)", "label": "No External Glazing (Enclosed / Blind Room)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal corridors, enclosed cloakrooms, or windowless plantrooms"},

    # Section 10: Ceiling & Roof U-Value Specifications (CIBSE Guide A & Historic England)
    {"row": 113, "key": "u_ceil_intermediate", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Intermediate Floor (Heated Space Above)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal ceiling below heated room on upper floor; zero envelope transmission loss"},
    {"row": 114, "key": "u_ceil_loft_uninsulated", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Uninsulated Loft (Bare Joists, No Quilt)", "default": 2.30, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "CIBSE Guide A Table 3.19: lath/plaster ceiling below cold ventilated attic without insulation"},
    {"row": 115, "key": "u_ceil_thatched", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Historic Thatched Roof (Straw / Reed Thatch)", "default": 0.30, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional 300-450mm straw/water reed thatch naturally provides continuous thermal insulation"},
    {"row": 116, "key": "u_ceil_loft_50mm", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Older Loft Insulation: 50mm (Pre-1980 Quilt)", "default": 0.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Thin historic mineral wool between ceiling joists, often degraded or compressed"},
    {"row": 117, "key": "u_ceil_loft_100mm", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Older Loft Insulation: 100mm (1980s-1990s Standard)", "default": 0.40, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Single layer mineral wool quilt filling joist depth only"},
    {"row": 118, "key": "u_ceil_loft_200mm", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Moderate Loft Insulation: 150-200mm (Pre-2002 Standard)", "default": 0.22, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Joists filled plus partial top-up layer, typical existing retrofitted loft"},
    {"row": 119, "key": "u_ceil_loft_270mm", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Modern Building Regs Loft: 270-300mm (Mineral Wool)", "default": 0.16, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Current Building Regulations Part L: 100mm between joists + 170mm cross-laid quilt"},
    {"row": 120, "key": "u_ceil_loft_350mm", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "High-Performance Deep Loft: 350-400mm (Quilt / Cellulose)", "default": 0.11, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Deep quilt or blown cellulose insulation achieving near-Passivhaus thermal resistance"},
    {"row": 121, "key": "u_ceil_sloping_uninsulated", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)", "default": 2.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Sloping roof with lath & plaster directly under slate/tile rafters, no loft void (e.g. Lean-to scullery)"},
    {"row": 122, "key": "u_ceil_sloping_insulated", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Insulated Sloping Rafters (50-100mm PIR / Woodfibre)", "default": 0.35, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Sloping cathedral ceiling retrofitted with rigid PIR or breathable woodfibre board between rafters"},
    {"row": 123, "key": "u_ceil_flatroof_modern", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Modern Insulated Warm Flat Roof Extension", "default": 0.18, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Flat timber deck roof with 120-150mm rigid PIR insulation above deck (Building Regs compliant)"},
    {"row": 124, "key": "u_ceil_roof_lantern", "section": "10. Ceiling & Roof Thermal Transmittance (U-Values)", "label": "Glazed Roof Lantern / Sloping Skylight (Orangery)", "default": 1.50, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Extensive sloping double glazed roof lantern over Orangery or garden room"},

    # Section 11: Wall Construction U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.3)
    {"row": 129, "key": "u_wall_brick_9in", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Brick: 9\" / 225mm (Uninsulated)", "default": 2.10, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "CIBSE Guide A Table 3.3: 1-brick thick solid wall with interior plaster (e.g. uninsulated lean-to/scullery)"},
    {"row": 130, "key": "u_wall_brick_13in", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Brick: 13.5\" / 330mm (Uninsulated)", "default": 1.70, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "1.5-brick thick traditional solid masonry wall with plaster"},
    {"row": 131, "key": "u_wall_brick_18in", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Brick: 18\" / 450mm (Georgian Facade)", "default": 1.40, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "2-brick thick substantial Georgian external facade with internal lath and plaster finish"},
    {"row": 132, "key": "u_wall_stone_rubble", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Stone: 450-500mm Sandstone / Limestone", "default": 1.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic England / CIBSE Guide A: Solid dressed or rubble stone wall (e.g. thatched gable wing)"},
    {"row": 133, "key": "u_wall_stone_granite", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Stone: 500-600mm Dense Granite / Whinstone", "default": 2.20, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Dense crystalline igneous or metamorphic stone without cavity"},
    {"row": 134, "key": "u_wall_timber_frame", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Historic Timber Frame (Wattle & Daub / Nogging)", "default": 1.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional exposed timber box-frame with wattle/daub or brick nogging infill"},
    {"row": 135, "key": "u_wall_cob", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Cob / Earth Construction (500mm+)", "default": 0.90, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional thick monolithic chalk/clay/straw cob wall with breathable lime render"},
    {"row": 136, "key": "u_wall_cavity_uninsulated", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Uninsulated Cavity Wall (Pre-1976)", "default": 1.50, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Early cavity construction (brick/block or brick/brick) with uninsulated 50mm air gap"},
    {"row": 137, "key": "u_wall_cavity_retrofill", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Retrofilled Cavity Wall (Blown Mineral / Bead)", "default": 0.50, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Existing masonry cavity injected with bonded EPS beads or blown mineral fibre"},
    {"row": 138, "key": "u_wall_cavity_partial", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Partial-Fill Cavity Wall (1980s-1990s)", "default": 0.45, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "25-50mm partial-fill insulation board retained against inner leaf"},
    {"row": 139, "key": "u_wall_cavity_regs", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Modern Building Regs Cavity (100mm PIR / Full Fill)", "default": 0.18, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Current Building Regulations Part L new build standard (e.g. New Build wing)"},
    {"row": 140, "key": "u_wall_cavity_passiv", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "High-Performance New Build (150mm+ PIR / Passivhaus)", "default": 0.12, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Ultra-low-energy extension specification with thick high-performance insulation"},
    {"row": 141, "key": "u_wall_iwi_50mm", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Wall + 50mm Breathable Woodfibre IWI", "default": 0.55, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic England recommended internal breathable woodfibre / cork lime retrofit"},
    {"row": 142, "key": "u_wall_iwi_100mm", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Wall + 100mm Woodfibre / PIR IWI", "default": 0.28, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Deep internal wall insulation with intelligent vapour control membrane"},
    {"row": 143, "key": "u_wall_ewi_100mm", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Solid Wall + 100mm External Wall Insulation (EWI)", "default": 0.25, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "External insulation with breathable render finish on non-heritage elevations"},
    {"row": 144, "key": "u_wall_party", "section": "11. Wall Construction Thermal Transmittance (U-Values)", "label": "Party Wall / Heated Boundary (No Heat Loss)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal partition or party wall adjoining heated space; zero envelope transmission loss"},

    # Section 12: Ground & Floor Construction U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.23 / 3.24)
    {"row": 148, "key": "u_floor_intermediate", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Intermediate Floor (Heated Space Below)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Internal intermediate floor with actively heated space beneath; zero envelope transmission loss"},
    {"row": 149, "key": "u_floor_modern_regs", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Modern Building Regs Insulated Slab (100mm PIR / Full Fill)", "default": 0.15, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Current Building Regulations Part L ground-bearing concrete slab with 100-120mm continuous PIR"},
    {"row": 150, "key": "u_floor_passivhaus", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "High-Performance Insulated Slab (150mm+ PIR / Passivhaus)", "default": 0.10, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Fully wrapped raft or slab with 150-200mm high-density insulation achieving near-zero ground heat loss"},
    {"row": 151, "key": "u_floor_timber_100mm", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Suspended Timber: Insulated (100-150mm PIR between Joists)", "default": 0.18, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Retrofitted timber floor with 100-150mm rigid PIR or dense woodfibre supported on breathable membrane"},
    {"row": 152, "key": "u_floor_timber_50mm", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Suspended Timber: Insulated (50mm Quilt / Board)", "default": 0.35, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Partial retrofit with 50mm insulation batts between joists over crawlspace void"},
    {"row": 153, "key": "u_floor_conc_50mm", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Solid Concrete: Moderate Insulation (50mm PIR / 1990s)", "default": 0.35, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "1990s to early 2000s standard slab with 50mm insulation beneath concrete or screed"},
    {"row": 154, "key": "u_floor_conc_perimeter", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Solid Concrete: Perimeter Insulation (1980s Standard)", "default": 0.45, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "1980s standard ground slab with vertical perimeter edge insulation only"},
    {"row": 155, "key": "u_floor_timber_carpet", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Suspended Timber: Carpet & Underlay over Uninsulated Void", "default": 0.60, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic suspended timber joists with tongue-and-groove boards, heavy underlay and fitted carpet"},
    {"row": 156, "key": "u_floor_timber_bare", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Suspended Timber: Bare Boards over Cold Uninsulated Void", "default": 0.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "CIBSE Guide A Table 3.24: Uninsulated suspended timber floorboards over ventilated sub-floor void"},
    {"row": 157, "key": "u_floor_conc_pre1976", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Solid Concrete Ground Slab (Uninsulated Pre-1976)", "default": 0.80, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional uninsulated solid concrete ground slab cast directly on hardcore/blinding"},
    {"row": 158, "key": "u_floor_solid_flags", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth", "default": 1.10, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Historic uninsulated flagstones, brick pavers, or quarry tiles bedded directly on ground / lime bed"},
    {"row": 159, "key": "u_floor_exposed_uninsulated", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Exposed Floor over Cold Undercroft / Garage (Uninsulated)", "default": 1.20, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "First floor cantilever or room over open driveway, unheated garage, or unheated archway without insulation"},
    {"row": 160, "key": "u_floor_exposed_insulated", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Exposed Floor over Cold Undercroft / Garage (100mm PIR Insulated)", "default": 0.20, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Cantilevered exposed floor retrofitted with 100mm rigid insulation board and soffit board"},
    {"row": 161, "key": "u_floor_party", "section": "12. Ground & Floor Construction Thermal Transmittance (U-Values)", "label": "Unheated Adjoining Boundary (Zero Transmission Loss)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Floor separating heated space from fully conditioned adjoining residential unit"},

    # Section 13: External Door Construction U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.28)
    {"row": 165, "key": "u_door_timber_uninsulated", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Solid Timber Door (Uninsulated / Historic Plank)", "default": 3.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional solid oak/pine external door (25-45mm) without weather seals"},
    {"row": 166, "key": "u_door_timber_draughtproof", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Solid Timber Door (Draught-Stripped / Heavy Ledge)", "default": 2.40, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Heavy historic plank door with compression draught-proofing seals"},
    {"row": 167, "key": "u_door_glazed_single", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Part-Glazed External Door (Single Glazed)", "default": 3.60, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Traditional timber entrance door with single-glazed upper lights"},
    {"row": 168, "key": "u_door_glazed_double", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Part-Glazed External Door (Double Glazed)", "default": 2.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Timber or aluminium framed door with 16mm argon double-glazed unit"},
    {"row": 169, "key": "u_door_composite_modern", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Modern High-Performance / Composite Insulated Door", "default": 1.20, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Part L compliant foam-filled composite or timber core door with thermal break"},
    {"row": 170, "key": "u_door_french_bifold", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "Modern French / Bi-fold Glazed Doors (Low-E Double)", "default": 1.40, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Modern high-performance multi-pane patio doors with thermal breaks"},
    {"row": 171, "key": "u_door_none", "section": "13. External Door Construction Thermal Transmittance (U-Values)", "label": "No External Door (Internal Boundary Only)", "default": 0.00, "unit": "W/m²K", "fmt": FORMATS["DECIMAL_2"], "notes": "Zero external door transmission loss (default for interior rooms)"},

    # Section 14: MCS Heat Pump & Radiator Design Parameters (BS EN 12831-1 & MIS 3005-D)
    {"row": 175, "key": "hp_design_flow_temp", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "Heat Pump Design Flow Temperature", "default": 45.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "MCS MIS 3005-D low-flow design standard (35°C to 50°C)"},
    {"row": 176, "key": "hp_design_return_temp", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "Heat Pump Design Return Temperature", "default": 40.0, "unit": "°C", "fmt": FORMATS["TEMP_C"], "notes": "Return temperature at design heat load condition"},
    {"row": 177, "key": "hp_system_delta_t", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "System Design Temperature Drop (ΔT)", "default": 5.0, "unit": "K", "fmt": FORMATS["TEMP_C"], "notes": "Heat pump optimal emitter temperature drop (typically 5K)"},
    {"row": 178, "key": "thermal_bridge_factor", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "Thermal Bridging Allowance (f_tb)", "default": 0.10, "unit": "%", "fmt": FORMATS["PERCENT"], "notes": "MCS MIS 3005 default (+10% on fabric losses for existing buildings)"},
    {"row": 179, "key": "rad_exponent_n", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "Radiator Output Exponent (n)", "default": 1.30, "unit": "exponent", "fmt": FORMATS["DECIMAL_2"], "notes": "Standard BS EN 442 panel radiator exponent (typically 1.25 to 1.33)"},
    {"row": 180, "key": "reheat_allowance_pct", "section": "14. MCS Heat Pump & Radiator Design Parameters", "label": "Intermittent Reheat Allowance (f_hu)", "default": 0.00, "unit": "%", "fmt": FORMATS["PERCENT"], "notes": "0% for continuous 24/7 heat pump weather-compensated operation"}
]

SECTION_HEADERS = [
    {"row": 3, "title": "1. WEATHER & DESIGN TEMPERATURES"},
    {"row": 10, "title": "2. VENTILATION & INFILTRATION RATES"},
    {"row": 19, "title": "3. DOMESTIC HOT WATER (DHW) SYSTEM"},
    {"row": 31, "title": "4. SWIMMING POOL THERMAL PARAMETERS"},
    {"row": 41, "title": "5. HEATING SYSTEM EFFICIENCIES & SCOPS"},
    {"row": 51, "title": "6. ENERGY TARIFFS, SOLAR PV & BATTERY STORAGE"},
    {"row": 63, "title": "7. CAPITAL EQUIPMENT COST RATES & CARBON FACTORS"},
    {"row": 74, "title": "8. INFILTRATION QUESTIONNAIRE SCORING & AIR CHANGE RATE (ACH) PENALTIES"},
    {"row": 96, "title": "9. WINDOW & GLAZING THERMAL TRANSMITTANCE (U-VALUES)"},
    {"row": 111, "title": "10. CEILING & ROOF THERMAL TRANSMITTANCE (U-VALUES)"},
    {"row": 127, "title": "11. WALL CONSTRUCTION THERMAL TRANSMITTANCE (U-VALUES)"},
    {"row": 146, "title": "12. GROUND & FLOOR CONSTRUCTION THERMAL TRANSMITTANCE (U-VALUES)"},
    {"row": 164, "title": "13. EXTERNAL DOOR CONSTRUCTION THERMAL TRANSMITTANCE (U-VALUES)"},
    {"row": 174, "title": "14. MCS HEAT PUMP & RADIATOR DESIGN PARAMETERS"}
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
        total_rows = max(max(item["row"] for item in INPUT_DEFINITIONS), max(sh["row"] for sh in SECTION_HEADERS)) + 1
        target_ws_rows = max(total_rows + 5, 185)
        try:
            ws = self.ss.worksheet("1_Inputs")
            if ws.row_count < target_ws_rows:
                ws.resize(rows=target_ws_rows, cols=max(ws.col_count, 6))
        except gspread.WorksheetNotFound:
            ws = self.ss.add_worksheet(title="1_Inputs", rows=target_ws_rows, cols=6)

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
