# 🏡 Old House Heat Loss & Heating Decarbonization Model (Google Sheets Engine)

A Python-driven, vibe-codable building physics and low-carbon heating model repository for Google Sheets.

This repository allows you to **vibe code** the entire building heat loss, domestic hot water (DHW), swimming pool, and low-carbon heating system architecture using Antigravity (`agy`). Whenever you or AGY update the code, pushing changes into Google Sheets builds **live, 100% transparent Google Sheets formulas** across all calculation tabs while **preserving all your custom edits in the `1_Inputs` tab**.

The original sheet is safely preserved as `_Legacy_Heating` for historical reference.

---

## 🌟 Key Features

1. **Vibe-Code with AGY**: You describe changes in plain English to AGY (e.g., *"Adjust Georgian end wall U-value to 0.60 after internal woodfibre insulation"*, *"Test an 800L cylinder vs a 500L cylinder recharge rate"*, or *"Update the off-peak electricity tariff to Octopus Cosy 13.5p"*), and AGY updates the Python model and syncs it.
2. **100% Native Google Sheets Formulas**: Calculations (`=SUM(...)`, `='1_Inputs'!C4 * ...`, `=ROUND(...)`) are written directly as spreadsheet formulas. You can click on any cell in Google Sheets to see exactly how numbers are derived for full auditability and traceability.
3. **Non-Destructive Inputs (`1_Inputs`)**:
   - The `1_Inputs` tab formats all adjustable parameters (temperatures, infiltration rates, DHW volumes, fuel prices, tariffs, solar capacities) in soft-blue editable cells.
   - When you re-run `python sync_model.py`, it **reads and preserves all your custom input edits in Google Sheets**, only writing new template rows/defaults if a cell was blank.
   - You can also force-reset defaults anytime with `--reset-inputs`.
4. **Physical & Engineering Rigor & MCS Compliance**:
   - **MCS MIS 3005-D & BS EN 12831 Compliance**: Built to meet UK heat pump installation standards and CIBSE Domestic Heating Design Guide (DHDG).
   - **External Doors & Net Wall Deduction**: Dedicated door area, door construction dropdown (uninsulated timber, solid hardwood, composite, insulated doorset), door U-value, and automatic deduction from gross external wall area.
   - **Thermal Bridging Allowance**: Default +10% thermal bridging factor applied to fabric transmission losses as recommended by MCS / BS EN 12831 for existing buildings (configurable in `1_Inputs!$C$178`).
   - **Brand New Low-Temperature Emitter Schedule**: Calculates design heat loss at 45°C flow / 40°C return, and applies the BS EN 442 non-linear derating exponent ($n=1.30$) to output the standard **$\Delta T 50$ catalogue radiator rating** needed to order radiators.
   - **Hydraulic Water Flow Rates & Pipe Sizing**: Computes design mass flow rate in litres/hour ($q = \Phi / (1.163 \times \Delta T)$ at $\Delta T = 5\,\text{K}$) and specifies minimum recommended copper pipe sizes (10mm, 15mm, or 22mm) per room for distribution balancing.
   - **Corrected Ground Floor $\Delta T$**: Sub-floor ground temperature is modeled at $\sim 10^\circ\text{C}$ ($\Delta T \approx 10\,\text{K}$) rather than $-4^\circ\text{C}$ outdoor air ($\Delta T = 24\,\text{K}$), eliminating artificial overestimation.
   - **Disaggregated Fabric Losses**: Dedicated formula columns for External Walls, Windows, External Doors, Ground Floor, and Roof.
   - **Infiltration Differentiation**: Peak room ACH (for emitter/radiator sizing) vs. diversified building average ACH (for central heat pump plant sizing).
   - **Hot Water Storage & Standing Losses**: 800L vessel sizing, peak reheat capacity, secondary pumped circulation loop losses, and weekly Legionella pasteurization boost.
   - **Heating Plant Comparison**: Side-by-side comparison of Ground Source Heat Pump (GSHP), Air Source Heat Pump (ASHP), and Oil Boiler.
   - **Renewables & Smart Tariffs**: Rooftop Solar PV, home battery storage load-shifting into cheap overnight slots (e.g. 15p), SEG export earnings, and net operational running costs.

---

## 📊 Tab Architecture

| Tab Name | Type | Description |
|---|---|---|
| **`0_Executive_Dashboard`** | Formula View | KPI cards (Peak kW, Annual kWh, Running Cost £/yr, Grid Import/Export, CO2 reduction), tech options comparison matrix, zone breakdown & renewable balance. |
| **`1_Inputs`** | User Control | Central assumptions and parameter control center (rows 1–182). Configures design temperatures, fuel prices, heat pump SCOPs, U-values (walls, windows, doors, floors, roofs), and MCS heat pump & radiator parameters (45°C flow, 40°C return, 5K delta T, +10% thermal bridging). Blue cells are editable and preserved non-destructively on sync. |
| **`2_Room_Heat_Loss`** | Dynamic Formula | Master 46-column room-by-room schedule compliant with MCS MIS 3005-D & BS EN 12831. Includes 5-badge compliance audit header, live dropdowns for walls/doors/windows/floors/ceilings, BS EN 12831 survey questionnaire, fabric, door & ventilation losses, +10% thermal bridging allowance, brand new low-temperature emitter sizing (BS EN 442 Delta T 50 catalogue rating), water flow rates (l/h), and recommended pipe sizes. |
| **`3_DHW_and_Pool`** | Dynamic Formula | Domestic hot water volume, storage vessel reheat power, secondary loop loss, Legionella cycle, and seasonal swimming pool thermal requirements. |
| **`4_Heating_and_Renewables`** | Dynamic Formula | Plant sizing, turnkey capital costs, seasonal SCOPs, annual fuel consumption, Solar PV + Battery load shifting, smart tariffs, and 10-year TCO. |
| **`_Archive_Wing_Heat_Loss`** | Protected Archive | Preserved macro 6-zone approximation with corrected ground ΔT (superseded by room schedule). |
| **`_Legacy_Heating`** | Protected Archive | The original sheet safely preserved with its original formulas and engineer comments. |

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
# Optional: create virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install requirements
pip install -r requirements.txt
```

### 2. Configure Spreadsheet ID
Set your Google Spreadsheet ID in `.env`:
```env
SPREADSHEET_ID=102deNBRxDmlTeM7eRQhrdFKXY1V0KkJsD2-yz_R0TMw
GOOGLE_SERVICE_ACCOUNT_FILE=service_account.json
```

### 3. Sync to Google Sheets
```bash
python sync_model.py
```

---

## 🛠️ CLI Options

| Flag | Description |
|---|---|
| `python sync_model.py` | Syncs the model and formulas, non-destructively preserving all user edits on `1_Inputs` and `2_Room_Heat_Loss`. |
| `python sync_model.py --reset-inputs` | Resets all input values on `1_Inputs` back to baseline defaults. |
| `python sync_model.py --export-csv` | Exports a backup CSV snapshot of all room data to `room_by_room_heat_loss_survey.csv`. |
| `python sync_model.py --sheet-id <ID_OR_URL>` | Syncs to a specific Google Spreadsheet ID or URL. |

---

## 📋 Google Sheets as the Master Database & Front End

Google Sheets is the single interactive front end and master database for the entire building model:
1. **Interactive Room Schedule (`2_Room_Heat_Loss`)**:
   - Edit room dimensions, ceiling heights, and external wall lengths directly in the sheet.
   - Select wall, window, floor, and ceiling constructions from dropdown menus linked to live U-values and infiltration adders on `1_Inputs`.
   - Select chimney / fireplace status from the single ventilation dropdown (air infiltration for fabric elements is automatically looked up from their respective specification tables on `1_Inputs`).
   - Automatically size low-temperature 45°C radiators and design flow rates (l/h) per room.
2. **Offline Backup & CSV Utilities**:
   - To create a local snapshot of your room data at any time, run:
     ```bash
     python sync_model.py --export-csv
     ```
   - To restore or import room survey data from a CSV file into `2_Room_Heat_Loss`:
     ```bash
     python import_survey_csv.py [path_to_csv]
     ```

---

## 💡 How to Vibe Code with AGY

When working with AGY in this repo, you can make natural language requests such as:

- *"Update the Georgian end walls to assume 100mm internal woodfibre insulation (U-value 0.28)."*
- *"Add an additional zone for the new workshop or annex."*
- *"Change the smart tariff structure to an Octopus Cosy heat pump tariff with two 13.5p off-peak windows."*
- *"Calculate required radiator sizes based on 45°C flow temperature for each room."*
- *"Simulate an extended 5-month pool season from May through September."*

AGY will edit the corresponding Python modules in `heat_loss_model/` and run `python sync_model.py` to push the formula updates directly into your Google Sheet!
