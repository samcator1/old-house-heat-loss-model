# Walkthrough: Window Glazing Dropdown Questionnaire & Dynamic U-Value Model

The Google Sheets model **"Heat Loss & Usage Calcs_HB Comments"** and its companion on-site survey tools have been upgraded with an **In-Cell Window Glazing Questionnaire & Dynamic U-Value Engine**. Room window U-values are now calculated dynamically via live `VLOOKUP` formulas against an editable lookup schedule rather than being hardcoded, enabling instant scenario modelling (e.g. secondary glazing vs slimline double glazing) room-by-room.

🔗 [Open Live Google Sheet](https://docs.google.com/spreadsheets/d/102deNBRxDmlTeM7eRQhrdFKXY1V0KkJsD2-yz_R0TMw/edit)

---

## 🎯 What Was Built & Upgraded

### 1. In-Cell Window Specification Dropdowns (`2_Room_Heat_Loss`)
`2_Room_Heat_Loss` has been expanded to a **36-column master schedule** (Columns A through AJ). A dedicated **Window Specification** column is inserted at **Column O (col 15)** featuring native in-cell dropdown chips:

| Col | Field | Formula / Input Type | Technical Rationale |
|---|---|---|---|
| **N** | **Window Area ($m^2$)** | Direct Numeric Input | Measured glass aperture area per room |
| **O** | **Window Specification** | **In-Cell Dropdown** | 11 CIBSE Guide A Table 3.29 & Historic England glazing typologies |
| **P** | **U Window ($W/m^2K$)** | `=VLOOKUP(O{r}, '1_Inputs'!$B$98:$C$108, 2, FALSE)` | **100% Native Live Formula** referencing central inputs schedule |
| **Q** | **Window Loss ($W$)** | `=N{r} * P{r} * (E{r} - '1_Inputs'!$C$5)` | Dynamic heat loss at design outdoor temperature ($-4^\circ\text{C}$) |

---

### 2. Central Glazing Typology Lookup Schedule (`1_Inputs`)
Section 9 has been added to `1_Inputs` (Rows 96–108, Range `$B$98:$C$108`) with 11 standard typologies in soft blue editable cells (`#E0F0FD`):

| # | Glazing Specification Option | Default $U_{\text{win}}$ ($W/m^2K$) | Technical & Historic Engineering Rationale |
|---|---|---|---|
| 1 | **Single Glazed (Historic Timber Sash / Casement)** | **4.80** | CIBSE Guide A Table 3.29 standard for single glass in timber frame. |
| 2 | **Single Glazed (Metal / Crittall / Stone Mullion)** | **5.60** | Uninsulated steel, iron, or direct stone mullion mounting. |
| 3 | **Single Glazed + Heavy Thermal Curtains / Working Shutters** | **3.00** | Night-time or draught-isolated historic window reduction. |
| 4 | **Single Glazed + Standard Secondary Glazing** | **2.60** | Retrofitted internal timber/aluminium secondary pane (100mm+ cavity). |
| 5 | **Single Glazed + Low-E Secondary Glazing** | **1.80** | Internal secondary glazing pane with pyrolytic low-emissivity coating. |
| 6 | **Heritage Slimline Double Glazing (Timber Frame)** | **1.90** | Vacuum or narrow gas cavity (6–8mm) slimline units in heritage sashes. |
| 7 | **Early Standard Double Glazing (Pre-2002, Air Gap)** | **2.80** | 12mm air cavity without soft low-E coating. |
| 8 | **Modern Double Glazing (Argon, Low-E, Warm Edge)** | **1.40** | 16–20mm 90% Argon, soft-coat low-E, warm-edge composite spacer. |
| 9 | **Modern High-Performance Triple Glazing** | **0.80** | Passivhaus standard dual-cavity Krypton/Argon insulated units. |
| 10 | **Glazed Roof Lantern / Sloping Skylight** | **1.80** | Sloping glazed architectural lanterns (includes convection penalty). |
| 11 | **No External Glazing (Enclosed / Blind Room)** | **0.00** | Enclosed plantrooms, internal store cupboards, or blind spaces. |

> [!TIP]
> **Live Scenario Modelling**: Changing any U-value in `1_Inputs` `$B$98:$C$108` instantly cascades across all 23 rooms in `2_Room_Heat_Loss`, updating peak heat loss, radiator sizing, heat pump capacity, and annual operating costs in real time.

---

### 3. Master Room Schedule 36-Column Layout (`2_Room_Heat_Loss`)

All columns in `2_Room_Heat_Loss` have been realigned and formatted:

```
[A-D]  ROOM IDENTIFICATION: Code (A), Room Name (B), Floor Level (C), Zone / Wing (D)
[E-J]  DESIGN & GEOMETRY: Design Ti (E), Length (F), Width (G), Area (H), Height (I), Volume (J)
[K-W]  FABRIC LOSS CALCULATIONS (W):
       - Ext Wall L (K), U Wall (L), Wall Loss W (M)
       - Window Area m² (N), Window Specification (O), U Window (P), Window Loss W (Q)
       - Exp Floor m² (R), U Floor (S), Floor Loss W (T)
       - Ceiling Area m² (U), U Ceiling (V), Ceiling Loss W (W)
[X-AD] INFILTRATION & VENTILATION QUESTIONNAIRE:
       - Base Construction (X), Chimney / Flue (Y), Windows & Doors (Z), Floor Const (AA), Ceiling Boundary (AB)
       - Calculated ACH (AC), Vent Loss W (AD)
[AE-AH] TOTAL HEAT LOSS & EMITTERS:
       - Room Loss W (AE), Intensity W/m² (AF), Rad 45°C ΔT30 W (AG), Boiler Rad ΔT50 W (AH)
[AI-AJ] SPECIFICATION & NOTES:
       - Recommended Emitter (AI), On-Site Survey Notes (AJ)
```

- **Total Whole House Row (Row 28)**:
  - Total Glazing Area: **$115.8\,m^2$**
  - Total Window Heat Loss: **$8,742\,W$**
  - Total Whole-House Peak Heat Loss: **$45,114\,W$ ($45.1\,\text{kW}$)**
  - Average Building Heat Intensity: **$80.4\,W/m^2$**
  - Whole-House 45°C Radiators Required: **$45.1\,\text{kW}$**

---

### 4. Cross-Sheet Coordination & Downstream Integrity

All downstream sheets that draw from `2_Room_Heat_Loss` have been updated to target Column AE:
1. **`0_Executive_Dashboard`**:
   - Card 1 Peak Heat Loss: `='2_Room_Heat_Loss'!$AE$28/1000` ($45.1\,\text{kW}$)
   - Card 1 Average Intensity: `='2_Room_Heat_Loss'!$AF$28` ($80.4\,W/m^2$)
   - Section 2 Architectural Wing Breakdown: `=SUMIF(..., '2_Room_Heat_Loss'!$AE$5:$AE$27)/1000`
2. **`4_Heating_and_Renewables`**:
   - Space Heating Annual Demand (Row 5): Degree-day formula targets `'2_Room_Heat_Loss'!$AE$28`
   - Plant Sizing (Rows 13–15): GSHP ($45.1\,\text{kW}$), ASHP ($45.1\,\text{kW}$), Oil Boiler ($54.1\,\text{kW}$) target `'2_Room_Heat_Loss'!$AE$28`

---

### 5. On-Site Companion Survey Tool (`room_survey_tool.html`)

The interactive on-site survey application ([room_survey_tool.html](file:///c:/Users/samca/repos/old-house-heat-loss-model/room_survey_tool.html)) has been updated:
- **Glazing Typology Dropdown**: Replaced the plain numeric U-value input with an interactive `<select>` dropdown populated with all 11 CIBSE/Historic England window typologies.
- **Dynamic Glazing Loss Calculation**: Automatically retrieves the appropriate U-value upon selection, computes glazing heat loss in real time, and updates room total loss and intensity badges.
- **Enhanced CSV Export & Importer**:
  - `exportCSV()` now outputs the `Window Specification` text alongside dimensional and infiltration parameters.
  - `import_survey_csv.py` batch writes `Window Area` to Col N and `Window Specification` to Col O, letting Google Sheets calculate `U Window` in Col P via formula without overwriting formulas.

---

## 📊 Sample Room Glazing & Heat Loss Breakdown

| Room Code & Name | Glazing Area ($m^2$) | Selected Window Specification | Live Formula $U_{\text{win}}$ | Window Loss ($W$) | Total Room Loss ($W$) | Room Intensity |
|---|---|---|---|---|---|---|
| **GF-01** Entrance Hall | $6.5\,m^2$ | Single Glazed (Historic Timber Sash) | **4.80** | $686\,W$ | $4,293\,W$ | $152.6\,W/m^2$ |
| **GF-02** Drawing Room | $9.2\,m^2$ | Single Glazed (Historic Timber Sash) | **4.80** | $1,104\,W$ | $5,622\,W$ | $105.8\,W/m^2$ |
| **GF-04** Kitchen Area | $8.0\,m^2$ | Modern Double Glazing (Argon, Low-E) | **1.40** | $246\,W$ | $753\,W$ | $15.5\,W/m^2$ |
| **GF-05** Orangery | $35.0\,m^2$ | Modern Double Glazing (Argon, Low-E) | **1.40** | $1,176\,W$ | $4,396\,W$ | $73.3\,W/m^2$ |
| **FF-03** Master Bedroom | $5.8\,m^2$ | Single Glazed (Historic Timber Sash) | **4.80** | $612\,W$ | $2,580\,W$ | $75.0\,W/m^2$ |
| **FF-05** Bedroom 2 Suite | $4.2\,m^2$ | Modern Double Glazing (Argon, Low-E) | **1.40** | $129\,W$ | $514\,W$ | $18.7\,W/m^2$ |
| **FF-08** Bedroom 4 | $3.5\,m^2$ | Single Glazed (Historic Timber Sash) | **4.80** | $370\,W$ | $2,735\,W$ | $121.5\,W/m^2$ |

---

## ✅ Verification & Validation Results

| Test / Check | Result | Verification Notes |
|---|---|---|
| **Formula Errors** | **0 Errors** across all 7 tabs | Verified: No `#REF!`, `#NAME?`, `#VALUE!`, `#N/A`, or `#DIV/0!` in any tab. |
| **Lookup Accuracy** | **100% Match** | Sampled historic sash ($4.80$), modern double ($1.40$), verified against `$B$98:$C$108`. |
| **Data Validations** | **Verified** | In-cell dropdown chips rendered on Col O (Glazing) and Cols X–AB (Infiltration). |
| **Traceability** | **100% Formula** | Col P formula: `=VLOOKUP(O5, '1_Inputs'!$B$98:$C$108, 2, FALSE)`. |
| **Dashboard Synchronization** | **$45.1\,\text{kW}$** | Card 1 & Zone breakdown reflect exact sum of 23 assessed rooms. |
| **Git Tracking** | **Committed** | Commit `6b46b86` on `master`. |
