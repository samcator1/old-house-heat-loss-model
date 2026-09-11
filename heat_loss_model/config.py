"""
Configuration, styling theme, formats, and baseline assumptions for the Old House Heat Loss Model.
"""

from typing import Dict, Any

# Visual Design Theme (Modern Clean Engineering Palette)
THEME = {
    # Headers
    "PRIMARY_HEADER_BG": {"red": 0.09, "green": 0.16, "blue": 0.24},   # Deep Slate/Navy (#172A3A)
    "SECONDARY_HEADER_BG": {"red": 0.13, "green": 0.23, "blue": 0.35}, # Steel Navy (#223A59)
    "SECTION_HEADER_BG": {"red": 0.18, "green": 0.38, "blue": 0.58},   # Architectural Blue (#2E6194)
    "HEADER_TEXT": {"red": 1.0, "green": 1.0, "blue": 1.0},            # Pure White
    
    # Cells
    "INPUT_BG": {"red": 0.88, "green": 0.94, "blue": 0.99},            # Soft Sky Blue / User Editable (#E0F0FD)
    "CALC_BG": {"red": 0.94, "green": 0.97, "blue": 0.94},             # Soft Sage / Calculated (#F0F7F0)
    "ACCENT_BG": {"red": 1.0, "green": 0.97, "blue": 0.88},            # Soft Amber / Key Metric (#FFF7E0)
    "TOTAL_BG": {"red": 0.92, "green": 0.95, "blue": 0.98},             # Muted Steel Gray (#EBF2FA)
    "ZEBRA_BG": {"red": 0.98, "green": 0.99, "blue": 1.0},             # Ultra Light Slate (#F8FAFC)
    "CARD_BG": {"red": 0.95, "green": 0.97, "blue": 0.99},              # Dashboard Card (#F1F5F9)
    "CARD_HEADER_BG": {"red": 0.18, "green": 0.28, "blue": 0.38},       # Dark Card Header
    
    # Text
    "DARK_TEXT": {"red": 0.06, "green": 0.09, "blue": 0.16},           # Off-Black (#0F172A)
    "MUTED_TEXT": {"red": 0.39, "green": 0.45, "blue": 0.55},          # Muted Slate (#64748B)
    "ACCENT_TEXT": {"red": 0.02, "green": 0.44, "blue": 0.70},         # Blue Accent
    "SUCCESS_TEXT": {"red": 0.06, "green": 0.50, "blue": 0.28},        # Emerald Green
}

# Standard Number Format Strings for Google Sheets
FORMATS = {
    "CURRENCY_GBP": "£#,##0",
    "CURRENCY_GBP_DEC": "£#,##0.00",
    "PERCENT": "0.0%",
    "PERCENT_INT": "0%",
    "DECIMAL_1": "0.0",
    "DECIMAL_2": "0.00",
    "INTEGER": "#,##0",
    "WATTS": '#,##0 "W"',
    "KW": '0.0 "kW"',
    "KWH": '#,##0 "kWh"',
    "M2": '0.0 "m²"',
    "M3": '0.0 "m³"',
    "W_M2": '0.0 "W/m²"',
    "U_VAL": '0.00 "W/m²K"',
    "ACH": '0.00 "ACH"',
    "TEMP_C": '0.0 "°C"',
    "TONNES_CO2": '0.0 "t CO₂"',
}

# Default Building Zones for Final Build
DEFAULT_ZONES = [
    {
        "name": "Georgian end",
        "length": 14.75,
        "width": 6.25,
        "floors": 2.0,
        "h1": 2.80,
        "h2": 2.80,
        "pct_ext_wall": 0.65,
        "u_wall": 1.40,
        "pct_glazing": 0.20,
        "u_window": 4.80,
        "u_ground": 0.80,
        "u_roof": 0.18,
        "zone_type": "Old House"
    },
    {
        "name": "Thatched gable ended",
        "length": 6.00,
        "width": 16.50,
        "floors": 2.5,
        "h1": 2.80,
        "h2": 2.60,
        "pct_ext_wall": 0.37,
        "u_wall": 1.80,
        "pct_glazing": 0.20,
        "u_window": 4.80,
        "u_ground": 0.60,
        "u_roof": 0.30,
        "zone_type": "Old House"
    },
    {
        "name": "New build",
        "length": 10.00,
        "width": 13.00,
        "floors": 2.0,
        "h1": 2.80,
        "h2": 2.60,
        "pct_ext_wall": 0.28,
        "u_wall": 0.22,
        "pct_glazing": 0.25,
        "u_window": 1.40,
        "u_ground": 0.15,
        "u_roof": 0.15,
        "zone_type": "New Build"
    },
    {
        "name": "Orangery",
        "length": 6.00,
        "width": 10.00,
        "floors": 1.0,
        "h1": 2.80,
        "h2": 0.00,
        "pct_ext_wall": 0.50,
        "u_wall": 0.25,
        "pct_glazing": 0.70,
        "u_window": 1.40,
        "u_ground": 0.15,
        "u_roof": 1.50,
        "zone_type": "New Build"
    },
    {
        "name": "Lean to West",
        "length": 4.00,
        "width": 6.00,
        "floors": 1.0,
        "h1": 2.60,
        "h2": 0.00,
        "pct_ext_wall": 0.50,
        "u_wall": 2.20,
        "pct_glazing": 0.30,
        "u_window": 4.80,
        "u_ground": 1.10,
        "u_roof": 2.00,
        "zone_type": "Old House"
    },
    {
        "name": "Lean to North",
        "length": 3.00,
        "width": 6.00,
        "floors": 1.0,
        "h1": 2.60,
        "h2": 2.00,
        "pct_ext_wall": 0.17,
        "u_wall": 2.20,
        "pct_glazing": 0.30,
        "u_window": 4.80,
        "u_ground": 1.10,
        "u_roof": 2.00,
        "zone_type": "Old House"
    }
]

# Chimney & Flue Ventilation Allowances (BS EN 12831 & CIBSE Guide A)
CHIMNEY_SPECIFICATIONS = [
    {
        "label": "No Chimney / Permanently Sealed",
        "ach": 0.00,
        "notes": "No open flue or chimney has been permanently capped/sealed"
    },
    {
        "label": "Room-Sealed Stove with Ext Air",
        "ach": 0.05,
        "notes": "Closed woodburner with dedicated direct external combustion air duct"
    },
    {
        "label": "Flue Damper / Chimney Balloon Fitted",
        "ach": 0.15,
        "notes": "Open chimney fitted with inflatable throat balloon or mechanical damper"
    },
    {
        "label": "Open Fireplace (Unsealed Flue)",
        "ach": 0.60,
        "notes": "Unrestricted open chimney flue driving continuous warm air stack extraction"
    }
]

# Mechanical Extract & Ventilation Rates (BS EN 12831 & Building Regs Part F)
MECHANICAL_VENTILATION_SPECIFICATIONS = [
    {
        "label": "None (Natural Infiltration Only)",
        "ach": 0.00,
        "notes": "No mechanical extract ventilation; room relies on natural envelope infiltration and stack effect"
    },
    {
        "label": "Intermittent Extract Fan (Bathroom / WC - 15 l/s)",
        "ach": 0.20,
        "notes": "Building Regs Part F intermittent extract fan (15 l/s / 54 m³/h boost on occupancy/light switch)"
    },
    {
        "label": "Intermittent Extract Fan (Kitchen / Utility - 30-60 l/s)",
        "ach": 0.30,
        "notes": "Part F intermittent extract cooker hood or utility fan (30-60 l/s boost)"
    },
    {
        "label": "Continuous Mechanical Extract (dMEV / MEV)",
        "ach": 0.30,
        "notes": "Decentralized or central continuous mechanical extract running low-rate trickle extraction"
    },
    {
        "label": "Mechanical Ventilation with Heat Recovery (MVHR - 85% Efficiency)",
        "ach": 0.05,
        "notes": "Continuous balanced supply and extract with 85%+ counter-flow heat exchanger (minimal net heat loss)"
    }
]

# Planned Emitter Types for Heat Pump Sizing (BS EN 442 & MCS MIS 3005-D)
EMITTER_SPECIFICATIONS = [
    {
        "label": "Type 22 Panel Radiator (Double Convector)",
        "notes": "Standard double panel double convector radiator; ideal balance of output and wall projection (approx 100mm deep)"
    },
    {
        "label": "Type 33 High-Output Radiator (Triple Convector)",
        "notes": "Triple panel triple convector radiator for high heat loss rooms where wall length is constrained (approx 160mm deep)"
    },
    {
        "label": "Type 21 Slimline Radiator (Double Panel Single Convector)",
        "notes": "Slimline double panel radiator (approx 70mm deep) suitable for narrow corridors or behind doors"
    },
    {
        "label": "Type 11 Single Panel Radiator",
        "notes": "Single panel single convector for very low heat loss areas, small cloakrooms, or tight spaces"
    },
    {
        "label": "Underfloor Heating (UFH - Screed / Low-Profile)",
        "notes": "Low-temperature radiant floor heating operating at 35-40°C flow with high self-regulating thermal comfort"
    },
    {
        "label": "Fan Coil / Trench Convector (Active Low-Water-Temp)",
        "notes": "Fan-assisted hydronic convector providing high kW output at 35-45°C flow with forced convection"
    }
]

# CIBSE Domestic Heating Design Guide (DHDG) Standard Room Design Temperatures
CIBSE_DESIGN_TEMPS = [18.0, 20.0, 21.0, 22.0]

# Backwards compatibility alias
INFILTRATION_QUESTIONNAIRE = {
    "chimney_flue": CHIMNEY_SPECIFICATIONS
}

# Window & Glazing U-Value Specifications (CIBSE Guide A Table 3.29 & Historic England)
WINDOW_SPECIFICATIONS = [
    {
        "label": 'Single Glazed (Historic Timber Sash / Casement)',
        "u_value": 4.80,
        "ach": 0.40,
        "notes": 'Uninsulated traditional timber sash, 3-4mm float glass (CIBSE Guide A Table 3.29)'
    },
    {
        "label": 'Single Glazed (Draught-Stripped / Weather-Sealed Timber Sash)',
        "u_value": 4.40,
        "ach": 0.15,
        "notes": 'Historic England recommended timber sash retrofit with perimeter routed brush pile seals and compression beading'
    },
    {
        "label": 'Single Glazed (Metal / Crittall / Stone Mullion)',
        "u_value": 5.60,
        "ach": 0.40,
        "notes": 'Single glazed in uninsulated steel frame or direct to stone mullion'
    },
    {
        "label": 'Single Glazed + Heavy Thermal Curtains / Working Shutters',
        "u_value": 3.00,
        "ach": 0.25,
        "notes": 'Traditional solid interior wooden shutters closed or heavy thermal interlined curtains'
    },
    {
        "label": 'Single Glazed + Standard Secondary Glazing',
        "u_value": 2.60,
        "ach": 0.10,
        "notes": 'Internal secondary glazing pane with >20mm sealed cavity, clear float glass'
    },
    {
        "label": 'Single Glazed + Low-E Secondary Glazing',
        "u_value": 1.80,
        "ach": 0.10,
        "notes": 'Secondary glazing pane with pyrolytic Low-E coating, draught-sealed perimeter'
    },
    {
        "label": 'Heritage Slimline Double Glazing (Timber Frame)',
        "u_value": 1.90,
        "ach": 0.15,
        "notes": 'Ultra-slim vacuum or Krypton units (6-8mm) fitting traditional historic glazing rebates'
    },
    {
        "label": 'Early Standard Double Glazing (Pre-2002, Air Gap)',
        "u_value": 2.80,
        "ach": 0.10,
        "notes": '1st-generation aluminium spacer double glazing without Low-E coating'
    },
    {
        "label": 'Modern Double Glazing (Argon, Low-E, Warm Edge)',
        "u_value": 1.40,
        "ach": 0.00,
        "notes": 'Current UK Building Regulations Part L standard domestic double glazing'
    },
    {
        "label": 'Modern High-Performance Triple Glazing',
        "u_value": 0.80,
        "ach": 0.00,
        "notes": 'Passivhaus standard triple glazing with two Low-E coats and warm edge spacer'
    },
    {
        "label": 'Glazed Roof Lantern / Sloping Skylight',
        "u_value": 1.80,
        "ach": 0.15,
        "notes": 'Sloping or horizontal glazed orientation with increased internal convective transfer'
    },
    {
        "label": 'No External Glazing (Enclosed / Blind Room)',
        "u_value": 0.00,
        "ach": 0.00,
        "notes": 'Internal corridors, enclosed cloakrooms, or windowless plantrooms'
    }
]

# Ceiling & Roof U-Value Specifications (CIBSE Guide A Table 3.19/3.20 & Historic England)
CEILING_SPECIFICATIONS = [
    {
        "label": 'Intermediate Floor (Heated Space Above)',
        "u_value": 0.00,
        "ach": 0.00,
        "notes": 'Internal ceiling below heated room on upper floor; zero envelope transmission loss'
    },
    {
        "label": 'Uninsulated Loft (Bare Joists, No Quilt)',
        "u_value": 2.30,
        "ach": 0.20,
        "notes": 'CIBSE Guide A Table 3.19: lath/plaster ceiling below cold ventilated attic without insulation'
    },
    {
        "label": 'Historic Thatched Roof (Straw / Reed Thatch)',
        "u_value": 0.30,
        "ach": 0.30,
        "notes": 'Traditional 300-450mm straw/water reed thatch naturally provides continuous thermal insulation'
    },
    {
        "label": 'Older Loft Insulation: 50mm (Pre-1980 Quilt)',
        "u_value": 0.80,
        "ach": 0.15,
        "notes": 'Thin historic mineral wool between ceiling joists, often degraded or compressed'
    },
    {
        "label": 'Older Loft Insulation: 100mm (1980s-1990s Standard)',
        "u_value": 0.40,
        "ach": 0.10,
        "notes": 'Single layer mineral wool quilt filling joist depth only'
    },
    {
        "label": 'Moderate Loft Insulation: 150-200mm (Pre-2002 Standard)',
        "u_value": 0.22,
        "ach": 0.05,
        "notes": 'Joists filled plus partial top-up layer, typical existing retrofitted loft'
    },
    {
        "label": 'Modern Building Regs Loft: 270-300mm (Mineral Wool)',
        "u_value": 0.16,
        "ach": 0.05,
        "notes": 'Current Building Regulations Part L: 100mm between joists + 170mm cross-laid quilt'
    },
    {
        "label": 'High-Performance Deep Loft: 350-400mm (Quilt / Cellulose)',
        "u_value": 0.11,
        "ach": 0.03,
        "notes": 'Deep quilt or blown cellulose insulation achieving near-Passivhaus thermal resistance'
    },
    {
        "label": 'Uninsulated Sloping Roof / Lean-to (Lath & Plaster to Rafters)',
        "u_value": 2.00,
        "ach": 0.30,
        "notes": 'Sloping roof with lath & plaster directly under slate/tile rafters, no loft void (e.g. Lean-to scullery)'
    },
    {
        "label": 'Insulated Sloping Rafters (50-100mm PIR / Woodfibre)',
        "u_value": 0.35,
        "ach": 0.10,
        "notes": 'Sloping cathedral ceiling retrofitted with rigid PIR or breathable woodfibre board between rafters'
    },
    {
        "label": 'Modern Insulated Warm Flat Roof Extension',
        "u_value": 0.18,
        "ach": 0.05,
        "notes": 'Flat timber deck roof with 120-150mm rigid PIR insulation above deck (Building Regs compliant)'
    },
    {
        "label": 'Glazed Roof Lantern / Sloping Skylight (Orangery)',
        "u_value": 1.50,
        "ach": 0.15,
        "notes": 'Extensive sloping double glazed roof lantern over Orangery or garden room'
    }
]

# Wall Construction U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.3)
WALL_SPECIFICATIONS = [
    {
        "label": 'Solid Brick: 9" / 225mm (Uninsulated)',
        "u_value": 2.10,
        "ach": 0.80,
        "notes": 'CIBSE Guide A Table 3.3: 1-brick thick solid wall with interior plaster (e.g. uninsulated lean-to/scullery)'
    },
    {
        "label": 'Solid Brick: 13.5" / 330mm (Uninsulated)',
        "u_value": 1.70,
        "ach": 0.80,
        "notes": '1.5-brick thick traditional solid masonry wall with plaster'
    },
    {
        "label": 'Solid Brick: 18" / 450mm (Georgian Facade)',
        "u_value": 1.40,
        "ach": 0.80,
        "notes": '2-brick thick substantial Georgian external facade with internal lath and plaster finish'
    },
    {
        "label": 'Solid Stone: 450-500mm Sandstone / Limestone',
        "u_value": 1.80,
        "ach": 0.80,
        "notes": 'Historic England / CIBSE Guide A: Solid dressed or rubble stone wall (e.g. thatched gable wing)'
    },
    {
        "label": 'Solid Stone: 500-600mm Dense Granite / Whinstone',
        "u_value": 2.20,
        "ach": 0.80,
        "notes": 'Dense crystalline igneous or metamorphic stone without cavity'
    },
    {
        "label": 'Historic Timber Frame (Wattle & Daub / Nogging)',
        "u_value": 1.80,
        "ach": 0.80,
        "notes": 'Traditional exposed timber box-frame with wattle/daub or brick nogging infill'
    },
    {
        "label": 'Cob / Earth Construction (500mm+)',
        "u_value": 0.90,
        "ach": 0.80,
        "notes": 'Traditional thick monolithic chalk/clay/straw cob wall with breathable lime render'
    },
    {
        "label": 'Uninsulated Cavity Wall (Pre-1976)',
        "u_value": 1.50,
        "ach": 0.50,
        "notes": 'Early cavity construction (brick/block or brick/brick) with uninsulated 50mm air gap'
    },
    {
        "label": 'Retrofilled Cavity Wall (Blown Mineral / Bead)',
        "u_value": 0.50,
        "ach": 0.40,
        "notes": 'Existing masonry cavity injected with bonded EPS beads or blown mineral fibre'
    },
    {
        "label": 'Partial-Fill Cavity Wall (1980s-1990s)',
        "u_value": 0.45,
        "ach": 0.35,
        "notes": '25-50mm partial-fill insulation board retained against inner leaf'
    },
    {
        "label": 'Modern Building Regs Cavity (100mm PIR / Full Fill)',
        "u_value": 0.18,
        "ach": 0.30,
        "notes": 'Current Building Regulations Part L new build standard (e.g. New Build wing)'
    },
    {
        "label": 'High-Performance New Build (150mm+ PIR / Passivhaus)',
        "u_value": 0.12,
        "ach": 0.15,
        "notes": 'Ultra-low-energy extension specification with thick high-performance insulation'
    },
    {
        "label": 'Solid Wall + 50mm Breathable Woodfibre IWI',
        "u_value": 0.55,
        "ach": 0.50,
        "notes": 'Historic England recommended internal breathable woodfibre / cork lime retrofit'
    },
    {
        "label": 'Solid Wall + 100mm Woodfibre / PIR IWI',
        "u_value": 0.28,
        "ach": 0.40,
        "notes": 'Deep internal wall insulation with intelligent vapour control membrane'
    },
    {
        "label": 'Solid Wall + 100mm External Wall Insulation (EWI)',
        "u_value": 0.25,
        "ach": 0.35,
        "notes": 'External insulation with breathable render finish on non-heritage elevations'
    },
    {
        "label": 'Party Wall / Heated Boundary (No Heat Loss)',
        "u_value": 0.00,
        "ach": 0.00,
        "notes": 'Internal partition or party wall adjoining heated space; zero envelope transmission loss'
    },
    {
        "label": 'Internal Partition to Unheated Garage / Store / Cellar',
        "u_value": 1.05,
        "ach": 0.00,
        "notes": 'BS EN 12831 unheated boundary: internal masonry/stud partition adjoining cold unconditioned space (bu = 0.50 applied to 2.1 W/m²K)'
    }
]

# Ground & Floor Construction U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.23 / 3.24)
FLOOR_SPECIFICATIONS = [
    {
        "label": 'Intermediate Floor (Heated Space Below)',
        "u_value": 0.00,
        "ach": 0.00,
        "notes": 'Internal intermediate floor with actively heated space beneath; zero envelope transmission loss'
    },
    {
        "label": 'Modern Building Regs Insulated Slab (100mm PIR / Full Fill)',
        "u_value": 0.15,
        "ach": 0.00,
        "notes": 'Current Building Regulations Part L ground-bearing concrete slab with 100-120mm continuous PIR'
    },
    {
        "label": 'High-Performance Insulated Slab (150mm+ PIR / Passivhaus)',
        "u_value": 0.10,
        "ach": 0.00,
        "notes": 'Fully wrapped raft or slab with 150-200mm high-density insulation achieving near-zero ground heat loss'
    },
    {
        "label": 'Suspended Timber: Insulated (100-150mm PIR between Joists)',
        "u_value": 0.18,
        "ach": 0.05,
        "notes": 'Retrofitted timber floor with 100-150mm rigid PIR or dense woodfibre supported on breathable membrane'
    },
    {
        "label": 'Suspended Timber: Insulated (50mm Quilt / Board)',
        "u_value": 0.35,
        "ach": 0.10,
        "notes": 'Partial retrofit with 50mm insulation batts between joists over crawlspace void'
    },
    {
        "label": 'Solid Concrete: Moderate Insulation (50mm PIR / 1990s)',
        "u_value": 0.35,
        "ach": 0.00,
        "notes": '1990s to early 2000s standard slab with 50mm insulation beneath concrete or screed'
    },
    {
        "label": 'Solid Concrete: Perimeter Insulation (1980s Standard)',
        "u_value": 0.45,
        "ach": 0.00,
        "notes": '1980s standard ground slab with vertical perimeter edge insulation only'
    },
    {
        "label": 'Suspended Timber: Carpet & Underlay over Uninsulated Void',
        "u_value": 0.60,
        "ach": 0.10,
        "notes": 'Historic suspended timber joists with tongue-and-groove boards, heavy underlay and fitted carpet'
    },
    {
        "label": 'Suspended Timber: Bare Boards over Cold Uninsulated Void',
        "u_value": 0.80,
        "ach": 0.35,
        "notes": 'CIBSE Guide A Table 3.24: Uninsulated suspended timber floorboards over ventilated sub-floor void'
    },
    {
        "label": 'Solid Concrete Ground Slab (Uninsulated Pre-1976)',
        "u_value": 0.80,
        "ach": 0.00,
        "notes": 'Traditional uninsulated solid concrete ground slab cast directly on hardcore/blinding'
    },
    {
        "label": 'Solid Ground Floor: Uninsulated Quarry Tiles / Stone / Earth',
        "u_value": 1.10,
        "ach": 0.00,
        "notes": 'Historic uninsulated flagstones, brick pavers, or quarry tiles bedded directly on ground / lime bed'
    },
    {
        "label": 'Exposed Floor over Cold Undercroft / Garage (Uninsulated)',
        "u_value": 1.20,
        "ach": 0.15,
        "notes": 'First floor cantilever or room over open driveway, unheated garage, or unheated archway without insulation'
    },
    {
        "label": 'Exposed Floor over Cold Undercroft / Garage (100mm PIR Insulated)',
        "u_value": 0.20,
        "ach": 0.05,
        "notes": 'Cantilevered exposed floor retrofitted with 100mm rigid insulation board and soffit board'
    },
    {
        "label": 'Unheated Adjoining Boundary (Zero Transmission Loss)',
        "u_value": 0.00,
        "ach": 0.00,
        "notes": 'Floor separating heated space from fully conditioned adjoining residential unit'
    },
    {
        "label": 'Suspended Timber / Floor over Unheated Cellar Void',
        "u_value": 0.50,
        "ach": 0.15,
        "notes": 'BS EN 12831: Floor above unheated cellar or semi-underground basement (bu = 0.60 applied to 0.80 W/m²K)'
    },
    {
        "label": 'Historic Breathable Insulated Limecrete Floor (Foamed Glass / Cork Lime)',
        "u_value": 0.22,
        "ach": 0.00,
        "notes": 'Historic England heritage specification: compacted foamed glass aggregate sub-base with breathable lime screed'
    }
]

# External Door U-Value Specifications (BS EN 12831-1 & CIBSE Guide A Table 3.28)
DOOR_SPECIFICATIONS = [
    {
        "label": "Solid Timber Door (Uninsulated / Historic Plank)",
        "u_value": 3.00,
        "notes": "Traditional solid oak/pine external door (25-45mm) without weather seals"
    },
    {
        "label": "Solid Timber Door (Draught-Stripped / Heavy Ledge)",
        "u_value": 2.40,
        "notes": "Heavy historic plank door with compression draught-proofing seals"
    },
    {
        "label": "Part-Glazed External Door (Single Glazed)",
        "u_value": 3.60,
        "notes": "Traditional timber entrance door with single-glazed upper lights"
    },
    {
        "label": "Part-Glazed External Door (Double Glazed)",
        "u_value": 2.00,
        "notes": "Timber or aluminium framed door with 16mm argon double-glazed unit"
    },
    {
        "label": "Modern High-Performance / Composite Insulated Door",
        "u_value": 1.20,
        "notes": "Part L compliant foam-filled composite or timber core door with thermal break"
    },
    {
        "label": "Modern French / Bi-fold Glazed Doors (Low-E Double)",
        "u_value": 1.40,
        "notes": "Modern high-performance multi-pane patio doors with thermal breaks"
    },
    {
        "label": "No External Door (Internal Boundary Only)",
        "u_value": 0.00,
        "notes": "Zero external door transmission loss (default for interior rooms)"
    }
]
