"""
config.py — Centralizirana konfiguracija aplikacije.

Sadrži:
  - PATHS: lokacije template Excela, projekata, outputa
  - THEME: paleta boja (light + dark mode), font, radius, shadows
  - CELL MAPS: točne adrese ćelija u Scenarios sheetu za Wind i Solar
  - DEFAULT SCENARIOS: 3 default scenarija (Base, Bank, Stress)

KRITIČNO: Cell mappings su verificirani protiv stvarnih .xlsm datoteka.
Wind koristi kolone G-L (scenariji 1-6) u Scenarios sheetu.
Solar koristi kolone H-M (scenariji 1-6) u Scenarios sheetu.
Aktivni scenarij se kontrolira preko Scenarios!E4 (vrijednost 1-6).
Naša aplikacija UVIJEK piše u "Scenario 1" stupac (G za Wind, H za Solar)
i postavlja E4=1 da forsira aktivnost.
"""
from __future__ import annotations
from pathlib import Path
from typing import Literal

# ════════════════════════════════════════════════════════════════════════════
# PATHS
# ════════════════════════════════════════════════════════════════════════════
ROOT_DIR = Path(__file__).parent.resolve()
TEMPLATES_DIR = ROOT_DIR / 'templates'
PROJECTS_DIR = ROOT_DIR / 'projects'

WIND_TEMPLATE = TEMPLATES_DIR / 'WIND_template.xlsm'
SOLAR_TEMPLATE = TEMPLATES_DIR / 'SOLAR_template.xlsm'

TEMPLATES = {
    'wind': WIND_TEMPLATE,
    'solar': SOLAR_TEMPLATE,
}

# ════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT — Active scenario mechanics
# ════════════════════════════════════════════════════════════════════════════
ACTIVE_SCENARIO_CELL = 'E4'        # Same in both models
ACTIVE_SCENARIO_VALUE = 1          # We always activate Scenario 1

# Column letters where "Scenario 1" lives in Scenarios sheet
SCENARIO_1_COL = {
    'wind': 'G',
    'solar': 'H',
}

# ════════════════════════════════════════════════════════════════════════════
# THEME — Akuo green palette (matches old Streamlit)
# ════════════════════════════════════════════════════════════════════════════
THEME = {
    'light': {
        'primary': '#1B5E3B',
        'primary_light': '#2E7D4A',
        'primary_dark': '#0D3D25',
        'accent': '#FF9800',
        'bg_main': '#F4F6F8',
        'bg_sidebar': '#1B5E3B',
        'bg_card': '#FFFFFF',
        'text_primary': '#1A1A1A',
        'text_secondary': '#6B7280',
        'border': '#E5E7EB',
        'success': '#10B981',
        'warning': '#F59E0B',
        'danger': '#EF4444',
    },
    'dark': {
        'primary': '#2E7D4A',
        'primary_light': '#4CAF50',
        'primary_dark': '#1B5E3B',
        'accent': '#FFB74D',
        'bg_main': '#121212',
        'bg_sidebar': '#1E1E1E',
        'bg_card': '#2D2D2D',
        'text_primary': '#E0E0E0',
        'text_secondary': '#A0A0A0',
        'border': '#404040',
        'success': '#4CAF50',
        'warning': '#FFC107',
        'danger': '#F44336',
    },
}

FONT_FAMILY = '"Inter", -apple-system, BlinkMacSystemFont, sans-serif'

# ════════════════════════════════════════════════════════════════════════════
# CELL MAPS — verified against actual .xlsm files
# ════════════════════════════════════════════════════════════════════════════
# Format: field_name -> (row, label, unit, default, min_val, max_val)
# row=None means the field is not in Scenarios sheet (handled via Inputs sheet)

WIND_CELL_MAP: dict[str, tuple] = {
    # ─── TECHNICAL (Scenarios rows 6-13) ─────────────────────────────────
    'capacity_mw':            (9,    'Capacity',                     'MW',     35.0,    1.0,    500.0),
    'num_turbines':           (10,   'Number of turbines',           'units',  5,       1,      200),
    'nominal_capacity':       (11,   'Nominal capacity per WTG',     'MW',     7.0,     1.0,    20.0),
    'yield_p50':              (12,   'Yield P50',                    'h',      4164.0,  500.0,  5000.0),
    'yield_p90':              (13,   'Yield P90',                    'h',      3752.0,  500.0,  5000.0),

    # ─── CAPEX (Scenarios rows 17-65) ────────────────────────────────────
    'capex_wind_turbines':    (17,   'Wind Turbines',                'k€',     35000,   0,      None),
    'capex_tsa_optionals':    (18,   'TSA Optionals',                'k€',     0,       0,      None),
    'capex_flow_parts':       (19,   'Flow Parts',                   'k€',     0,       0,      None),
    'capex_procurement_fees': (20,   'Procurement Fees',             'k€',     0,       0,      None),
    'capex_logistics':        (21,   'Logistics & Transport',        'k€',     0,       0,      None),
    'capex_electrical_bop':   (24,   'Electrical BOP',               'k€',     7000,    0,      None),
    'capex_civil_bop':        (27,   'Civil BOP',                    'k€',     5560,    0,      None),
    'capex_grid_connection':  (32,   'Grid connection cost',         'k€',     1000,    0,      None),
    'capex_telecom':          (37,   'Telecom',                      'k€',     50,      0,      None),
    'capex_scada':            (38,   'SCADA',                        'k€',     50,      0,      None),
    'capex_ems':              (39,   'EMS',                          'k€',     0,       0,      None),
    'capex_om_building':      (41,   'O&M Building',                 'k€',     500,     0,      None),
    'capex_weather_station':  (42,   'Weather Station',              'k€',     50,      0,      None),
    'capex_temp_roads':       (43,   'Temporary Access Roads',       'k€',     200,     0,      None),
    'capex_special_vehicles': (44,   'Special vehicles',             'k€',     0,       0,      None),
    'capex_es_mitigation':    (45,   'E&S/Mitigation',               'k€',     200,     0,      None),
    'capex_local_involve':    (46,   'Local Involvement',            'k€',     0,       0,      None),
    'capex_insurance_trc':    (48,   'TRC Insurance',                'k€',     150,     0,      None),
    'capex_insurance_rc':     (49,   'RC Insurance',                 'k€',     20,      0,      None),
    'capex_insurance_do':     (50,   'DO Insurance',                 'k€',     50,      0,      None),
    'capex_insurance_dsu':    (51,   'DSU Insurance',                'k€',     200,     0,      None),
    'capex_marine_cargo':     (52,   'Marine Cargo DSU',             'k€',     30,      0,      None),
    'capex_other_insurance':  (53,   'Other insurance',              'k€',     0,       0,      None),
    'capex_land_lease':       (55,   'Land lease',                   'k€',     400,     0,      None),
    'capex_easement':         (56,   'Easement',                     'k€',     0,       0,      None),
    'capex_expropriation':    (57,   'Expropriation',                'k€',     50,      0,      None),
    'capex_dev_costs':        (59,   'Development costs',            'k€',     1100,    0,      None),
    'capex_akuo_dev_pct':     (60,   'Akuo Dev Services (% capex)',  '%',      0.035,   0,      0.20),
    'capex_acquisition':      (61,   'Acquisition costs',            'k€',     0,       0,      None),
    'capex_akuo_constr_pct':  (62,   'Akuo Construction (% capex)',  '%',      0.025,   0,      0.20),
    'contingency_pct':        (64,   'Contingencies %',              '%',      0.06,    0,      0.30),

    # ─── OPEX (Scenarios rows 68-160) ────────────────────────────────────
    'opex_asset_mgmt':        (68,   'Asset Management',             'k€/y',   80,      0,      None),
    'opex_operation_mgmt':    (69,   'Operation Management',         'k€/y',   60,      0,      None),
    'opex_perf_monitoring':   (70,   'Performance monitoring',       'k€/y',   30,      0,      None),
    'opex_tech_inspection':   (71,   'Technical inspections',        'k€/y',   20,      0,      None),
    'opex_weather_forecast':  (72,   'Weather forecast service',     'k€/y',   15,      0,      None),
    'opex_scada':             (73,   'SCADA',                        'k€/y',   10,      0,      None),
    'opex_minor_maint':       (109,  'Minor Maintenance',            'k€/y',   50,      0,      None),
    'opex_substation_om':     (110,  'Substation & O&M Building',    'k€/y',   80,      0,      None),
    'opex_regulatory_insp':   (111,  'Regulatory inspections',       'k€/y',   20,      0,      None),
    'opex_hse_plan':          (112,  'HSE Prevention',               'k€/y',   25,      0,      None),
    'opex_met_station_maint': (113,  'Met Station Maintenance',      'k€/y',   10,      0,      None),
    'opex_blade_maint':       (114,  'Blade Maintenance',            'k€/y',   30,      0,      None),
    'opex_vehicle_maint':     (115,  'Vehicle Maintenance',          'k€/y',   15,      0,      None),
    'opex_other_infra':       (116,  'Other infrastructure',         'k€/y',   0,       0,      None),
    'opex_vegetation_mgmt':   (118,  'Vegetation management',        'k€/y',   30,      0,      None),
    'opex_repair_roads':      (119,  'Road repair',                  'k€/y',   10,      0,      None),
    'opex_pest_control':      (120,  'Pest control',                 'k€/y',   5,       0,      None),
    'opex_site_inspections':  (121,  'Site inspections',             'k€/y',   10,      0,      None),
    'opex_clean_blades':      (124,  'Clean blades',                 'k€/y',   20,      0,      None),
    'opex_water_supply':      (125,  'Water supply',                 'k€/y',   5,       0,      None),
    'opex_surveillance_sys':  (128,  'Surveillance systems',         'k€/y',   30,      0,      None),
    'opex_surveillance_pat':  (129,  'Surveillance patrols',         'k€/y',   40,      0,      None),
    'opex_op_all_risk':       (131,  'Operation All Risk insurance', 'k€/y',   200,     0,      None),
    'opex_tpl':               (132,  'TPL Insurance',                'k€/y',   30,      0,      None),
    'opex_substation_cov':    (133,  'Substation coverage',          'k€/y',   20,      0,      None),
    'opex_spare_parts_ins':   (134,  'Spare parts insurance',        'k€/y',   10,      0,      None),
    'opex_land_lease_op':     (136,  'Land Lease (operation)',       'k€/y',   100,     0,      None),
    'opex_property_tax':      (137,  'Property Tax',                 'k€/y',   80,      0,      None),
    'opex_power_consumption': (139,  'Power consumption',            'k€/y',   30,      0,      None),
    'opex_grid_usage_fee':    (140,  'Grid usage fee',               'k€/y',   50,      0,      None),
    'opex_balancing_costs':   (141,  'Balancing costs',              'k€/y',   100,     0,      None),
    'opex_wake_compensation': (162,  'Wake effect compensation',     'k€/y',   0,       0,      None),

    # ─── REVENUE (Scenarios rows 167-171) ────────────────────────────────
    'ppa_term':               (167,  'PPA Term',                     'years',  12,      0,      30),
    'ppa_index':              (168,  'PPA price inflation',          '%',      0.01,    0,      0.05),
    'ppa_tariff':             (169,  'PPA Base Tariff',              '€/MWh',  65.0,    0,      500.0),
    'pct_via_ppa':            (170,  '% production via PPA',         '%',      0.7,     0,      1.0),

    # ─── TAX (Scenarios rows 174-176) ────────────────────────────────────
    'tax_property':           (174,  'Property Tax',                 'k€/y',   0,       0,      None),
    'tax_property_land':      (175,  'Property Land Tax',            'k€/y',   0,       0,      None),
    'tax_concession_fee':     (176,  'Concession Fee',               'k€/y',   0,       0,      None),

    # ─── FINANCING (Scenarios rows 179-188) ──────────────────────────────
    'senior_maturity':        (179,  'Senior Debt Maturity',         'years',  14,      5,      20),
    'senior_margin_bps':      (180,  'Senior Debt Margin',           'bps',    220,     50,     500),
    'swap_rate':              (181,  'Swap rate',                    '%',      0.0314,  0,      0.10),
    'hedge_coverage':         (184,  'Hedge coverage',               '%',      1.0,     0,      1.0),
    'gearing_max':            (185,  'Gearing max',                  '%',      0.80,    0.50,   0.85),
    'dscr_ppa':               (186,  'DSCR PPA',                     'x',      1.15,    1.0,    2.0),
    'dscr_market':            (187,  'DSCR market',                  'x',      1.35,    1.0,    2.5),
    'shl_rate':               (188,  'SHL interest rate',            '%',      0.08,    0,      0.20),
}


SOLAR_CELL_MAP: dict[str, tuple] = {
    # ─── TECHNICAL ───────────────────────────────────────────────────────
    'capacity_mw':            (8,    'Capacity DC',                  'MW',     75.26,   1.0,    500.0),
    'capacity_ac':            (9,    'Capacity AC',                  'MW',     55.0,    1.0,    500.0),
    'yield_p50':              (10,   'Yield P50',                    'h',      1500.0,  500.0,  3000.0),
    'yield_p90':              (11,   'Yield P90',                    'h',      1400.0,  500.0,  3000.0),

    # ─── CAPEX (PV + BESS specific items) ────────────────────────────────
    'capex_pv_modules':       (29,   'PV Modules',                   'k€',     20000,   0,      None),
    'capex_procurement':      (31,   'Procurement (EPC)',            'k€',     5000,    0,      None),
    'capex_site_construction':(39,   'Site Construction',            'k€',     8000,    0,      None),
    'capex_bess_supply':      (50,   'BESS Supply + BOP',            'k€',     10000,   0,      None),
    'capex_project_mgmt':     (57,   'Project Management',           'k€',     500,     0,      None),
    'capex_takeover_comm':    (66,   'Take-over & Commissioning',    'k€',     200,     0,      None),
    'capex_spare_parts':      (67,   'Spare Parts',                  'k€',     300,     0,      None),
    'capex_om_default':       (69,   'O&M during default notif.',    'k€',     100,     0,      None),
    'capex_grid_subscription':(71,   'Grid Subscription Fees',       'k€',     150,     0,      None),
    'capex_grid_conn_fee':    (72,   'Grid Connection Fee',          'k€',     500,     0,      None),
    'capex_ohl_supply':       (73,   'OHL Supply & Cabling',         'k€',     1000,    0,      None),
    'capex_substation_akuo':  (77,   'Substation (Akuo)',            'k€',     1500,    0,      None),
    'capex_substation_tso':   (78,   'Substation (TSO/DSO)',         'k€',     2000,    0,      None),
    'contingency_pct':        (315,  'OpEx Contingency %',           '%',      0.04,    0,      0.30),

    # ─── REVENUE — PV ────────────────────────────────────────────────────
    'ppa_term':               (320,  'PPA Term',                     'years',  12,      0,      30),
    'ppa_index':              (321,  'PPA price inflation',          '%',      0.01,    0,      0.05),
    'ppa_tariff':             (322,  'PPA Base Tariff',              '€/MWh',  60.0,    0,      500.0),
    'pct_via_ppa':            (323,  '% via PPA',                    '%',      0.7,     0,      1.0),

    # ─── REVENUE — BESS (Solar-only) ─────────────────────────────────────
    'bess_term':              (328,  'BESS market term',             'years',  12,      0,      30),
    'storage_lifetime':       (329,  'Storage lifetime',             'years',  10,      5,      20),
    'bess_index':             (330,  'BESS price inflation',         '%',      0.02,    0,      0.05),
    'bess_tariff':            (331,  'BESS Base Tariff',             '€/MWh',  150.0,   0,      500.0),

    # ─── TAX ─────────────────────────────────────────────────────────────
    'tax_property':           (339,  'Property Tax',                 'k€/y',   0,       0,      None),
    'tax_property_land':      (340,  'Property Land Tax',            'k€/y',   0,       0,      None),
    'tax_concession_fee':     (341,  'Concession Fee',               'k€/y',   0,       0,      None),

    # ─── FINANCING ───────────────────────────────────────────────────────
    'senior_maturity':        (345,  'Senior Maturity',              'years',  14,      5,      20),
    'senior_margin_bps':      (346,  'Senior Margin',                'bps',    220,     50,     500),
    'swap_rate':              (347,  'Swap rate',                    '%',      0.0314,  0,      0.10),
    'hedge_coverage':         (348,  'Hedge coverage',               '%',      1.0,     0,      1.0),
    'gearing_max':            (349,  'Gearing max',                  '%',      0.80,    0.50,   0.85),
    'dscr_ppa':               (350,  'DSCR PPA',                     'x',      1.15,    1.0,    2.0),
    'dscr_pv_market':         (351,  'DSCR PV market',               'x',      1.35,    1.0,    2.5),
    'dscr_storage_market':    (352,  'DSCR Storage market',          'x',      1.40,    1.0,    2.5),

    # ─── SCHEDULE ────────────────────────────────────────────────────────
    'fc_date':                (353,  'Start of construction',        'date',   None,    None,   None),
    'construction_months':    (354,  'Construction duration',        'months', 18,      6,      48),

    # ─── SHL ─────────────────────────────────────────────────────────────
    'shl_rate':               (355,  'SHL rate',                     '%',      0.08,    0,      0.20),
}


def get_cell_map(project_type: Literal['wind', 'solar']) -> dict:
    """Return cell map for given project type."""
    return WIND_CELL_MAP if project_type == 'wind' else SOLAR_CELL_MAP


def get_excel_address(project_type: Literal['wind', 'solar'], field_name: str) -> str | None:
    """Get full Excel address (e.g. 'G169') for a given input field."""
    cell_map = get_cell_map(project_type)
    if field_name not in cell_map:
        return None
    row = cell_map[field_name][0]
    if row is None:
        return None
    col = SCENARIO_1_COL[project_type]
    return f'{col}{row}'


# ════════════════════════════════════════════════════════════════════════════
# DEFAULT SCENARIOS (3 templates)
# ════════════════════════════════════════════════════════════════════════════
DEFAULT_SCENARIOS = ['Base', 'Bank', 'Stress']

SCENARIO_OVERRIDES = {
    'Base': {
        # Use defaults from cell map
    },
    'Bank': {
        # Conservative: lower yield, lower tariff, higher DSCR target
        'yield_p50': lambda d: d * 0.95,
        'ppa_tariff': lambda d: d * 0.95,
        'dscr_ppa': lambda d: max(d, 1.20),
        'dscr_market': lambda d: max(d, 1.40),
    },
    'Stress': {
        # Stress test: yield P90, much higher debt cost
        'yield_p50': lambda d: d * 0.85,
        'ppa_tariff': lambda d: d * 0.85,
        'senior_margin_bps': lambda d: d + 100,
        'gearing_max': lambda d: min(d, 0.70),
    },
}


# ════════════════════════════════════════════════════════════════════════════
# CALCULATION SETTINGS
# ════════════════════════════════════════════════════════════════════════════
PERIODS_PER_YEAR = 2  # Semestrial (matches Excel default)
BANK_YEAR_DAYS = 360  # For interest accrual
DEFAULT_DISCOUNT_RATE = 0.08

# Numerical tolerances
SOLVER_TOLERANCE = 1e-6
MAX_ITERATIONS = 100

# ════════════════════════════════════════════════════════════════════════════
# UI SETTINGS
# ════════════════════════════════════════════════════════════════════════════
APP_NAME = "Akuo Energy — Project Finance Modeler"
APP_VERSION = "1.0.0"

KPI_DEFINITIONS = [
    {'key': 'project_irr',  'label': 'Project IRR',   'format': '.2%',   'icon': 'mdi:trending-up'},
    {'key': 'project_npv',  'label': 'Project NPV',   'format': ',.0f',  'unit': 'k€', 'icon': 'mdi:cash-multiple'},
    {'key': 'total_capex',  'label': 'Total CapEx',   'format': ',.0f',  'unit': 'k€', 'icon': 'mdi:hammer-wrench'},
    {'key': 'min_dscr',     'label': 'Min DSCR',      'format': '.2f',   'unit': 'x',  'icon': 'mdi:bank-check'},
    {'key': 'lcoe',         'label': 'LCOE',          'format': '.1f',   'unit': '€/MWh', 'icon': 'mdi:flash'},
]
