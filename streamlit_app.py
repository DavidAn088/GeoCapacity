
from pathlib import Path
from datetime import datetime
import ast
import io
import json
import math
import zipfile
import re
import urllib.request
import urllib.error
import base64

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle, RegularPolygon, Polygon, Patch

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage

try:
    import streamlit as st
except Exception:
    raise RuntimeError("Please install Streamlit: pip install streamlit")

APP_DIR = Path(__file__).parent
TEMPLATE_DIR = APP_DIR / "templates"
SAMPLE_DIR = APP_DIR / "samples"
HISTORY_DIR = APP_DIR / "history"
OUTPUT_DIR = APP_DIR / "output"
DOCUMENT_DIR = APP_DIR / "documents"
MANUAL_UPLOAD_DIR = APP_DIR / "manual_uploads"
EQUATION_PRESET_DIR = APP_DIR / "equation_presets"
SHALLOW_EQUATION_PRESET_DIR = APP_DIR / "shallow_equation_presets"
for p in [TEMPLATE_DIR, SAMPLE_DIR, HISTORY_DIR, OUTPUT_DIR, DOCUMENT_DIR, MANUAL_UPLOAD_DIR, EQUATION_PRESET_DIR, SHALLOW_EQUATION_PRESET_DIR]:
    p.mkdir(exist_ok=True)

# This can be changed from the sidebar. 9.81 reproduces the uploaded Excel.
TON_TO_KN = 9.81

st.set_page_config(page_title="GeoCapacity", page_icon="🌍", layout="wide", initial_sidebar_state="expanded")

APP_CSS = """
<style>
.main .block-container {padding-top:0.75rem; padding-bottom:1.6rem; max-width:1600px;}
.app-title{padding:1.0rem 1.2rem;border-radius:24px;background:linear-gradient(135deg,#0f172a 0%,#1d4ed8 60%,#22c55e 100%);color:white;box-shadow:0 15px 36px rgba(15,23,42,.24);margin-bottom:0.8rem;}
.app-title h1{margin:0;font-size:2.05rem;letter-spacing:.2px;}.app-title p{margin:.35rem 0 0 0;opacity:.94;font-size:1.02rem;line-height:1.55;}

.small-note{font-size:.86rem;color:#64748b}.software-note{padding:.75rem 1rem;border-radius:18px;background:#eff6ff;border:1px solid #bfdbfe;color:#1e3a8a;}
.metric-card{border-radius:20px;padding:0.9rem 1rem;background:white;border:1px solid #e5e7eb;box-shadow:0 8px 22px rgba(15,23,42,.07);}
.top-toolbar{display:flex;gap:.5rem;flex-wrap:wrap;margin:-.2rem 0 .8rem 0}.toolbar-chip{padding:.56rem .88rem;border-radius:16px;background:linear-gradient(180deg,#fff 0%,#f8fafc 100%);border:1px solid #dbe4ef;box-shadow:0 8px 20px rgba(15,23,42,.075);font-weight:800;color:#0f172a}.toolbar-chip span{opacity:.68;font-weight:650}
.project-tree{font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;font-size:.86rem;line-height:1.55;padding:.75rem;border-radius:16px;background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.10)}.project-tree .ok{color:#86efac;font-weight:700}.project-tree .wait{color:#fcd34d;font-weight:700}
div.stButton > button:first-child{width:100%;border-radius:18px;min-height:3.0rem;font-weight:800;border:0;color:white;background:linear-gradient(135deg,#2563eb 0%,#22c55e 100%);box-shadow:0 10px 22px rgba(37,99,235,.28);transition:all .2s ease-in-out}div.stButton > button:first-child:hover{transform:translateY(-1px);box-shadow:0 12px 28px rgba(37,99,235,.35);color:white}
div.stDownloadButton > button:first-child{width:100%;border-radius:16px;min-height:2.6rem;font-weight:700;border:1px solid #dbeafe;background:linear-gradient(180deg,#fff 0%,#eff6ff 100%);color:#1d4ed8;}
.stFileUploader{border-radius:18px;background:#f8fafc;padding:.2rem .6rem .8rem .6rem;border:1px dashed #cbd5e1;}
.stRadio [role="radiogroup"] label{width:100%;min-height:44px;display:flex;align-items:center;border:1px solid rgba(148,163,184,.35);padding:10px 12px;border-radius:14px;margin:0;transition:all .15s ease-in-out}.stRadio [role="radiogroup"] label:hover{background:rgba(148,163,184,.10);border-color:rgba(148,163,184,.55)}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)

def inject_left_board_theme(theme_name):
    if str(theme_name).lower().startswith("bright"):
        st.markdown("""
        <style>
        [data-testid="stSidebar"]{background:linear-gradient(180deg,#f8fafc 0%,#e2e8f0 100%) !important;} [data-testid="stSidebar"] *{color:#0f172a !important;}
        .project-tree{background:rgba(255,255,255,.72) !important;border:1px solid rgba(15,23,42,.12) !important}.project-tree .ok{color:#15803d !important}.project-tree .wait{color:#a16207 !important}
        .stRadio [role="radiogroup"] label{background:rgba(255,255,255,.82) !important;border:1px solid rgba(15,23,42,.10) !important;}
        </style>""", unsafe_allow_html=True)


# -----------------------------
# Supported engineering options
# -----------------------------
VALID_PILE_TYPE_MAP = {
    "B": "Bored pile",
    "BORED": "Bored pile",
    "BORED PILE": "Bored pile",
    "CAST-IN-PLACE": "Bored pile",
    "CAST IN PLACE": "Bored pile",
    "CAST_IN_PLACE": "Bored pile",
    "D": "Driven pile",
    "DRIVEN": "Driven pile",
    "DRIVEN PILE": "Driven pile",
    "DRIVE": "Driven pile",
}
VALID_SECTION_MAP = {
    "C": "Solid circular",
    "CIRCLE": "Solid circular",
    "CIRCULAR": "Solid circular",
    "S": "Solid square",
    "SQUARE": "Solid square",
    "H": "Solid hexagonal",
    "HEX": "Solid hexagonal",
    "HEXAGON": "Solid hexagonal",
    "SC": "Hollow spun circular",
    "SPUN_CIRCLE": "Hollow spun circular",
    "HOLLOW_CIRCLE": "Hollow spun circular",
    "SS": "Hollow square",
    "SPUN_SQUARE": "Hollow square",
    "HOLLOW_SQUARE": "Hollow square",
    "I": "I section",
    "I_SECTION": "I section",
    "CUSTOM": "User-defined section",
}

def pile_type_reference_df():
    return pd.DataFrame([
        {"pile_type": "B", "description": "Bored pile / cast-in-place pile", "used_in_equations_as": "pile_type_code = 1"},
        {"pile_type": "D", "description": "Driven pile", "used_in_equations_as": "pile_type_code = 2"},
    ])

def pile_section_reference_df():
    return pd.DataFrame([
        {"pile_section": "C", "description": "Solid circular", "pile_size means": "diameter", "extra input": "none"},
        {"pile_section": "S", "description": "Solid square", "pile_size means": "width", "extra input": "none"},
        {"pile_section": "H or HEX", "description": "Solid hexagonal", "pile_size means": "across flats", "extra input": "none"},
        {"pile_section": "SC", "description": "Hollow spun circular", "pile_size means": "outside diameter", "extra input": "inner_size = inside diameter"},
        {"pile_section": "SS", "description": "Hollow square", "pile_size means": "outside width", "extra input": "inner_size = inside width"},
        {"pile_section": "I", "description": "I section", "pile_size means": "overall depth", "extra input": "provide pile_perimeter and pile_cross_section_area for exact calculation"},
        {"pile_section": "CUSTOM", "description": "User-defined geometry", "pile_size means": "optional", "extra input": "pile_perimeter and pile_cross_section_area"},
    ])

def normalize_pile_type_text(value):
    key = clean_text(value, "").upper().replace("-", " ")
    key = " ".join(key.split())
    if key in VALID_PILE_TYPE_MAP:
        return "D" if VALID_PILE_TYPE_MAP[key].startswith("Driven") else "B"
    return ""

def normalize_section_text(value):
    key = clean_text(value, "").upper().replace("-", "_")
    key = "_".join(key.split())
    if key in VALID_SECTION_MAP:
        if key in ["CIRCLE", "CIRCULAR"]: return "C"
        if key == "SQUARE": return "S"
        if key in ["HEX", "HEXAGON"]: return "H"
        if key in ["SPUN_CIRCLE", "HOLLOW_CIRCLE"]: return "SC"
        if key in ["SPUN_SQUARE", "HOLLOW_SQUARE"]: return "SS"
        if key in ["I", "I_SECTION"]: return "I"
        return key
    return ""

# -----------------------------
# Basic helpers
# -----------------------------
def get_default_file(name):
    for folder in [APP_DIR, TEMPLATE_DIR, SAMPLE_DIR]:
        p = folder / name
        if p.exists():
            return p
    return None

def read_csv_from_upload(uploaded_file, fallback_path=None):
    if uploaded_file is not None:
        return pd.read_csv(uploaded_file)
    if fallback_path is not None and Path(fallback_path).exists():
        return pd.read_csv(fallback_path)
    return pd.DataFrame()

def clean_text(v, default=""):
    if pd.isna(v):
        return default
    return str(v).strip()

def to_float(v, default=0.0):
    if pd.isna(v) or v == "":
        return default
    try:
        return float(v)
    except Exception:
        return default

def is_missing(v):
    try:
        return pd.isna(v)
    except Exception:
        return False

def avg(*args):
    vals = []
    for a in args:
        try:
            x = float(a)
            if not math.isnan(x) and not math.isinf(x):
                vals.append(x)
        except Exception:
            pass
    return sum(vals) / len(vals) if vals else 0.0

def unit_label(unit):
    unit = clean_text(unit, "kN").lower()
    return "ton" if unit in ["t", "ton", "tons", "tonne", "tonnes"] else "kN"

def force_to_kn(value, unit):
    return to_float(value, 0.0) * TON_TO_KN if unit_label(unit) == "ton" else to_float(value, 0.0)

def force_from_ton(value_ton, unit):
    return to_float(value_ton, 0.0) if unit_label(unit) == "ton" else to_float(value_ton, 0.0) * TON_TO_KN

def stress_from_t_m2(value_t_m2, unit):
    return to_float(value_t_m2, 0.0) if unit_label(unit) == "ton" else to_float(value_t_m2, 0.0) * TON_TO_KN

def cap_unit(unit):
    return "ton" if unit_label(unit) == "ton" else "kN"

def stress_unit(unit):
    return "t/m²" if unit_label(unit) == "ton" else "kPa"

def gamma_unit(unit):
    return "t/m³" if unit_label(unit) == "ton" else "kN/m³"

def display_no_missing(df):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].replace([np.inf, -np.inf], np.nan).fillna(0.0)
        else:
            out[c] = out[c].fillna("")
    return out

def round_for_display(df, decimals=None):
    """Round only for screen/export display. Use None for full precision."""
    out = display_no_missing(df)
    if out is None or out.empty or decimals is None:
        return out
    try:
        d = int(decimals)
    except Exception:
        return out
    num_cols = out.select_dtypes(include=[np.number]).columns
    out[num_cols] = out[num_cols].round(d)
    return out

def read_reference_table(uploaded_file=None, sheet_name="Test2"):
    """Read optional reference CSV/XLSX uploaded by user for comparison."""
    if uploaded_file is None:
        return None
    try:
        name = getattr(uploaded_file, "name", "").lower()
        if name.endswith(".xlsx") or name.endswith(".xls"):
            return pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return pd.read_csv(uploaded_file)
    except Exception as e:
        st.warning(f"Could not read reference file: {e}")
        return None

def normalize_reference_columns(df):
    """Clean a reference table and keep numeric columns where possible."""
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out = out.dropna(how="all").reset_index(drop=True)
    # Drop completely blank columns and unnamed empty columns.
    out = out.loc[:, ~out.columns.astype(str).str.match(r"^Unnamed") | out.notna().any(axis=0)]
    out.columns = [str(c).strip() for c in out.columns]
    return out

# -----------------------------
# Shape functions
# -----------------------------
def shape_geometry(section, size_m, inner_size_m=0.0, flange_width_m=0.3, web_thickness_m=0.12, flange_thickness_m=0.12):
    section = clean_text(section, "C").upper()
    D = to_float(size_m, 0.6)
    di = to_float(inner_size_m, 0.0)
    if section in ["S", "SQUARE"]:
        return 4 * D, D**2
    if section in ["C", "CIRCLE", "CIRCULAR"]:
        return math.pi * D, math.pi * D**2 / 4
    if section in ["SC", "SPUN_CIRCLE", "HOLLOW_CIRCLE"]:
        return math.pi * D, math.pi * max(D**2 - di**2, 0.0) / 4
    if section in ["SS", "SPUN_SQUARE", "HOLLOW_SQUARE"]:
        return 4 * D, max(D**2 - di**2, 0.0)
    if section in ["H", "HEX", "HEXAGON"]:
        side = D / math.sqrt(3)
        return 6 * side, (3 * math.sqrt(3) / 2) * side**2
    if section in ["I", "I_SECTION"]:
        # If exact perimeter/area are supplied in the CSV, normalize_pile_case keeps them.
        # This fallback is only for preview/rough geometry.
        bf = max(flange_width_m, 0.25*D)
        tw = min(max(web_thickness_m, 0.12*D), bf)
        tf = min(max(flange_thickness_m, 0.12*D), D/3)
        area = max(2*bf*tf + tw*max(D-2*tf, 0.0), 0.0)
        per = 4*bf + 2*(D-2*tf) + 4*tf
        return per, area
    if section in ["CUSTOM"]:
        # Geometry will be taken from pile_perimeter and pile_cross_section_area.
        # This fallback prevents crashes if the custom values are still blank.
        return math.pi * D, math.pi * D**2 / 4
    return math.pi * D, math.pi * D**2 / 4

# -----------------------------
# Safe equation evaluator
# -----------------------------
ALLOWED_FUNCS = {
    "sin": math.sin, "cos": math.cos, "tan": math.tan, "exp": math.exp, "log": math.log,
    "log10": math.log10, "sqrt": math.sqrt, "atan": math.atan, "radians": math.radians, "degrees": math.degrees,
    "min": min, "max": max, "abs": abs, "avg": avg,
}
ALLOWED_NAMES = {"pi": math.pi, "e": math.e, "nan": float("nan"), **ALLOWED_FUNCS}

class SafeEval(ast.NodeVisitor):
    allowed_nodes = (
        ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name, ast.Load,
        ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.Mod, ast.USub, ast.UAdd,
        ast.Call, ast.IfExp, ast.Compare, ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
        ast.BoolOp, ast.And, ast.Or,
    )
    def __init__(self, variables):
        self.variables = variables
    def visit(self, node):
        if not isinstance(node, self.allowed_nodes):
            raise ValueError(f"Unsupported expression: {type(node).__name__}")
        return super().visit(node)
    def visit_Call(self, node):
        if not isinstance(node.func, ast.Name) or node.func.id not in ALLOWED_FUNCS:
            raise ValueError("Only safe math functions are allowed.")
        for arg in node.args:
            self.visit(arg)
    def visit_Name(self, node):
        if node.id not in ALLOWED_NAMES and node.id not in self.variables:
            raise ValueError(f"Unknown variable: {node.id}")
    def eval(self, expression):
        tree = ast.parse(str(expression), mode="eval")
        self.visit(tree)
        return eval(compile(tree, "<equation>", "eval"), {"__builtins__": {}}, {**ALLOWED_NAMES, **self.variables})

def safe_eval(expression, variables, default=0.0):
    try:
        value = SafeEval(variables).eval(expression)
        if isinstance(value, (bool, np.bool_)):
            return float(value)
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            return default
        return value
    except Exception:
        return default

# -----------------------------
# Equations
# -----------------------------
def default_equations_df():
    p = get_default_file("equation_library_bangkok_excel.csv")
    if p:
        return pd.read_csv(p)
    return pd.DataFrame([
        ["su2_t_m2", "nan if SPT_N_missing > 0.5 else 0.52 * SPT_N", r"s_{u2}=0.52N\quad(\mathrm{blank\ if\ SPT\ is\ blank})", "Excel match"],
        ["Su_t_m2", "avg(su1_t_m2, su2_t_m2)", r"S_u=\mathrm{AVERAGE}(s_{u1},s_{u2})", "Excel match"],
        ["alpha_clay", "0.93 * exp(-0.0536 * Su_t_m2)", r"\alpha=0.93e^{-0.0536S_u}", "default"],
    ], columns=["parameter", "equation", "latex", "notes"])

def ensure_equation_columns(eq_df):
    """Return a clean equation table.

    Required columns:
    - parameter: internal parameter name used by the solver
    - equation: simple Python-style equation used for calculation

    Optional columns are kept for user guidance and reporting:
    - latex: equation preview shown in the interface
    - unit: expected output unit of the parameter
    - soil_type: clay/sand/all/help text
    - source: design standard, textbook, company method, or advisor note
    - notes: practical comments for the user
    """
    if eq_df is None or eq_df.empty or "parameter" not in eq_df.columns or "equation" not in eq_df.columns:
        return default_equations_df()
    out = eq_df.copy()
    for c in ["latex", "unit", "soil_type", "source", "notes"]:
        if c not in out.columns:
            out[c] = ""
    return out[["parameter", "equation", "latex", "unit", "soil_type", "source", "notes"]]

def equations_df_to_dict(eq_df):
    eq_df = ensure_equation_columns(eq_df)
    return {clean_text(r["parameter"]): clean_text(r["equation"]) for _, r in eq_df.iterrows()}


def equation_parameter_reference_df():
    return pd.DataFrame([
        ["su2_t_m2", "Estimate Su from SPT", "0.52 * SPT_N", "t/m²", "Clay", "Can be edited for local correlation"],
        ["Su_t_m2", "Final Su used in clay equations", "avg(su1_t_m2, su2_t_m2)", "t/m²", "Clay", "avg ignores blank/nan values"],
        ["alpha_clay", "Adhesion factor", "0.93 * exp(-0.0536 * Su_t_m2)", "-", "Clay", "Used in fs_clay_t_m2"],
        ["fs_clay_t_m2", "Unit shaft resistance in clay", "alpha_clay * Su_t_m2", "t/m²", "Clay", "Active in clay layers"],
        ["CN", "SPT overburden correction", "0.77 * log10(20 * 95.76 / max(sigma_v02_t_m2 * g_factor, 0.0001))", "-", "Sand", "Can be changed to another standard"],
        ["N_corr", "Corrected SPT", "CN * SPT_N", "blows/ft", "Sand", "Used for phi correlation"],
        ["phi_deg", "Friction angle", "27.1 + 0.3 * N_corr - 0.00054 * N_corr**2", "degree", "Sand", "Use degree in equation output"],
        ["K0", "At-rest pressure coefficient", "1 - sin(radians(phi_deg))", "-", "Sand", "Internal variable"],
        ["Ks_driven", "Driven-pile lateral coefficient", "1.5 * K0", "-", "Sand", "Driven pile only"],
        ["Ks_bored", "Bored-pile lateral coefficient", "0.85 * K0", "-", "Sand", "Bored pile only"],
        ["Ks", "Selected lateral coefficient", "Ks_driven if pile_type_code > 1.5 else Ks_bored", "-", "Sand", "Bored=1, Driven=2"],
        ["delta_deg", "Interface friction angle", "0.75 * phi_deg", "degree", "Sand", "Used in beta"],
        ["beta_sand", "Beta factor", "Ks * tan(radians(delta_deg))", "-", "Sand", "Used in fs_sand_t_m2"],
        ["fs_sand_t_m2", "Unit shaft resistance in sand", "beta_sand * sigma_v02_t_m2", "t/m²", "Sand", "Active in sand layers"],
        ["Nc", "Clay bearing factor", "9", "-", "Clay", "Used in tip resistance"],
        ["Nq_driven", "Driven-pile Nq", "exp(pi * tan(radians(phi_deg))) * tan(radians(45 + phi_deg/2))**2 * (1 + tan(radians(phi_deg)))", "-", "Sand", "Advisor note: trigonometric phi converted from degree to radian"],
        ["Nq_bored", "Bored-pile Nq", "0.6934 * exp(0.0974 * phi_deg)", "-", "Sand", "Advisor note: empirical relation; phi_deg is used directly in degrees"],
        ["Nq", "Selected Nq", "Nq_driven if pile_type_code > 1.5 else Nq_bored", "-", "Sand", "Bored=1, Driven=2"],
        ["Ngamma", "Bearing capacity factor", "2 * (Nq + 1) * tan(radians(phi_deg))", "-", "Sand", "Available for custom equations"],
        ["qe_clay_t_m2", "Clay tip resistance", "Nc * Su_t_m2", "t/m²", "Clay", "Used at pile tip layer"],
        ["qe_sand_t_m2", "Sand tip resistance", "(Nq - 1) * sigma_v01_t_m2", "t/m²", "Sand", "Used at pile tip layer"],
    ], columns=["parameter", "meaning", "example_equation", "unit", "soil_type", "note"])

def equation_variable_reference_df():
    return pd.DataFrame([
        ["su1_t_m2", "Input Su from borehole CSV", "t/m²", "Clay input"],
        ["SPT_N", "SPT N-value from borehole CSV", "blows/ft", "Sand input; blank SPT becomes 0 and SPT_N_missing=1"],
        ["SPT_N_missing", "Flag for blank SPT", "0 or 1", "Useful for conditional equations"],
        ["sigma_v01_t_m2", "Vertical effective stress at layer bottom", "t/m²", "Uses water table and unit weight"],
        ["sigma_v02_t_m2", "Vertical effective stress at layer midpoint", "t/m²", "Often used for shaft friction"],
        ["gamma_t_m3", "Total unit weight", "t/m³", "From borehole layer CSV"],
        ["gamma_eff_t_m3", "Effective unit weight used by app", "t/m³", "gamma or gamma - gamma_w depending on water table"],
        ["depth_from_m", "Top depth of current layer", "m", "Layer coordinate"],
        ["depth_to_m", "Bottom depth of current layer", "m", "Layer coordinate"],
        ["H_m", "Layer thickness", "m", "bottom - top"],
        ["D_m", "Pile size/diameter/width", "m", "From pile CSV"],
        ["pile_size_m", "Same as D_m", "m", "Alias for convenience"],
        ["pile_type_code", "Bored=1, Driven=2", "-", "Use for if-else equations"],
        ["g_factor", "1 ton force conversion", "kN/ton", "User chooses 9.81 or 10"],
        ["pi", "Mathematical pi", "-", "Available constant"],
        ["e", "Euler's number", "-", "Available constant"],
        ["nan", "Blank value", "-", "Used when a value should be ignored by avg"],
    ], columns=["variable", "meaning", "unit", "notes"])

def equation_function_reference_df():
    return pd.DataFrame([
        ["exp(x)", "exponential", "exp(-0.05 * Su_t_m2)"],
        ["log(x)", "natural logarithm", "log(x)"],
        ["log10(x)", "base-10 logarithm", "log10(20 / sigma_v02_t_m2)"],
        ["sqrt(x)", "square root", "sqrt(SPT_N)"],
        ["sin(x), cos(x), tan(x)", "trigonometry; x must be radians", "tan(radians(phi_deg))"],
        ["radians(x)", "convert degree to radian", "sin(radians(phi_deg))"],
        ["degrees(x)", "convert radian to degree", "degrees(atan_value)"],
        ["min(a,b), max(a,b)", "limit values", "max(sigma_v02_t_m2, 0.0001)"],
        ["abs(x)", "absolute value", "abs(depth_to_m - depth_from_m)"],
        ["avg(a,b,...)", "average values and ignore nan", "avg(su1_t_m2, su2_t_m2)"],
        ["x**2", "power", "N_corr**2"],
        ["a if condition else b", "conditional equation", "Ks_driven if pile_type_code > 1.5 else Ks_bored"],
    ], columns=["syntax", "meaning", "example"])

def equation_examples_df():
    return pd.DataFrame([
        ["Clay alpha from Su", "alpha_clay", "0.93 * exp(-0.0536 * Su_t_m2)", r"\alpha=0.93e^{-0.0536S_u}"],
        ["Clay shaft friction", "fs_clay_t_m2", "alpha_clay * Su_t_m2", r"f_s=\alpha S_u"],
        ["Use different Ks for bored/driven", "Ks", "Ks_driven if pile_type_code > 1.5 else Ks_bored", r"K_s=K_{s,D}\ \mathrm{or}\ K_{s,B}"],
        ["Sand beta", "beta_sand", "Ks * tan(radians(delta_deg))", r"\beta=K_s\tan\delta"],
        ["Limit stress to avoid log error", "CN", "0.77 * log10(20 * 95.76 / max(sigma_v02_t_m2 * g_factor, 0.0001))", r"C_N=0.77\log_{10}(\cdots)"],
        ["Use a constant Nc", "Nc", "9", r"N_c=9"],
    ], columns=["case", "parameter", "simple_input", "latex_preview"])


def nq_formula_check_df():
    rows = []
    for phi in [20, 25, 30, 35, 40]:
        phi_rad = math.radians(phi)
        nqd = math.exp(math.pi * math.tan(phi_rad)) * math.tan(math.radians(45 + phi/2))**2 * (1 + math.tan(phi_rad))
        nqb = 0.6934 * math.exp(0.0974 * phi)
        rows.append({
            "phi_deg": phi,
            "tan_phi": math.tan(phi_rad),
            "Nq_driven": nqd,
            "Nq_bored": nqb,
        })
    return pd.DataFrame(rows)

def validate_equation_set(eq_df):
    """Evaluate the equation set on representative clay and sand sample values."""
    eqs = equations_df_to_dict(eq_df)
    base_vars = {
        "layer_no": 1, "depth_from_m": 5.0, "depth_to_m": 6.0, "depth_m": 5.0, "H_m": 1.0,
        "gamma_t_m3": 1.8, "gamma_eff_t_m3": 0.8, "sigma_v01_t_m2": 8.0, "sigma_v02_t_m2": 7.5,
        "SPT_N": 15.0, "SPT_N_missing": 0, "su1_t_m2": 4.0, "pile_type_code": 2.0,
        "D_m": 0.6, "pile_size_m": 0.6, "g_factor": TON_TO_KN,
    }
    variables = base_vars.copy()
    rows = []
    ok_all = True
    for key in EVAL_ORDER:
        expr = eqs.get(key, "")
        if clean_text(expr, "") == "":
            rows.append({"parameter": key, "status": "Missing", "value": "", "message": "No equation found for this parameter."})
            ok_all = False
            continue
        try:
            raw_value = SafeEval(variables).eval(expr)
            value = float(raw_value)
            if math.isnan(value) or math.isinf(value):
                if key == "su2_t_m2":
                    value = np.nan
                else:
                    raise ValueError("Equation returned nan or infinity.")
            variables[key] = value
            rows.append({"parameter": key, "status": "OK", "value": value, "message": ""})
        except Exception as e:
            variables[key] = 0.0
            rows.append({"parameter": key, "status": "Error", "value": "", "message": str(e)})
            ok_all = False
    return ok_all, pd.DataFrame(rows)

EVAL_ORDER = [
    "su2_t_m2", "Su_t_m2", "alpha_clay", "CN", "N_corr", "phi_deg", "K0",
    "Ks_driven", "Ks_bored", "Ks", "delta_deg", "beta_sand", "fs_clay_t_m2", "fs_sand_t_m2",
    "Nc", "Nq_driven", "Nq_bored", "Nq", "Ngamma", "qe_clay_t_m2", "qe_sand_t_m2",
]

# -----------------------------
# Input normalization
# -----------------------------
def required_pile_columns():
    return ["pile_id","boring_no","pile_type","pile_section","pile_size","pile_perimeter","pile_cross_section_area","pile_top_elevation","water_table_level","unit_weight_of_water","pile_fc_ksc","factor_of_safety_FS","pile_length_m","length_increment","required_load_value","required_load_unit","display_force_unit"]

def required_layer_columns():
    return ["boring_no","top_depth_m","bottom_depth_m","gamma_t_m3","soil_type","soil_name","su1_t_m2","spt_N"]

def normalize_pile_case(row):
    case = {k: v for k, v in dict(row).items() if not pd.isna(v)}
    # flexible column names
    alias = {
        "pile_size_m": "pile_size", "pile_diameter_m": "pile_size", "diameter_m": "pile_size",
        "pile_area_m2": "pile_cross_section_area", "area_m2": "pile_cross_section_area",
        "perimeter_m": "pile_perimeter", "pile_perimeter_m": "pile_perimeter",
        "FS": "factor_of_safety_FS", "safety_factor": "factor_of_safety_FS",
        "length_m": "pile_length_m", "required_load_kN": "required_load_value", "required_load_t": "required_load_value",
        "water_table_m": "water_table_level", "gamma_w_t_m3": "unit_weight_of_water",
        "fc_ksc": "pile_fc_ksc", "concrete_fc_ksc": "pile_fc_ksc", "concrete_strength_ksc": "pile_fc_ksc", "fck_ksc": "pile_fc_ksc",
    }
    for old, new in alias.items():
        if old in case and new not in case:
            case[new] = case[old]
    if "required_load_kN" in row.index and not pd.isna(row.get("required_load_kN")):
        case["required_load_unit"] = "kN"
    if "required_load_t" in row.index and not pd.isna(row.get("required_load_t")):
        case["required_load_unit"] = "ton"
    # Normalize supported pile type and section codes for clearer results.
    pt_norm = normalize_pile_type_text(case.get("pile_type", ""))
    if pt_norm:
        case["pile_type"] = pt_norm
    sec_norm = normalize_section_text(case.get("pile_section", ""))
    if sec_norm:
        case["pile_section"] = sec_norm
    if "display_force_unit" not in case or clean_text(case.get("display_force_unit")) == "":
        case["display_force_unit"] = case.get("required_load_unit", "kN")
    return case

def pile_type_code(case):
    normalized = normalize_pile_type_text(case.get("pile_type", case.get("pile_type_excel", "B")))
    if normalized == "D":
        return 2
    if normalized == "B":
        return 1
    raise ValueError("Unsupported pile_type. This version supports only B/Bored and D/Driven.")

def soil_family(row):
    t = clean_text(row.get("soil_type", "")).upper()
    n = clean_text(row.get("soil_name", "")).upper()
    s = t + " " + n
    if t.startswith("C") or "CLAY" in s:
        return "C"
    if t.startswith("S") or "SAND" in s:
        return "S"
    if to_float(row.get("spt_N"), 0) > 0:
        return "S"
    if to_float(row.get("su1_t_m2"), 0) > 0:
        return "C"
    return "O"

def validate_inputs(pile_df, layers_df):
    msgs = []
    missing_pile = [c for c in ["pile_id", "boring_no", "pile_type", "pile_section", "pile_length_m"] if c not in pile_df.columns]
    missing_layers = [c for c in ["boring_no", "top_depth_m", "bottom_depth_m", "gamma_t_m3", "soil_type"] if c not in layers_df.columns]
    if missing_pile:
        msgs.append("Pile CSV missing important columns: " + ", ".join(missing_pile))
    if missing_layers:
        msgs.append("Borehole CSV missing important columns: " + ", ".join(missing_layers))
    if msgs:
        return msgs

    # Only bored and driven piles are allowed in this version.
    invalid_pt = []
    for i, v in pile_df["pile_type"].items():
        if normalize_pile_type_text(v) == "":
            invalid_pt.append(f"row {i+1}: {v}")
    if invalid_pt:
        msgs.append("Unsupported pile_type. Use only B/Bored or D/Driven. Check: " + "; ".join(invalid_pt[:8]))

    # Section validation and geometry-related checks.
    invalid_sec = []
    for i, v in pile_df["pile_section"].items():
        if normalize_section_text(v) == "":
            invalid_sec.append(f"row {i+1}: {v}")
    if invalid_sec:
        msgs.append("Unknown pile_section. Use C, S, H/HEX, SC, SS, or CUSTOM. Check: " + "; ".join(invalid_sec[:8]))

    for i, row in pile_df.iterrows():
        sec = normalize_section_text(row.get("pile_section", ""))
        if sec in ["SC", "SS"]:
            if "inner_size" not in pile_df.columns or pd.isna(row.get("inner_size")) or to_float(row.get("inner_size"), 0) <= 0:
                msgs.append(f"row {i+1}: {sec} hollow section requires inner_size.")
            elif to_float(row.get("inner_size"), 0) >= to_float(row.get("pile_size"), 0):
                msgs.append(f"row {i+1}: inner_size must be smaller than pile_size.")
        if sec == "CUSTOM":
            if to_float(row.get("pile_perimeter"), 0) <= 0 or to_float(row.get("pile_cross_section_area"), 0) <= 0:
                msgs.append(f"row {i+1}: CUSTOM section requires pile_perimeter and pile_cross_section_area.")
        if to_float(row.get("pile_length_m"), 0) <= 0:
            msgs.append(f"row {i+1}: pile_length_m must be greater than 0.")
        if "required_load_unit" in pile_df.columns and clean_text(row.get("required_load_unit"), "").lower() not in ["", "t", "ton", "tons", "tonne", "tonnes", "kn"]:
            msgs.append(f"row {i+1}: required_load_unit must be ton or kN.")

    if "boring_no" in pile_df.columns and "boring_no" in layers_df.columns:
        missing_bh = sorted(set(pile_df["boring_no"].astype(str)) - set(layers_df["boring_no"].astype(str)))
        if missing_bh:
            msgs.append("boring_no not found in layers: " + ", ".join(missing_bh[:10]))

    # Borehole layer checks.
    for i, row in layers_df.iterrows():
        if to_float(row.get("bottom_depth_m"), 0) <= to_float(row.get("top_depth_m"), 0):
            msgs.append(f"layer row {i+1}: bottom_depth_m must be greater than top_depth_m.")
        if to_float(row.get("gamma_t_m3"), 0) <= 0:
            msgs.append(f"layer row {i+1}: gamma_t_m3 should be positive.")
        if "spt_N" in layers_df.columns and not pd.isna(row.get("spt_N")) and to_float(row.get("spt_N"), 0) < 0:
            msgs.append(f"layer row {i+1}: spt_N cannot be negative.")
        if "su1_t_m2" in layers_df.columns and not pd.isna(row.get("su1_t_m2")) and to_float(row.get("su1_t_m2"), 0) < 0:
            msgs.append(f"layer row {i+1}: su1_t_m2 cannot be negative.")
    return msgs

# -----------------------------
# Excel-compatible calculation
# -----------------------------
def calculate_excel_style_table(layers, pile_case, eqs):
    case = pile_case
    layers = layers.copy().sort_values(["top_depth_m", "bottom_depth_m"]).reset_index(drop=True)
    section = case.get("pile_section", "C")
    size = to_float(case.get("pile_size"), 0.6)
    perim_auto, area_auto = shape_geometry(section, size, to_float(case.get("inner_size", 0), 0))
    perimeter = to_float(case.get("pile_perimeter"), perim_auto)
    area = to_float(case.get("pile_cross_section_area"), area_auto)
    FS = max(to_float(case.get("factor_of_safety_FS"), 2.5), 1e-9)
    pile_top = to_float(case.get("pile_top_elevation"), 0.0)
    wt = to_float(case.get("water_table_level"), 1.0)
    gamma_w = to_float(case.get("unit_weight_of_water"), 1.0)
    fc_ksc = to_float(case.get("pile_fc_ksc"), 350.0)
    pt_code = pile_type_code(case)

    rows = []
    sigma_v01 = 0.0
    Qs = 0.0

    for i, layer in layers.iterrows():
        top = to_float(layer.get("top_depth_m"), 0.0)
        bottom = to_float(layer.get("bottom_depth_m"), top)
        H = max(bottom - top, 0.0)
        gamma = to_float(layer.get("gamma_t_m3"), 0.0)
        gamma_eff = gamma if top < wt else gamma - gamma_w
        gamma_eff = max(gamma_eff, 0.0)
        gammaH = gamma_eff * H
        sigma_v01 += gammaH
        sigma_v02 = sigma_v01 - gammaH / 2.0
        raw_spt = layer.get("spt_N")
        SPT_N_missing = 1 if pd.isna(raw_spt) or raw_spt == "" else 0
        SPT_N = 0.0 if SPT_N_missing else to_float(raw_spt, 0.0)
        su1 = np.nan if pd.isna(layer.get("su1_t_m2")) else to_float(layer.get("su1_t_m2"), np.nan)

        variables = {
            "layer_no": i + 1,
            "depth_from_m": top,
            "depth_to_m": bottom,
            "depth_m": top,
            "H_m": H,
            "gamma_t_m3": gamma,
            "gamma_eff_t_m3": gamma_eff,
            "sigma_v01_t_m2": sigma_v01,
            "sigma_v02_t_m2": sigma_v02,
            "SPT_N": SPT_N,
            "SPT_N_missing": SPT_N_missing,
            "su1_t_m2": su1,
            "pile_type_code": pt_code,
            "D_m": size,
            "pile_size_m": size,
            "g_factor": TON_TO_KN,
        }
        for key in EVAL_ORDER:
            default_value = np.nan if key == "su2_t_m2" else 0.0
            variables[key] = safe_eval(eqs.get(key, "0"), variables, default_value)

        fam = soil_family(layer)
        qe_cap_t_m2 = to_float(case.get("qe_cap_t_m2", 500.0), 500.0)
        apply_qe_cap = qe_cap_t_m2 > 0
        qe_clay_raw = max(variables.get("qe_clay_t_m2", 0.0), 0.0)
        qe_sand_raw = max(variables.get("qe_sand_t_m2", 0.0), 0.0)
        qe_clay_used = min(qe_clay_raw, qe_cap_t_m2) if apply_qe_cap else qe_clay_raw
        qe_sand_used = min(qe_sand_raw, qe_cap_t_m2) if apply_qe_cap else qe_sand_raw
        if fam == "C":
            fs_raw = variables.get("fs_clay_t_m2", 0.0)
            qe_raw = qe_clay_raw
            qe = qe_clay_used
        elif fam == "S":
            fs_raw = variables.get("fs_sand_t_m2", 0.0)
            qe_raw = qe_sand_raw
            qe = qe_sand_used
        else:
            fs_raw = 0.0
            qe_raw = 0.0
            qe = 0.0

        fs = 0.0 if top < pile_top else max(fs_raw, 0.0)
        Asi = perimeter * H
        Qsi = fs * Asi
        Qs += Qsi
        soil_name = clean_text(layer.get("soil_name"), "Clay" if fam == "C" else "Sand" if fam == "S" else clean_text(layer.get("soil_type"), "Soil"))
        rows.append({
            "No.": i + 1,
            "Depth from (m)": top,
            "Depth to (m)": bottom,
            "g (t/m³)": gamma,
            "Soil type": fam if fam != "O" else clean_text(layer.get("soil_type"), "O"),
            "Soil name": soil_name,
            "Input su1 (t/m²)": 0.0 if pd.isna(su1) else su1,
            "SPT (blow/ft)": SPT_N,
            "H (m)": H,
            "g' (t/m³)": gamma_eff,
            "g'H (t/m²)": gammaH,
            "s'v01 bottom (t/m²)": sigma_v01,
            "s'v02 mid (t/m²)": sigma_v02,
            "su2=0.52SPT (t/m²)": variables.get("su2_t_m2", 0.0),
            "su (t/m²)": variables.get("Su_t_m2", 0.0),
            "alpha": variables.get("alpha_clay", 0.0),
            "CN": variables.get("CN", 0.0),
            "N'=CN*SPT": variables.get("N_corr", 0.0),
            "phi (deg)": variables.get("phi_deg", 0.0),
            "K0": variables.get("K0", 0.0),
            "Ks driven": variables.get("Ks_driven", 0.0),
            "Ks bored": variables.get("Ks_bored", 0.0),
            "Ks": variables.get("Ks", 0.0),
            "delta (deg)": variables.get("delta_deg", 0.0),
            "beta": variables.get("beta_sand", 0.0),
            "fsc=alpha*su (t/m²)": variables.get("fs_clay_t_m2", 0.0),
            "fss=beta*s'v02 (t/m²)": variables.get("fs_sand_t_m2", 0.0),
            "fsc or fss (t/m²)": fs_raw,
            "fs active (t/m²)": fs,
            "Asi (m²)": Asi,
            "Qsi (ton)": Qsi,
            "Qs cumulative (ton)": Qs,
            "Nq driven": variables.get("Nq_driven", 0.0),
            "Nq bored": variables.get("Nq_bored", 0.0),
            "Nq": variables.get("Nq", 0.0),
            "Ngamma": variables.get("Ngamma", 0.0),
            "qec=9su (t/m²)": qe_clay_used,
            "qes=(Nq-1)s'v01 (t/m²)": qe_sand_used,
            "qe raw before cap (t/m²)": qe_raw,
            "qe cap (t/m²)": qe_cap_t_m2 if apply_qe_cap else 0.0,
            "qe (t/m²)": qe,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    qe = df["qe (t/m²)"].astype(float).tolist()
    qe_ave = []
    qe_windows = []
    # Excel-style q_e average: one row above + current row + two rows below.
    # At the top and bottom of the profile, only available rows are averaged.
    for i in range(len(qe)):
        start_i = max(0, i - 1)
        end_i = min(len(qe), i + 3)
        vals = [qe[j] for j in range(start_i, end_i) if not pd.isna(qe[j])]
        qe_ave.append(sum(vals)/len(vals) if vals else 0.0)
        qe_windows.append(f"rows {start_i + 1}-{end_i}")
    df["qe,ave window"] = qe_windows
    df["qe,ave (t/m²)"] = qe_ave
    df["Qe (ton)"] = df["qe,ave (t/m²)"] * area
    df["Qult (ton)"] = df["Qs cumulative (ton)"] + df["Qe (ton)"]
    df["Qall (ton)"] = df["Qult (ton)"] / FS
    df["fc,all (t/m²)"] = 0.25 * fc_ksc * 10.0
    df["Qc,all (ton)"] = area * df["fc,all (t/m²)"]
    df["Concrete status"] = np.where(df["Qall (ton)"] < df["Qc,all (ton)"], "OK", "NOK")
    required_ton = force_to_kn(case.get("required_load_value", 0.0), case.get("required_load_unit", "kN")) / TON_TO_KN
    df["Required load (ton)"] = required_ton
    df["Load safety ratio"] = np.where(required_ton > 0, df["Qall (ton)"] / required_ton, 0.0)
    df["Load status"] = np.where((required_ton > 0) & (df["Qall (ton)"] >= required_ton), "OK", "NG")
    return df.replace([np.inf, -np.inf], np.nan).fillna(0.0)

def display_calculation_table(base_df, output_unit="ton", decimals=6):
    u = unit_label(output_unit)
    df = base_df.copy()
    if df.empty:
        return df
    if u == "kN":
        rename = {}
        for c in df.columns:
            if "(ton)" in c:
                df[c] = df[c] * TON_TO_KN
                rename[c] = c.replace("(ton)", "(kN)")
            if "(t/m²)" in c:
                df[c] = df[c] * TON_TO_KN
                rename[c] = c.replace("(t/m²)", "(kPa)")
            if "(t/m³)" in c:
                df[c] = df[c] * TON_TO_KN
                rename[c] = c.replace("(t/m³)", "(kN/m³)")
        df = df.rename(columns=rename)
    return round_for_display(df, decimals)

def result_at_pile_length(base_df, case):
    L = to_float(case.get("pile_length_m"), to_float(case.get("length_m"), 0.0))
    if base_df.empty:
        return {}
    idx = (base_df["Depth from (m)"] - L).abs().idxmin()
    row = base_df.loc[idx].to_dict()
    return row

def run_batch(pile_df, layers_df, equations_df):
    eqs = equations_df_to_dict(equations_df)
    results = []
    errors = []
    for idx, prow in pile_df.iterrows():
        try:
            case = normalize_pile_case(prow)
            if normalize_pile_type_text(case.get("pile_type", "")) == "":
                raise ValueError("Unsupported pile_type. Use only B/Bored or D/Driven.")
            if normalize_section_text(case.get("pile_section", "")) == "":
                raise ValueError("Unsupported pile_section. Use C, S, H/HEX, SC, SS, or CUSTOM.")
            bh = clean_text(case.get("boring_no"))
            pile_id = clean_text(case.get("pile_id"), f"P-{idx+1:04d}")
            L = to_float(case.get("pile_length_m"), 0.0)
            selected_layers = layers_df[layers_df["boring_no"].astype(str) == bh].copy()
            if selected_layers.empty:
                raise ValueError(f"No borehole layers found for {bh}")
            table = calculate_excel_style_table(selected_layers, case, eqs)
            r = result_at_pile_length(table, case)
            out_unit = unit_label(case.get("display_force_unit", case.get("required_load_unit", "kN")))
            req_kn = force_to_kn(case.get("required_load_value", 0.0), case.get("required_load_unit", "kN"))
            qall_kn = r.get("Qall (ton)", 0.0) * TON_TO_KN
            results.append({
                "pile_id": pile_id,
                "boring_no": bh,
                "pile_type": case.get("pile_type", ""),
                "pile_section": case.get("pile_section", ""),
                "pile_size_m": to_float(case.get("pile_size"), 0.0),
                "pile_length_m": L,
                "output_unit": out_unit,
                "Qs": force_from_ton(r.get("Qs cumulative (ton)", 0.0), out_unit),
                "Qb": force_from_ton(r.get("Qe (ton)", 0.0), out_unit),
                "Qult": force_from_ton(r.get("Qult (ton)", 0.0), out_unit),
                "Qall": force_from_ton(r.get("Qall (ton)", 0.0), out_unit),
                "required_load": force_from_ton(req_kn / TON_TO_KN, out_unit),
                "safety_ratio": qall_kn / req_kn if req_kn > 0 else 0.0,
                "load_status": "OK" if req_kn > 0 and qall_kn >= req_kn else "NG",
                "concrete_status": r.get("Concrete status", ""),
                "selected_depth_m": r.get("Depth from (m)", 0.0),
            })
        except Exception as e:
            errors.append({"row": idx+1, "pile_id": prow.get("pile_id", ""), "error": str(e)})
    return display_no_missing(pd.DataFrame(results).round(4)), pd.DataFrame(errors)

# -----------------------------
# Plotting
# -----------------------------
def tight_x_limits(ax, values, include_zero=True, pad=0.08):
    s = pd.Series(values).replace([np.inf, -np.inf], np.nan).dropna()
    if s.empty:
        return
    vmin, vmax = float(s.min()), float(s.max())
    if include_zero:
        vmin = min(vmin, 0.0)
    if abs(vmax - vmin) < 1e-12:
        vmax = vmin + 1.0
    m = (vmax - vmin) * pad
    ax.set_xlim(vmin - m, vmax + m)

def soil_color_map(layers):
    keys = []
    for _, r in layers.iterrows():
        fam = soil_family(r)
        label = "Clay" if fam == "C" else "Sand" if fam == "S" else clean_text(r.get("soil_name"), "Other")
        keys.append(label)
    base = {"Clay":"#cfe8ff", "Sand":"#ffe8a3", "Silt":"#d9f99d", "Fill":"#e5e7eb", "Other":"#e5e7eb"}
    palette = ["#cfe8ff", "#ffe8a3", "#d9f99d", "#fecaca", "#ddd6fe", "#bae6fd", "#fef3c7", "#e5e7eb"]
    cmap = {}
    for k in keys:
        cmap[k] = base.get(k, palette[len(cmap) % len(palette)]) if k not in cmap else cmap[k]
    return cmap

def section_display_name(section):
    sec = clean_text(section, "C").upper()
    names = {
        "C": "Circle", "S": "Square", "H": "Hexagon", "SC": "Hollow spun", "SS": "Hollow square", "I": "I-section", "CUSTOM": "Custom"
    }
    return names.get(sec, sec)


def draw_section_shape(ax, case, face="white", text="#0f172a", spine="#475569"):
    sec = normalize_section_text(case.get("pile_section", "")) or clean_text(case.get("pile_section", "C")).upper()
    D = max(to_float(case.get("pile_size", 0.6), 0.6), 0.05)
    inner = max(to_float(case.get("inner_size", 0.0), 0.0), 0.0)
    ax.set_facecolor(face)
    ax.set_aspect('equal')
    ax.set_xlim(-1.15, 1.15)
    ax.set_ylim(-1.15, 1.15)
    ax.axis('off')
    if sec == 'C':
        ax.add_patch(Circle((0,0), 0.75, facecolor="#dbeafe", edgecolor=spine, lw=2))
    elif sec == 'SC':
        ax.add_patch(Circle((0,0), 0.78, facecolor="#dbeafe", edgecolor=spine, lw=2))
        ax.add_patch(Circle((0,0), 0.40, facecolor=face, edgecolor=spine, lw=1.6))
    elif sec == 'S':
        ax.add_patch(Rectangle((-0.72,-0.72),1.44,1.44, facecolor="#dbeafe", edgecolor=spine, lw=2))
    elif sec == 'SS':
        ax.add_patch(Rectangle((-0.78,-0.78),1.56,1.56, facecolor="#dbeafe", edgecolor=spine, lw=2))
        ax.add_patch(Rectangle((-0.40,-0.40),0.80,0.80, facecolor=face, edgecolor=spine, lw=1.6))
    elif sec == 'H':
        ax.add_patch(RegularPolygon((0,0), numVertices=6, radius=0.82, orientation=np.pi/6, facecolor="#dbeafe", edgecolor=spine, lw=2))
    elif sec == 'I':
        pts = np.array([[-0.78,0.78],[0.78,0.78],[0.78,0.48],[0.20,0.48],[0.20,-0.48],[0.78,-0.48],[0.78,-0.78],[-0.78,-0.78],[-0.78,-0.48],[-0.20,-0.48],[-0.20,0.48],[-0.78,0.48]])
        ax.add_patch(Polygon(pts, closed=True, facecolor="#dbeafe", edgecolor=spine, lw=2))
    else:
        ax.add_patch(Rectangle((-0.72,-0.72),1.44,1.44, facecolor="#dbeafe", edgecolor=spine, lw=2))
    ax.text(0, 0.98, section_display_name(sec), ha='center', va='bottom', color=text, fontsize=9, fontweight='bold')
    ax.text(0, -1.02, f"size = {D:g} m", ha='center', va='top', color=text, fontsize=8)
    if sec in ['SC','SS'] and inner>0:
        ax.text(0, -1.18, f"inner = {inner:g} m", ha='center', va='top', color=text, fontsize=7.6)

def plot_aligned_panels(layers, table_ton, case, output_unit="ton", board_theme="Bright"):
    u = unit_label(output_unit)
    df = table_ton.copy()
    max_depth = max(to_float(layers["bottom_depth_m"].max(), 0), to_float(df["Depth to (m)"].max(), 0) if not df.empty else 0)
    face = "#0b1220" if str(board_theme).lower().startswith("dark") else "white"
    text = "#e5e7eb" if str(board_theme).lower().startswith("dark") else "#0f172a"
    grid = "#334155" if str(board_theme).lower().startswith("dark") else "#cbd5e1"
    spine = "#94a3b8" if str(board_theme).lower().startswith("dark") else "#475569"

    fig, axes = plt.subplots(1, 6, figsize=(22.5, 8.5),
                             gridspec_kw={"width_ratios": [1.45, 0.9, 1.15, 1.15, 1.7, 1.7], "wspace": 0.12}, facecolor=face)
    for ax in axes:
        ax.set_facecolor(face)
        for sp in ax.spines.values():
            sp.set_color(spine)
    depth_axes = [axes[0], axes[2], axes[3], axes[4], axes[5]]
    for ax in depth_axes:
        ax.set_ylim(max_depth, 0)
        ax.tick_params(colors=text, labelsize=8)
        ax.grid(True, color=grid, alpha=.50, linewidth=.5)
        for _, layer in layers.iterrows():
            ax.axhline(to_float(layer.get("top_depth_m")), color=grid, lw=.45, alpha=.7)
            ax.axhline(to_float(layer.get("bottom_depth_m")), color=grid, lw=.45, alpha=.7)

    ax0 = axes[0]
    cmap = soil_color_map(layers)
    legend_handles = [
        Patch(facecolor=cmap.get("Clay", "#c8a97e"), edgecolor=spine, label="Clay"),
        Patch(facecolor=cmap.get("Sand", "#f6d365"), edgecolor=spine, label="Sand"),
        Patch(facecolor="#d1d5db", edgecolor=spine, label="Other"),
    ]
    for _, layer in layers.iterrows():
        top = to_float(layer.get("top_depth_m"))
        bottom = to_float(layer.get("bottom_depth_m"))
        fam = soil_family(layer)
        label = "Clay" if fam == "C" else "Sand" if fam == "S" else clean_text(layer.get("soil_name"), "Other")
        col = cmap.get(label, "#d1d5db")
        ax0.fill_betweenx([top, bottom], 0, 1, color=col, edgecolor=spine, linewidth=.65, alpha=0.95)
        if bottom-top >= max(max_depth/45, 0.6):
            ax0.text(0.14, (top+bottom)/2, label, ha='left', va='center', fontsize=7.4, color="#0f172a")
    pile_L = to_float(case.get('pile_length_m', max_depth*0.65), max_depth*0.65)
    pile_D = to_float(case.get('pile_size', 0.6), 0.6)
    x1, x2 = 0.40, 0.60
    ax0.add_patch(Rectangle((x1, 0), x2-x1, pile_L, facecolor="#cbd5e1", edgecolor=spine, lw=1.8, zorder=4))
    req_txt = f"Q applied = {to_float(case.get('required_load_value',0),0):g} {clean_text(case.get('required_load_unit','ton'))}"
    ax0.annotate('', xy=((x1+x2)/2, 0.2), xytext=((x1+x2)/2, -0.7), arrowprops=dict(arrowstyle='simple', color='#2563eb'))
    ax0.text(0.66, 0.25, req_txt, color='#2563eb', fontsize=7.8, va='bottom', ha='left', fontweight='bold')
    # Shaft friction (Qs) shown along the pile shaft direction
    ax0.annotate('', xy=(x1-0.04, pile_L*0.44), xytext=(x1-0.04, pile_L*0.56), arrowprops=dict(arrowstyle='->', color='#38bdf8', lw=2))
    ax0.annotate('', xy=(x2+0.04, pile_L*0.44), xytext=(x2+0.04, pile_L*0.56), arrowprops=dict(arrowstyle='->', color='#38bdf8', lw=2))
    ax0.text(0.12, pile_L*0.50, 'Qs', color='#38bdf8', fontsize=8, va='center', ha='center', fontweight='bold', rotation=0)
    ax0.annotate('', xy=((x1+x2)/2, pile_L-0.25), xytext=((x1+x2)/2, min(pile_L+max(3.5, max_depth*0.22), max_depth*1.04)), arrowprops=dict(arrowstyle='->', color='#f97316', lw=2.4), annotation_clip=False)
    ax0.text((x1+x2)/2, pile_L + max(1.35, max_depth*0.055), 'Qe', color='#f97316', fontsize=8, va='top', ha='center', fontweight='bold', clip_on=False)
    ax0.annotate('', xy=(0.93, 0), xytext=(0.93, pile_L), arrowprops=dict(arrowstyle='<->', color=spine, lw=1.2))
    ax0.text(0.945, pile_L/2, f"L = {pile_L:g} m", rotation=90, va='center', ha='left', color=text, fontsize=8)
    ax0.annotate('', xy=(x1, min(pile_L*0.2, max_depth*0.14)), xytext=(x2, min(pile_L*0.2, max_depth*0.14)), arrowprops=dict(arrowstyle='<->', color=spine, lw=1.2))
    ax0.text((x1+x2)/2, min(pile_L*0.2, max_depth*0.14)-0.25, f"D = {pile_D:g} m", ha='center', va='top', color=text, fontsize=8)
    wt = to_float(case.get("water_table_level", np.nan), np.nan)
    if not pd.isna(wt) and 0 <= wt <= max_depth:
        ax0.axhline(wt, color="#0284c7", linewidth=2.2, zorder=7)
        ax0.scatter([.1, .22, .34], [wt, wt, wt], marker="v", color="#0284c7", s=36, zorder=8)
        ax0.text(.36, wt, "Water table", ha="left", va="bottom", fontsize=7.6, color="#0284c7", fontweight="bold", zorder=8)
    ax0.legend(handles=legend_handles, fontsize=7, loc='best', framealpha=0.95)
    ax0.set_xlim(0, 1)
    ax0.set_xticks([])
    ax0.set_ylabel("Depth (m)", color=text)
    ax0.set_title("Soil profile + pile", color=text, fontsize=10, fontweight="bold", pad=14)

    draw_section_shape(axes[1], case, face=face, text=text, spine=spine)

    y = df["Depth from (m)"]
    fs = df["fs active (t/m²)"].apply(lambda x: stress_from_t_m2(x, u))
    qe = df["qe,ave (t/m²)"].apply(lambda x: stress_from_t_m2(x, u))
    qs = df["Qs cumulative (ton)"].apply(lambda x: force_from_ton(x, u))
    qbe = df["Qe (ton)"].apply(lambda x: force_from_ton(x, u))
    qult = df["Qult (ton)"].apply(lambda x: force_from_ton(x, u))
    qall = df["Qall (ton)"].apply(lambda x: force_from_ton(x, u))
    qcall = df["Qc,all (ton)"].apply(lambda x: force_from_ton(x, u)) if "Qc,all (ton)" in df.columns else pd.Series([], dtype=float)
    req = force_to_kn(case.get("required_load_value", 0), case.get("required_load_unit", "kN"))
    req_out = req if u == "kN" else req / TON_TO_KN


    axes[2].plot(fs, y, color="#ef4444", linewidth=1.55, label='fs')
    axes[2].set_xlabel(f"fs ({stress_unit(u)})", color=text)
    axes[2].set_title("Unit skin friction", color=text, fontsize=10, fontweight="bold")
    tight_x_limits(axes[2], fs)
    axes[2].legend(fontsize=7, loc='best', framealpha=0.92)

    axes[3].plot(qe, y, color="#60a5fa", linewidth=1.55, label='qe,ave')
    axes[3].set_xlabel(f"qe,ave ({stress_unit(u)})", color=text)
    axes[3].set_title("Unit end bearing", color=text, fontsize=10, fontweight="bold")
    tight_x_limits(axes[3], qe)
    axes[3].legend(fontsize=7, loc='best', framealpha=0.92)

    axes[4].plot(qs, y, color="#22c55e", linewidth=1.55, label="Qs")
    axes[4].plot(qbe, y, color="#f97316", linewidth=1.55, label="Qe")
    axes[4].plot(qult, y, color="#8b5cf6", linewidth=1.55, label="Qult")
    axes[4].set_xlabel(f"Capacity ({cap_unit(u)})", color=text)
    axes[4].set_title("Capacity distribution", color=text, fontsize=10, fontweight="bold")
    axes[4].legend(fontsize=7, loc="best", framealpha=0.92)
    tight_x_limits(axes[4], pd.concat([qs, qbe, qult]))

    axes[5].plot(qall, y, color="#60a5fa", linewidth=1.8, label="Qall")
    if len(qcall) == len(y):
        axes[5].plot(qcall, y, color="#10b981", linewidth=1.7, linestyle="--", label="Qc,all")
    if req_out > 0:
        axes[5].axvline(req_out, color="#ef4444", linewidth=1.7, label="Required")
    axes[5].set_xlabel(f"Capacity ({cap_unit(u)})", color=text)
    axes[5].set_title("Qall / Qc,all / required", color=text, fontsize=10, fontweight="bold")
    axes[5].legend(fontsize=7, loc="best", framealpha=0.92)
    lim_series = [qall]
    if len(qcall) == len(y):
        lim_series.append(qcall)
    if req_out > 0:
        lim_series.append(pd.Series([req_out]))
    tight_x_limits(axes[5], pd.concat(lim_series, ignore_index=True))
    for ax in [axes[2], axes[3], axes[4], axes[5]]:
        ax.tick_params(labelleft=False)
    fig.tight_layout()
    return fig


# -----------------------------
# Shallow foundation helpers
# -----------------------------
def required_shallow_columns():
    return [
        'case_id','footing_shape','analysis_type','B_m','L_m','D_m','t_m','Cx_m','Cy_m','Ox_m','Oy_m','eccentricity_mode','water_table_depth_m',
        'Qc_t','Mcx_tm','Mcy_tm','Hx_t','Hy_t','beta_deg','g1u_t_m3','g1sat_t_m3','c1_t_m2','phi1_deg','g2u_t_m3','g2sat_t_m3','c2_t_m2','phi2_deg',
        'gc_t_m3','gw_t_m3','FS','result_unit'
    ]


def ensure_shallow_columns(df):
    out = df.copy() if df is not None else pd.DataFrame()
    for c in required_shallow_columns():
        if c not in out.columns:
            out[c] = ''
    return out[required_shallow_columns()]


def validate_shallow_inputs(df):
    msgs=[]
    if df is None or df.empty:
        msgs.append('Upload shallow footing cases CSV or use the sample template.')
        return msgs
    req = ['case_id','footing_shape','analysis_type','B_m','L_m','D_m','Qc_t','g1u_t_m3','g1sat_t_m3','g2u_t_m3','g2sat_t_m3','c2_t_m2','phi1_deg','phi2_deg','FS']
    missing = [c for c in req if c not in df.columns]
    if missing:
        msgs.append('Missing shallow columns: ' + ', '.join(missing))
    return msgs


def strip_input_warning_messages(df):
    if df is None or df.empty:
        return []
    try:
        sdf = ensure_shallow_columns(df)
        has_strip = any(sdf['footing_shape'].astype(str).str.upper().str.strip() == 'S')
    except Exception:
        has_strip = False
    if has_strip:
        return ["Strip S: key in L=1, Cy=1, Oy=0, Mcx=0, Hy=0. Cy is read from the CSV. Optional beta_deg can be filled directly. Optional eccentricity_mode=DIRECT uses Ox_m as ex directly; Oy/ey is forced to 0 for strip."]
    return []


def _avg_top_gamma(dw,D,g_u,g_sat):
    dw=float(dw); D=float(D); g_u=float(g_u); g_sat=float(g_sat)
    if dw >= D:
        return g_u
    if dw <= 0:
        return g_sat
    return (g_u*dw + g_sat*(D-dw))/D


def _avg_top_gamma_eff(dw,D,g_u,g_sat,gw):
    dw=float(dw); D=float(D); g_u=float(g_u); g_sat=float(g_sat); gw=float(gw)
    if dw >= D:
        return g_u
    if dw <= 0:
        return g_sat-gw
    return (g_u*dw + (g_sat-gw)*(D-dw))/D


def _avg_bottom_gamma(dw,D,B,g_u,g_sat):
    dw=float(dw); D=float(D); B=float(B); g_u=float(g_u); g_sat=float(g_sat)
    if dw >= D+B:
        return g_u
    if dw <= D:
        return g_sat
    return (g_u*(dw-D) + g_sat*(B-(dw-D)))/B


def _avg_bottom_gamma_eff(dw,D,B,g_u,g_sat,gw):
    dw=float(dw); D=float(D); B=float(B); g_u=float(g_u); g_sat=float(g_sat); gw=float(gw)
    if dw >= D+B:
        return g_u
    if dw <= D:
        return g_sat-gw
    return (g_u*(dw-D) + (g_sat-gw)*(B-(dw-D)))/B


def calc_shallow_case(row, result_unit='ton', equations_df=None):
    c = {k: row.get(k,'') for k in required_shallow_columns()}
    case_id = clean_text(c.get('case_id'), 'SF')
    shape = clean_text(c.get('footing_shape'),'R').upper()
    analysis = clean_text(c.get('analysis_type'),'E').upper()
    B = to_float(c.get('B_m'),0.0)
    L = to_float(c.get('L_m'),1.0)
    D = to_float(c.get('D_m'),0.0)
    t = to_float(c.get('t_m'),0.0)
    Cx = to_float(c.get('Cx_m'),0.0)
    Cy = to_float(c.get('Cy_m'),0.0)
    Ox = to_float(c.get('Ox_m'),0.0)
    Oy = to_float(c.get('Oy_m'),0.0)
    eccentricity_mode = clean_text(c.get('eccentricity_mode'), 'MOMENT').upper()
    dw = to_float(c.get('water_table_depth_m'),999.0)
    Qc = to_float(c.get('Qc_t'),0.0)
    Mcx = to_float(c.get('Mcx_tm'),0.0)
    Mcy = to_float(c.get('Mcy_tm'),0.0)
    Hx = to_float(c.get('Hx_t'),0.0)
    Hy = to_float(c.get('Hy_t'),0.0)
    beta_input_text = clean_text(c.get('beta_deg'), '')
    g1u=to_float(c.get('g1u_t_m3'),0.0)
    g1sat=to_float(c.get('g1sat_t_m3'),0.0)
    c1=to_float(c.get('c1_t_m2'),0.0)
    phi1=to_float(c.get('phi1_deg'),0.0)
    g2u=to_float(c.get('g2u_t_m3'),0.0)
    g2sat=to_float(c.get('g2sat_t_m3'),0.0)
    c2=to_float(c.get('c2_t_m2'),0.0)
    phi2=to_float(c.get('phi2_deg'),0.0)
    gc=to_float(c.get('gc_t_m3'),0.0)
    gw=to_float(c.get('gw_t_m3'),1.0)
    FS=max(to_float(c.get('FS'),3.0),1e-6)
    if shape == 'S':
        # Strip footing is analyzed as a 1 m longitudinal slice.
        # Cy is read from the CSV input. For Excel-style strip TSL, key in Cy = 1.
        L = 1.0
        if clean_text(c.get('Cy_m'), '') == '':
            Cy = 1.0
        Oy = 0.0
        Mcx = 0.0
        Hy = 0.0
    g1ave = _avg_top_gamma(dw,D,g1u,g1sat)
    g1ave_eff = _avg_top_gamma_eff(dw,D,g1u,g1sat,gw)
    g2ave = _avg_bottom_gamma(dw,D,B,g2u,g2sat)
    g2ave_eff = _avg_bottom_gamma_eff(dw,D,B,g2u,g2sat,gw)
    Wsoil = g1ave * max(B*L - Cx*Cy, 0.0) * max(D-t, 0.0)
    Wcolumn = gc * Cx * Cy * max(D-t, 0.0)
    Wslab = t * B * L * gc
    Qt = Wsoil + Wcolumn + Wslab + Qc
    Mtx = (Wcolumn + Qc) * Oy + Hy * D + Mcx
    Mty = (Wcolumn + Qc) * Ox + Hx * D + Mcy
    Ht = math.sqrt(Hx**2 + Hy**2)
    beta_from_load = math.degrees(math.atan(Ht/Qt)) if Qt else 0.0
    if beta_input_text != '':
        beta = max(0.0, to_float(c.get('beta_deg'), beta_from_load))
        beta_source = 'input beta_deg'
    else:
        beta = beta_from_load
        beta_source = 'calculated from Hx_t and Hy_t'
    ex_moment = Mty/Qt if Qt else 0.0
    ey_moment = Mtx/Qt if Qt else 0.0
    if eccentricity_mode in ["DIRECT", "DIRECT_ECCENTRICITY", "INPUT", "USER"]:
        eccentricity_mode = "DIRECT"
        ex = Ox
        ey = Oy
        eccentricity_source = "direct eccentricity from Ox_m and Oy_m"
    else:
        eccentricity_mode = "MOMENT"
        ex = ex_moment
        ey = ey_moment
        eccentricity_source = "moment-derived eccentricity from Ox_m and Oy_m"
    if shape == 'S':
        ey = 0.0
    Bx = B - 2*abs(ex)
    Ly = L - 2*abs(ey)

    # Strip footing: use B' = B - 2ex, while keeping 1 m longitudinal strip length.
    # Rectangular case is unchanged.
    if shape == 'S':
        B_eff = max(Bx, 1e-6)
        L_eff = 1.0
    else:
        B_eff = max(Bx, 1e-6)
        L_eff = max(Ly, 1e-6)
    A_eff = B_eff * L_eff

    # Strip S follows the Excel sheet for lower unit-weight averaging:
    # use original B for g2ave/g'2ave, while B' remains 1 for the bearing term.
    g2_width_for_gamma = B if shape == 'S' else B_eff
    g2ave = _avg_bottom_gamma(dw, D, g2_width_for_gamma, g2u, g2sat)
    g2ave_eff = _avg_bottom_gamma_eff(dw, D, g2_width_for_gamma, g2u, g2sat, gw)

    g1used = g1ave if analysis == 'T' else g1ave_eff
    g2used = g2ave if analysis == 'T' else g2ave_eff
    q = g1used * D
    phi1r = math.radians(phi1)
    phi2r = math.radians(phi2)
    Nq = math.exp(math.pi*math.tan(phi2r)) * math.tan(math.pi/4 + phi2r/2)**2
    Nc = (Nq - 1.0)/math.tan(phi2r) if abs(math.tan(phi2r)) > 1e-12 else 5.14
    Ny = 2*(Nq + 1.0)*math.tan(phi2r)
    ratio = B_eff/max(L_eff,1e-6)

    # Shape factors
    # Strip footing: Fcs = Fqs = Fys = 1.0
    # Rectangular footing: use B'/L' relationship.
    if shape == 'S':
        Fcs = 1.0
        Fqs = 1.0
        Fys = 1.0
    else:
        Fcs = 1.0 + ratio*(Nq/max(Nc,1e-6))
        Fqs = 1.0 + ratio*math.tan(phi2r)
        Fys = 1.0 - 0.4*ratio

    # Depth factors
    # User rule: depth factor uses phi1.
    D_over_B = D/max(B_eff,1e-6)
    depth_argument = D_over_B if D_over_B <= 1.0 else math.atan(D_over_B)
    Fcd = 1.0 + 0.4*depth_argument
    Fqd = 1.0 + 2*math.tan(phi1r)*(1-math.sin(phi1r))**2*depth_argument
    Fyd = 1.0

    # Inclination factors
    # beta is calculated in degrees, so 90 and phi2 are also used in degrees.
    beta_deg = beta
    Fci = (1-beta_deg/90.0)**2
    Fqi = (1-beta_deg/90.0)**2
    Fyi = (1-beta_deg/max(phi2,1e-6))**2 if phi2>0 else 1.0

    # Optional shallow equation library connection.
    # If users load/upload a shallow equation table in Equation Lab, recognized parameters below
    # can update the calculation. If an equation is missing, the app keeps the original value.
    # Strip footing shape factors are forced to 1.0 after custom equations, as required.
    shallow_eqs = shallow_equations_df_to_dict(equations_df) if equations_df is not None else {}
    eq_vars = {
        "B": B, "L": L, "D": D, "Df": D, "t": t, "Cx": Cx, "Cy": Cy, "Ox": Ox, "Oy": Oy,
        "Qc": Qc, "Qc_t": Qc, "Qt": Qt, "Q": Qt, "Hx": Hx, "Hy": Hy, "Ht": Ht,
        "Mcx": Mcx, "Mcy": Mcy, "Mtx": Mtx, "Mty": Mty, "eB": abs(ex), "eL": abs(ey),
        "ex": ex, "ey": ey, "ex_moment": ex_moment, "ey_moment": ey_moment, "ex_direct": Ox, "ey_direct": Oy,
        "Bx": Bx, "Ly": Ly, "B_eff": B_eff, "L_eff": L_eff, "A_eff": A_eff,
        "q": q, "g1": g1used, "g2": g2used, "g1used": g1used, "g2used": g2used,
        "gamma1": g1used, "gamma2": g2used, "c1": c1, "c2": c2,
        "phi1_deg": phi1, "phi2_deg": phi2, "phi1": phi1r, "phi2": phi2r,
        "FS": FS, "D_over_B": D_over_B, "depth_argument": depth_argument,
        "beta_deg": beta_deg, "beta_from_load": beta_from_load,
        "footing_shape": shape, "shape_code": 1.0 if shape == "S" else 2.0, "is_strip": 1.0 if shape == "S" else 0.0,
        "Nq": Nq, "Nc": Nc, "Ny": Ny, "Ngamma": Ny,
        "Fcs": Fcs, "Fqs": Fqs, "Fys": Fys, "Fcd": Fcd, "Fqd": Fqd, "Fyd": Fyd,
        "Fci": Fci, "Fqi": Fqi, "Fyi": Fyi,
    }
    final_params = {"term_c", "term_q", "term_y", "qult", "qall", "Qult", "Qall", "qmax", "qmin"}
    for parameter, equation in shallow_eqs.items():
        if parameter not in final_params:
            eq_vars[parameter] = safe_eval(equation, eq_vars, eq_vars.get(parameter, 0.0))

    if shape == 'S':
        eq_vars["Fcs"] = 1.0
        eq_vars["Fqs"] = 1.0
        eq_vars["Fys"] = 1.0

    Nq = eq_vars.get("Nq", Nq)
    Nc = eq_vars.get("Nc", Nc)
    Ny = eq_vars.get("Ny", eq_vars.get("Ngamma", Ny))
    Fcs = eq_vars.get("Fcs", Fcs)
    Fqs = eq_vars.get("Fqs", Fqs)
    Fys = eq_vars.get("Fys", Fys)
    Fcd = eq_vars.get("Fcd", Fcd)
    Fqd = eq_vars.get("Fqd", Fqd)
    Fyd = eq_vars.get("Fyd", Fyd)
    Fci = eq_vars.get("Fci", Fci)
    Fqi = eq_vars.get("Fqi", Fqi)
    Fyi = eq_vars.get("Fyi", Fyi)
    B_eff = eq_vars.get("B_eff", B_eff)
    L_eff = eq_vars.get("L_eff", L_eff)
    A_eff = eq_vars.get("A_eff", B_eff * L_eff)

    if shape == 'S':
        Fcs = 1.0
        Fqs = 1.0
        Fys = 1.0
        eq_vars["Fcs"] = Fcs
        eq_vars["Fqs"] = Fqs
        eq_vars["Fys"] = Fys

    term_c = c2 * Nc * Fcs * Fcd * Fci
    term_q = q * Nq * Fqs * Fqd * Fqi
    term_y = 0.5 * B_eff * g2used * Ny * Fys * Fyd * Fyi
    qult = term_c + term_q + term_y
    qall = qult/FS
    Qult = qult*A_eff
    Qall = qall*A_eff
    qmax = Qt/(B*L) * (1 + 6*abs(ex)/B + 6*abs(ey)/L) if B*L>0 else 0.0
    qmin = Qt/(B*L) * (1 - 6*abs(ex)/B - 6*abs(ey)/L) if B*L>0 else 0.0

    eq_vars.update({
        "Nq": Nq, "Nc": Nc, "Ny": Ny, "Ngamma": Ny,
        "Fcs": Fcs, "Fqs": Fqs, "Fys": Fys, "Fcd": Fcd, "Fqd": Fqd, "Fyd": Fyd,
        "Fci": Fci, "Fqi": Fqi, "Fyi": Fyi,
        "B_eff": B_eff, "L_eff": L_eff, "A_eff": A_eff,
        "term_c": term_c, "term_q": term_q, "term_y": term_y,
        "qult": qult, "qall": qall, "Qult": Qult, "Qall": Qall, "qmax": qmax, "qmin": qmin,
    })
    for parameter in ["term_c", "term_q", "term_y", "qult", "qall", "Qult", "Qall", "qmax", "qmin"]:
        if parameter in shallow_eqs:
            eq_vars[parameter] = safe_eval(shallow_eqs[parameter], eq_vars, eq_vars.get(parameter, 0.0))
    term_c = eq_vars.get("term_c", term_c)
    term_q = eq_vars.get("term_q", term_q)
    term_y = eq_vars.get("term_y", term_y)
    qult = eq_vars.get("qult", qult)
    qall = eq_vars.get("qall", qult/FS)
    Qult = eq_vars.get("Qult", qult*A_eff)
    Qall = eq_vars.get("Qall", qall*A_eff)
    qmax = eq_vars.get("qmax", qmax)
    qmin = eq_vars.get("qmin", qmin)
    ru = unit_label(c.get('result_unit', result_unit))
    out = {
        'case_id': case_id, 'shape': shape, 'analysis_type': analysis, 'result_unit': ru,
        'B_m': B, 'L_m': L, 'D_m': D, 'water_table_depth_m': dw,
        'Qt': force_from_ton(Qt, ru), 'qult': stress_from_t_m2(qult, ru), 'qall': stress_from_t_m2(qall, ru),
        'Qult': force_from_ton(Qult, ru), 'Qall': force_from_ton(Qall, ru), 'qmax': stress_from_t_m2(qmax, ru), 'qmin': stress_from_t_m2(qmin, ru),
        'eB_m': abs(ex), 'eL_m': abs(ey), 'ex_moment_m': ex_moment, 'ey_moment_m': ey_moment, 'eccentricity_mode': eccentricity_mode, 'eccentricity_source': eccentricity_source,
        'Bx_m': Bx, 'Ly_m': Ly, 'B_eff_m': B_eff, 'L_eff_m': L_eff, 'A_eff_m2': A_eff,
        'B_used_for_factors_m': B_eff, 'D_over_B_eff': D_over_B, 'beta_deg': beta, 'beta_source': beta_source,
        'Nq': Nq, 'Nc': Nc, 'Ny': Ny, 'Fcs': Fcs, 'Fqs': Fqs, 'Fys': Fys, 'Fcd': Fcd, 'Fqd': Fqd, 'Fyd': Fyd,
        'Fci': Fci, 'Fqi': Fqi, 'Fyi': Fyi, 'FS_using_Qt': Qult/max(Qt,1e-6), 'FS_using_qmax': qult/max(qmax,1e-6),
        'criterion_Qall_ge_Qt': 'OK' if Qall >= Qt else 'NG', 'criterion_qall_ge_qmax': 'OK' if qall >= qmax else 'NG', 'criterion_qmin_ge_0': 'OK' if qmin >= 0 else 'NG',
        'criterion_t_ge_0_2': 'OK' if t >= 0.2 else 'NG', 'criterion_D_ge_1': 'OK' if D >= 1 else 'NG', 'foundation_class': 'Shallow' if D_over_B <= 1 else 'Deep'
    }
    detail = pd.DataFrame([
        {'factor':'c2','value':c2,'N':Nc,'shape_factor':Fcs,'depth_factor':Fcd,'inclination_factor':Fci,'B_used_m':B_eff,'beta_deg':beta,'beta_source':beta_source,'total_term_t_m2':term_c},
        {'factor':'q','value':q,'N':Nq,'shape_factor':Fqs,'depth_factor':Fqd,'inclination_factor':Fqi,'B_used_m':B_eff,'beta_deg':beta,'beta_source':beta_source,'total_term_t_m2':term_q},
        {'factor':"0.5B'g2",'value':0.5*B_eff*g2used,'N':Ny,'shape_factor':Fys,'depth_factor':Fyd,'inclination_factor':Fyi,'B_used_m':B_eff,'beta_deg':beta,'beta_source':beta_source,'total_term_t_m2':term_y},
    ])
    return out, detail


def calc_shallow_batch(df, result_unit='ton', equations_df=None):
    rows=[]
    details={}
    errs=[]
    for _, r in ensure_shallow_columns(df).iterrows():
        try:
            out, detail = calc_shallow_case(r, result_unit=result_unit, equations_df=equations_df)
            rows.append(out)
            details[str(out['case_id'])] = detail
        except Exception as e:
            errs.append({'case_id': r.get('case_id',''), 'error': str(e)})
    return pd.DataFrame(rows), details, pd.DataFrame(errs)


def shallow_equation_table():
    return pd.DataFrame([
        {'parameter':'Nq','equation':'exp(pi*tan(phi2))*tan(pi/4 + phi2/2)^2','latex':r'N_q=e^{\pi\tan\phi_2}\tan^2\left(\frac{\pi}{4}+\frac{\phi_2}{2}\right)','note':'Use phi2 for bearing factor'},
        {'parameter':'Nc','equation':'(Nq-1)/tan(phi2)','latex':r'N_c=(N_q-1)\cot\phi_2','note':'Use phi2'},
        {'parameter':'Ny','equation':'2*(Nq+1)*tan(phi2)','latex':r'N_\gamma=2(N_q+1)\tan\phi_2','note':'Use phi2'},
        {'parameter':'qult','equation':'c2*Nc*Fcs*Fcd*Fci + q*Nq*Fqs*Fqd*Fqi + 0.5*B_eff*g2*Ny*Fys*Fyd*Fyi','latex':r'q_{ult}=c_2N_cF_{cs}F_{cd}F_{ci}+qN_qF_{qs}F_{qd}F_{qi}+0.5B\'\gamma_2N_\gamma F_{\gamma s}F_{\gamma d}F_{\gamma i}','note':'Depth factor uses phi1; all other factors use phi2'},
        {'parameter':'equivalent dimensions','equation':'B_eff = B - 2*eB ; L_eff = L - 2*eL','latex':r'B\'=B-2e_B,\quad L\'=L-2e_L','note':'for load eccentricity'},
        {'parameter':'contact stress','equation':'qmax = Q/(BL)*(1+6eB/B+6eL/L); qmin = Q/(BL)*(1-6eB/B-6eL/L)','latex':r'q_{max}=\frac{Q}{BL}\left(1+6\frac{e_B}{B}+6\frac{e_L}{L}\right),\quad q_{min}=\frac{Q}{BL}\left(1-6\frac{e_B}{B}-6\frac{e_L}{L}\right)','note':'Check qmin >= 0'},
    ])


def plot_shallow_geometry(row, board_theme='Bright'):
    face = '#0b1220' if str(board_theme).lower().startswith('dark') else 'white'
    text = '#e5e7eb' if str(board_theme).lower().startswith('dark') else '#0f172a'
    grid = '#334155' if str(board_theme).lower().startswith('dark') else '#cbd5e1'
    B = to_float(row.get('B_m'),1.5)
    L=to_float(row.get('L_m'),2.0)
    D=to_float(row.get('D_m'),1.0)
    t=to_float(row.get('t_m'),0.2)
    Cx=to_float(row.get('Cx_m'),0.2)
    Cy=to_float(row.get('Cy_m'),0.2)
    Ox=to_float(row.get('Ox_m'),0.1)
    Oy=to_float(row.get('Oy_m'),0.15)
    shape = clean_text(row.get('footing_shape'), 'R').upper()
    if shape == 'S':
        L = 1.0
        Oy = 0.0
    Qc=to_float(row.get('Qc_t'),20.0)
    Mcx=to_float(row.get('Mcx_tm'),3.0)
    Mcy=to_float(row.get('Mcy_tm'),2.0)
    Hx=to_float(row.get('Hx_t'),1.5)
    Hy=to_float(row.get('Hy_t'),2.5)
    phi1=to_float(row.get('phi1_deg'),30.0)
    phi2=to_float(row.get('phi2_deg'),33.0)
    g1=to_float(row.get('g1u_t_m3'),1.9)
    g2=to_float(row.get('g2u_t_m3'),1.5)
    g1sat=to_float(row.get('g1sat_t_m3'),g1)
    g2sat=to_float(row.get('g2sat_t_m3'),g2)
    c1=to_float(row.get('c1_t_m2'),1.0)
    c2=to_float(row.get('c2_t_m2'),2.0)
    same_soil_color = (
        abs(g1 - g2) < 1e-9 and
        abs(g1sat - g2sat) < 1e-9 and
        abs(c1 - c2) < 1e-9 and
        abs(phi1 - phi2) < 1e-9
    )
    soil1_color = '#b08968'
    soil2_color = soil1_color if same_soil_color else '#f2cc8f'
    wt=to_float(row.get('water_table_depth_m'),999.0)
    fig, axs = plt.subplots(1,2, figsize=(12.5,5.2), facecolor=face, gridspec_kw={'width_ratios':[1.05,1.1]})
    for ax in axs:
        ax.set_facecolor(face)
        for sp in ax.spines.values():
            sp.set_color(grid)
        ax.tick_params(colors=text, labelsize=8)
    ax=axs[0]
    ax.set_xlim(-0.8, B+0.8)
    ax.set_ylim(D+1.2, -0.8)
    col_x = B/2 + Ox
    ax.fill_between([-0.8,B+0.8],0,D, color=soil1_color, alpha=0.88)
    ax.fill_between([-0.8,B+0.8],D,D+1.2, color=soil2_color, alpha=0.96)
    ax.add_patch(Rectangle((0,D-t),B,t, facecolor='#d1d5db', edgecolor=grid, lw=1.5))
    ax.add_patch(Rectangle((col_x-Cx/2,0),Cx,D-t, facecolor='#e5e7eb', edgecolor=grid, lw=1.5))
    ax.axhline(0, color=grid, lw=1.2)
    ax.axhline(D, color=grid, lw=1.0, linestyle='--')
    ax.axvline(B/2, color='#111827', lw=1.25, linestyle='-', alpha=0.80)
    if abs(Ox) > 1e-9:
        ax.annotate('', xy=(col_x, D*0.18), xytext=(B/2, D*0.18), arrowprops=dict(arrowstyle='<->', color='#7c3aed', lw=1.2))
        ax.text((col_x+B/2)/2 - 0.06*max(B,1.0), D*0.18-0.05, f'Ox={Ox:g} m', color='#7c3aed', fontsize=7, ha='right', va='top')
    if 0 <= wt <= D+1.2:
        ax.axhline(wt, color='#0284c7', lw=2)
        ax.scatter([0.05*B,0.18*B,0.31*B],[wt,wt,wt], marker='v', color='#0284c7', s=35)
    ax.annotate('', xy=(col_x, -0.1), xytext=(col_x, -0.65), arrowprops=dict(arrowstyle='simple', color='#2563eb'))
    ax.text(col_x+0.05, -0.3, f'Qc={Qc:g} t', color='#2563eb', fontsize=8, va='center')
    ax.annotate('', xy=(col_x,0.4), xytext=(col_x+0.45,0.4), arrowprops=dict(arrowstyle='->', color='#ef4444', lw=1.8))
    ax.text(col_x+0.48,0.4, f'Hx={Hx:g} t', color='#ef4444', fontsize=7.6, va='center')
    ax.annotate('', xy=(col_x,0.6), xytext=(col_x,0.1), arrowprops=dict(arrowstyle='->', color='#22c55e', lw=1.8))
    ax.text(col_x+0.04,0.08, f'Hy={Hy:g} t', color='#22c55e', fontsize=7.6, va='center')
    ax.text(B+0.08, D/2, 'Soil 1', color=text, fontsize=8)
    ax.text(B+0.08, D+0.55, 'Soil 2', color=text, fontsize=8)
    ax.text(B+0.08, D/2+0.18, fr'$\gamma_1={g1:g}, c_1={c1:g}, \phi_1={phi1:g}^\circ$', color=text, fontsize=7)
    ax.text(B+0.08, D+0.73, fr'$\gamma_2={g2:g}, c_2={c2:g}, \phi_2={phi2:g}^\circ$', color=text, fontsize=7)
    ax.annotate('', xy=(0,D+0.12), xytext=(B,D+0.12), arrowprops=dict(arrowstyle='<->', color=text, lw=1.2))
    ax.text(B/2, D+0.24, f'B={B:g} m', color=text, fontsize=8, ha='center', va='top')
    ax.annotate('', xy=(-0.18,0), xytext=(-0.18,D), arrowprops=dict(arrowstyle='<->', color=text, lw=1.2))
    ax.text(-0.2,D/2, f'D={D:g} m', color=text, fontsize=8, rotation=90, va='center', ha='right')
    ax.text(B+0.08, D-t/2, f't={t:g} m', color=text, fontsize=7.5)
    ax.set_title('Section / geometry', color=text, fontsize=10, fontweight='bold')
    ax.set_xticks([])
    ax.set_yticks([])

    ax=axs[1]
    ax.set_xlim(-0.3, B+0.3)
    ax.set_ylim(-0.3, L+0.3)
    ax.add_patch(Rectangle((0,0),B,L, facecolor='#ede9fe', edgecolor=grid, lw=1.6, alpha=0.95))
    if Cy > 0:
        ax.add_patch(Rectangle((B/2-Cx/2+Ox,L/2-Cy/2+Oy),Cx,Cy, facecolor='#d1d5db', edgecolor=grid, lw=1.2))
    else:
        ax.plot([B/2-Cx/2+Ox, B/2+Cx/2+Ox], [L/2, L/2], color='#d1d5db', lw=4, solid_capstyle='round')
    ax.scatter([B/2],[L/2], color='#8b5cf6', s=35)
    ax.annotate('', xy=(B/2+0.30,L/2), xytext=(B/2,L/2), arrowprops=dict(arrowstyle='->', color='#ef4444', lw=1.6))
    ax.text(B/2+0.32,L/2, f'Ox={Ox:g}', color='#ef4444', fontsize=7, va='center')
    ax.annotate('', xy=(B/2,L/2+0.35), xytext=(B/2,L/2), arrowprops=dict(arrowstyle='->', color='#22c55e', lw=1.6))
    ax.text(B/2+0.03,L/2+0.38, f'Oy={Oy:g}', color='#22c55e', fontsize=7, va='bottom')
    ax.text(B/2, L+0.08, f'L={L:g} m', color=text, fontsize=8, ha='center')
    ax.text(B+0.08, L/2, f'Mcx={Mcx:g} tm\nMcy={Mcy:g} tm', color=text, fontsize=7.5, va='center')
    ax.set_title('Plan / eccentricity', color=text, fontsize=10, fontweight='bold')
    ax.set_xticks([])
    ax.set_yticks([])
    fig.tight_layout()
    return fig

# -----------------------------
# Length-diameter design chart helpers
# -----------------------------
def parse_number_list(text, default_values=None):
    vals = []
    for part in str(text).replace(";", ",").split(","):
        part = part.strip()
        if part == "":
            continue
        try:
            vals.append(float(part))
        except Exception:
            pass
    return vals if vals else (default_values or [])


def make_float_range(start, stop, step):
    start = float(start); stop = float(stop); step = float(step)
    if step <= 0:
        step = 1.0
    if stop < start:
        start, stop = stop, start
    n = int(np.floor((stop - start) / step)) + 1
    vals = [round(start + i * step, 6) for i in range(n + 1) if start + i * step <= stop + step * 0.25]
    return vals


def case_with_design_size(case, diameter_m=None, length_m=None):
    c = dict(case)
    if diameter_m is not None:
        D = float(diameter_m)
        c["pile_size"] = D
        perimeter, area = shape_geometry(c.get("pile_section", "C"), D, to_float(c.get("inner_size", 0), 0))
        c["pile_perimeter"] = perimeter
        c["pile_cross_section_area"] = area
    if length_m is not None:
        c["pile_length_m"] = float(length_m)
    return c


def qall_for_design(layers, case, eqs, diameter_m, length_m, output_unit="ton"):
    c = case_with_design_size(case, diameter_m, length_m)
    table = calculate_excel_style_table(layers, c, eqs)
    if table.empty:
        return {"Qall": 0.0, "Qult": 0.0, "Qs": 0.0, "Qb": 0.0, "selected_depth_m": 0.0}
    r = result_at_pile_length(table, c)
    return {
        "Qall": force_from_ton(r.get("Qall (ton)", 0.0), output_unit),
        "Qult": force_from_ton(r.get("Qult (ton)", 0.0), output_unit),
        "Qs": force_from_ton(r.get("Qs cumulative (ton)", 0.0), output_unit),
        "Qb": force_from_ton(r.get("Qe (ton)", 0.0), output_unit),
        "selected_depth_m": r.get("Depth from (m)", 0.0),
    }


def capacity_vs_length_table(layers, case, eqs, diameters, lengths, output_unit="ton"):
    rows = []
    for L in lengths:
        row = {"Length (m)": float(L)}
        for D in diameters:
            key = f"D={D:g} m Qall ({cap_unit(output_unit)})"
            row[key] = qall_for_design(layers, case, eqs, D, L, output_unit)["Qall"]
        rows.append(row)
    return display_no_missing(pd.DataFrame(rows).round(4))


def capacity_vs_diameter_table(layers, case, eqs, diameters, selected_lengths, output_unit="ton"):
    rows = []
    for D in diameters:
        row = {"Diameter/size (m)": float(D)}
        for L in selected_lengths:
            key = f"L={L:g} m Qall ({cap_unit(output_unit)})"
            row[key] = qall_for_design(layers, case, eqs, D, L, output_unit)["Qall"]
        rows.append(row)
    return display_no_missing(pd.DataFrame(rows).round(4))


def capacity_heatmap_table(layers, case, eqs, diameters, lengths, output_unit="ton"):
    rows = []
    for L in lengths:
        for D in diameters:
            vals = qall_for_design(layers, case, eqs, D, L, output_unit)
            rows.append({
                "Length (m)": float(L),
                "Diameter/size (m)": float(D),
                f"Qall ({cap_unit(output_unit)})": vals["Qall"],
                f"Qult ({cap_unit(output_unit)})": vals["Qult"],
                f"Qs ({cap_unit(output_unit)})": vals["Qs"],
                f"Qb ({cap_unit(output_unit)})": vals["Qb"],
            })
    return display_no_missing(pd.DataFrame(rows).round(4))


def plot_capacity_vs_length(design_df, output_unit="ton", board_theme="Bright"):
    face = "#0b1220" if str(board_theme).lower().startswith("dark") else "white"
    text = "#e5e7eb" if str(board_theme).lower().startswith("dark") else "#0f172a"
    grid = "#334155" if str(board_theme).lower().startswith("dark") else "#cbd5e1"
    fig, ax = plt.subplots(figsize=(8.2, 5.4), facecolor=face)
    ax.set_facecolor(face)
    x = design_df["Length (m)"]
    for c in design_df.columns:
        if c != "Length (m)":
            ax.plot(x, design_df[c], linewidth=1.8, label=c.split(" Qall")[0])
    ax.set_xlabel("Pile length (m)", color=text)
    ax.set_ylabel(f"Allowable capacity, Qall ({cap_unit(output_unit)})", color=text)
    ax.set_title("Qall vs pile length", color=text, fontweight="bold")
    ax.grid(True, color=grid, alpha=.55)
    ax.tick_params(colors=text)
    for sp in ax.spines.values(): sp.set_color(grid)
    ax.legend(fontsize=8)
    fig.tight_layout()
    return fig


def plot_capacity_vs_diameter(design_df, output_unit="ton", board_theme="Bright"):
    face = "#0b1220" if str(board_theme).lower().startswith("dark") else "white"
    text = "#e5e7eb" if str(board_theme).lower().startswith("dark") else "#0f172a"
    grid = "#334155" if str(board_theme).lower().startswith("dark") else "#cbd5e1"
    fig, ax = plt.subplots(figsize=(8.2, 5.4), facecolor=face)
    ax.set_facecolor(face)
    x = design_df["Diameter/size (m)"]
    for c in design_df.columns:
        if c != "Diameter/size (m)":
            ax.plot(x, design_df[c], linewidth=1.8, label=c.split(" Qall")[0])
    ax.set_xlabel("Pile diameter/size (m)", color=text)
    ax.set_ylabel(f"Allowable capacity, Qall ({cap_unit(output_unit)})", color=text)
    ax.set_title("Qall vs pile diameter/size", color=text, fontweight="bold")
    ax.grid(True, color=grid, alpha=.55)
    ax.tick_params(colors=text)
    for sp in ax.spines.values(): sp.set_color(grid)
    ax.legend(fontsize=8)
    fig.tight_layout()
    return fig


def plot_capacity_heatmap(heat_df, output_unit="ton", board_theme="Bright"):
    face = "#0b1220" if str(board_theme).lower().startswith("dark") else "white"
    text = "#e5e7eb" if str(board_theme).lower().startswith("dark") else "#0f172a"
    value_col = f"Qall ({cap_unit(output_unit)})"
    pivot = heat_df.pivot(index="Length (m)", columns="Diameter/size (m)", values=value_col)
    fig, ax = plt.subplots(figsize=(8.2, 5.4), facecolor=face)
    ax.set_facecolor(face)
    im = ax.imshow(pivot.values, aspect="auto", origin="lower",
                   extent=[pivot.columns.min(), pivot.columns.max(), pivot.index.min(), pivot.index.max()])
    ax.set_xlabel("Pile diameter/size (m)", color=text)
    ax.set_ylabel("Pile length (m)", color=text)
    ax.set_title(f"Qall heatmap ({cap_unit(output_unit)})", color=text, fontweight="bold")
    ax.tick_params(colors=text)
    for sp in ax.spines.values(): sp.set_color("#94a3b8")
    cb = fig.colorbar(im, ax=ax)
    cb.set_label(f"Qall ({cap_unit(output_unit)})", color=text)
    cb.ax.tick_params(colors=text)
    fig.tight_layout()
    return fig


# -----------------------------
# PDF report helpers
# -----------------------------
def _pdf_table_from_df(df, max_rows=28, max_cols=10):
    d = display_no_missing(df).copy()
    if d.empty:
        return [["No data"]]
    d = d.iloc[:max_rows, :max_cols]
    data = [list(map(str, d.columns.tolist()))]
    for _, row in d.iterrows():
        vals = []
        for x in row.tolist():
            if isinstance(x, (float, np.floating)):
                vals.append(f"{x:.4g}")
            else:
                vals.append(str(x))
        data.append(vals)
    return data


def _fig_to_png_bytes(fig):
    bio = io.BytesIO()
    fig.savefig(bio, format="png", dpi=160, bbox_inches="tight")
    bio.seek(0)
    return bio.getvalue()



def _build_pdf_report(title, subtitle, sections):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), rightMargin=12*mm, leftMargin=12*mm, topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("AppTitle", parent=styles["Title"], fontSize=18, leading=22, textColor=colors.HexColor("#0f172a"))
    h_style = ParagraphStyle("Heading", parent=styles["Heading2"], fontSize=12, leading=15, textColor=colors.HexColor("#1d4ed8"))
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=8, leading=10)
    story = [Paragraph(title, title_style), Paragraph(subtitle, small), Spacer(1, 5*mm)]
    for sec in sections:
        story.append(Paragraph(sec.get("heading", "Section"), h_style))
        if sec.get("text"):
            story.append(Paragraph(sec["text"], small))
            story.append(Spacer(1, 2*mm))
        if sec.get("image_bytes") is not None:
            try:
                img = RLImage(io.BytesIO(sec["image_bytes"]))
                max_w = sec.get("image_width_mm", 250) * mm
                max_h = sec.get("image_height_mm", 135) * mm
                ratio = img.imageHeight / max(img.imageWidth, 1)
                img.drawWidth = max_w
                img.drawHeight = max_w * ratio
                if img.drawHeight > max_h:
                    img.drawHeight = max_h
                    img.drawWidth = max_h / max(ratio, 1e-9)
                story.append(img)
                story.append(Spacer(1, 4*mm))
            except Exception:
                pass
        if sec.get("df") is not None:
            table_data = _pdf_table_from_df(sec["df"], sec.get("max_rows", 28), sec.get("max_cols", 10))
            tbl = Table(table_data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#dbeafe")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0f172a")),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE", (0,0), (-1,-1), 6.5),
                ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 4*mm))
    doc.build(story)
    bio.seek(0)
    return bio.getvalue()


def selected_pile_pdf_report_bytes(pile_id, case, summary_row, calc_table, output_unit="ton", figure=None):
    summary_df = pd.DataFrame([summary_row]) if isinstance(summary_row, dict) else pd.DataFrame()
    subtitle = f"Selected pile report - {pile_id} | Unit: {cap_unit(output_unit)} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    sections = [
        {"heading":"Selected pile summary", "df":summary_df, "max_rows":5, "max_cols":12},
    ]
    if figure is not None:
        sections.append({"heading":"Diagram", "image_bytes": _fig_to_png_bytes(figure), "image_width_mm": 250, "image_height_mm": 135})
    pdf_calc_table = deep_pdf_table_view(calc_table)
    sections.append({"heading":"Excel-style calculation table", "df":pdf_calc_table, "max_rows":42, "max_cols":12})
    return _build_pdf_report("Selected Pile Capacity Report", subtitle, sections)


def selected_shallow_pdf_report_bytes(case_id, summary_row, factor_table, figure=None):
    summary_df = pd.DataFrame([summary_row]) if isinstance(summary_row, (dict, pd.Series)) else pd.DataFrame()
    subtitle = f"Selected shallow foundation report - {case_id} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    sections = [
        {"heading":"Selected case summary", "df":summary_df, "max_rows":5, "max_cols":14},
    ]
    if figure is not None:
        sections.append({"heading":"Diagram", "image_bytes": _fig_to_png_bytes(figure), "image_width_mm": 250, "image_height_mm": 135})
    if factor_table is not None and not factor_table.empty:
        sections.append({"heading":"Factor table", "df":factor_table, "max_rows":8, "max_cols":9})
    return _build_pdf_report("Selected Shallow Foundation Report", subtitle, sections)


def batch_summary_pdf_report_bytes(results_df, errors_df=None, run_name="pile_capacity_check"):
    subtitle = f"Batch summary report - {run_name} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    total = len(results_df) if results_df is not None else 0
    ok = int((results_df["load_status"] == "OK").sum()) if results_df is not None and not results_df.empty and "load_status" in results_df else 0
    ng = total - ok
    summary = pd.DataFrame([{"total_piles": total, "load_OK": ok, "load_NG": ng}])
    sections = [
        {"heading":"Summary", "df":summary, "max_rows":5, "max_cols":8},
        {"heading":"Batch result table", "df":results_df, "max_rows":36, "max_cols":12},
    ]
    if errors_df is not None and not errors_df.empty:
        sections.append({"heading":"Input/calculation errors", "df":errors_df, "max_rows":20, "max_cols":8})
    return _build_pdf_report("Batch Pile Capacity Summary", subtitle, sections)

def deep_pdf_table_view(calc_table):
    if calc_table is None or calc_table.empty:
        return pd.DataFrame()
    preferred = [
        "Depth from (m)", "Depth to (m)", "Soil type", "Soil name",
        "fs active", "Qs cumulative", "qe,ave", "Qe", "Qult", "Qall", "Qc,all", "Required load"
    ]
    chosen = []
    for pref in preferred:
        for col in calc_table.columns:
            if col == pref or str(col).startswith(pref):
                if col not in chosen:
                    chosen.append(col)
                break
    return calc_table[chosen] if chosen else calc_table


def excel_reference_comparison(app_table, reference_source=None, output_unit="ton", decimals=None):
    """Return summary comparison between app output and a reference table.

    reference_source can be a path or a DataFrame. App values are not rounded unless decimals is set.
    """
    if app_table is None or app_table.empty:
        return pd.DataFrame()
    if isinstance(reference_source, pd.DataFrame):
        ref = normalize_reference_columns(reference_source)
    elif reference_source is not None and Path(reference_source).exists():
        try:
            ref = normalize_reference_columns(pd.read_csv(reference_source))
        except Exception:
            return pd.DataFrame()
    else:
        return pd.DataFrame()
    if ref.empty:
        return pd.DataFrame()
    app = display_calculation_table(app_table, "ton", decimals=None)
    pairs = [
        ("Depth from (m)", "Depth from (m)"),
        ("Depth to (m)", "Depth to (m)"),
        ("fs active (t/m²)", "fs active (t/m²)"),
        ("qe (t/m²)", "qe (t/m²)"),
        ("qe,ave (t/m²)", "qe,ave (t/m²)"),
        ("Qe (ton)", "Qe (ton)"),
        ("Qult (ton)", "Qult (ton)"),
        ("Qall (ton)", "Qall (ton)"),
        ("Qc,all (ton)", "Qc,all (ton)"),
    ]
    if unit_label(output_unit) == "kN":
        pairs += [
            ("Qe (ton)", "Qe (kN)"),
            ("Qult (ton)", "Qult (kN)"),
            ("Qall (ton)", "Qall (kN)"),
            ("Qc,all (ton)", "Qc,all (kN)"),
        ]
    rows = []
    n = min(len(app), len(ref))
    for app_col, ref_col in pairs:
        if app_col in app.columns and ref_col in ref.columns:
            a = pd.to_numeric(app[app_col].iloc[:n], errors="coerce")
            r = pd.to_numeric(ref[ref_col].iloc[:n], errors="coerce")
            if unit_label(output_unit) == "kN" and ref_col.endswith("(kN)") and app_col.endswith("(ton)"):
                a = a * TON_TO_KN
            diff_signed = a - r
            diff = diff_signed.abs()
            rows.append({
                "column": app_col,
                "reference_column": ref_col,
                "rows_checked": int(diff.notna().sum()),
                "max_abs_difference": float(diff.max(skipna=True)) if diff.notna().any() else 0.0,
                "mean_abs_difference": float(diff.mean(skipna=True)) if diff.notna().any() else 0.0,
                "max_signed_difference": float(diff_signed.max(skipna=True)) if diff_signed.notna().any() else 0.0,
                "min_signed_difference": float(diff_signed.min(skipna=True)) if diff_signed.notna().any() else 0.0,
            })
    return round_for_display(pd.DataFrame(rows), decimals)

def excel_reference_detail_comparison(app_table, reference_source=None, output_unit="ton", decimals=None):
    """Return row-by-row comparison for exact checking."""
    if app_table is None or app_table.empty:
        return pd.DataFrame()
    if isinstance(reference_source, pd.DataFrame):
        ref = normalize_reference_columns(reference_source)
    elif reference_source is not None and Path(reference_source).exists():
        try:
            ref = normalize_reference_columns(pd.read_csv(reference_source))
        except Exception:
            return pd.DataFrame()
    else:
        return pd.DataFrame()
    if ref.empty:
        return pd.DataFrame()
    app = display_calculation_table(app_table, "ton", decimals=None)
    columns = ["fs active (t/m²)", "qe (t/m²)", "qe,ave (t/m²)", "Qe (ton)", "Qult (ton)", "Qall (ton)", "Qc,all (ton)"]
    n = min(len(app), len(ref))
    rows = []
    for i in range(n):
        base = {"row": i + 1}
        if "Depth from (m)" in app.columns:
            base["Depth from (m)"] = app["Depth from (m)"].iloc[i]
        if "Depth to (m)" in app.columns:
            base["Depth to (m)"] = app["Depth to (m)"].iloc[i]
        for c in columns:
            if c in app.columns and c in ref.columns:
                av = pd.to_numeric(pd.Series([app[c].iloc[i]]), errors="coerce").iloc[0]
                rv = pd.to_numeric(pd.Series([ref[c].iloc[i]]), errors="coerce").iloc[0]
                base[f"app {c}"] = av
                base[f"excel {c}"] = rv
                base[f"diff {c}"] = av - rv
        rows.append(base)
    return round_for_display(pd.DataFrame(rows), decimals)


def generic_reference_comparison(app_df, reference_source=None, decimals=None):
    """Compare common numeric columns between an app table and a user reference table."""
    if app_df is None or len(app_df) == 0:
        return pd.DataFrame(), pd.DataFrame()
    if isinstance(reference_source, pd.DataFrame):
        ref = normalize_reference_columns(reference_source)
    elif reference_source is not None and Path(reference_source).exists():
        try:
            ref = normalize_reference_columns(pd.read_csv(reference_source))
        except Exception:
            return pd.DataFrame(), pd.DataFrame()
    else:
        return pd.DataFrame(), pd.DataFrame()
    if ref.empty:
        return pd.DataFrame(), pd.DataFrame()
    app = display_no_missing(app_df).copy()
    common = [c for c in app.columns if c in ref.columns]
    rows = []
    detail_rows = []
    n = min(len(app), len(ref))
    for c in common:
        a = pd.to_numeric(app[c].iloc[:n], errors="coerce")
        r = pd.to_numeric(ref[c].iloc[:n], errors="coerce")
        valid = a.notna() & r.notna()
        if not valid.any():
            continue
        diff_signed = a[valid] - r[valid]
        diff_abs = diff_signed.abs()
        rows.append({
            "column": c,
            "rows_checked": int(valid.sum()),
            "max_abs_difference": float(diff_abs.max()),
            "mean_abs_difference": float(diff_abs.mean()),
            "max_signed_difference": float(diff_signed.max()),
            "min_signed_difference": float(diff_signed.min()),
        })
        for idx in diff_signed.index:
            detail_rows.append({
                "row": int(idx) + 1,
                "column": c,
                "app_value": float(a.loc[idx]),
                "reference_value": float(r.loc[idx]),
                "difference": float(a.loc[idx] - r.loc[idx]),
            })
    return round_for_display(pd.DataFrame(rows), decimals), round_for_display(pd.DataFrame(detail_rows), decimals)


def _deep_qall_from_case_layers(layers, case, eqs, output_unit="ton"):
    table = calculate_excel_style_table(layers, case, eqs)
    if table.empty:
        return 0.0
    row = result_at_pile_length(table, case)
    return force_from_ton(row.get("Qall (ton)", 0.0), output_unit)

def deep_sensitivity_table(layers, case, eqs, output_unit="ton"):
    """Feature-importance style one-at-a-time sensitivity ranking."""
    base_case = dict(case)
    base_layers = layers.copy()
    base_qall = _deep_qall_from_case_layers(base_layers, base_case, eqs, output_unit)

    scenarios = []
    base_size = to_float(base_case.get("pile_size"), 0.6)
    base_length = to_float(base_case.get("pile_length_m"), 20.0)
    base_wt = to_float(base_case.get("water_table_level"), 1.0)

    scenarios.append(("Pile length", base_length, base_length * 0.90, base_length * 1.10, "m",
                      lambda v: (dict(base_case, pile_length_m=v), base_layers.copy())))
    scenarios.append(("Pile size", base_size, base_size * 0.90, base_size * 1.10, "m",
                      lambda v: (case_with_design_size(base_case, diameter_m=v, length_m=base_length), base_layers.copy())))
    scenarios.append(("Water table", base_wt, max(base_wt - 1.0, 0.0), base_wt + 1.0, "m",
                      lambda v: (dict(base_case, water_table_level=v), base_layers.copy())))

    if "spt_N" in base_layers.columns:
        scenarios.append(("SPT N", 1.0, 0.90, 1.10, "multiplier",
                          lambda v: (base_case.copy(), base_layers.assign(spt_N=pd.to_numeric(base_layers["spt_N"], errors="coerce") * v))))
    if "su1_t_m2" in base_layers.columns:
        scenarios.append(("Su", 1.0, 0.90, 1.10, "multiplier",
                          lambda v: (base_case.copy(), base_layers.assign(su1_t_m2=pd.to_numeric(base_layers["su1_t_m2"], errors="coerce") * v))))
    if "gamma_t_m3" in base_layers.columns:
        scenarios.append(("Unit weight γ", 1.0, 0.95, 1.05, "multiplier",
                          lambda v: (base_case.copy(), base_layers.assign(gamma_t_m3=pd.to_numeric(base_layers["gamma_t_m3"], errors="coerce") * v))))

    rows = []
    for parameter, base_value, low_value, high_value, unit, builder in scenarios:
        low_case, low_layers = builder(low_value)
        high_case, high_layers = builder(high_value)
        q_low = _deep_qall_from_case_layers(low_layers, low_case, eqs, output_unit)
        q_high = _deep_qall_from_case_layers(high_layers, high_case, eqs, output_unit)
        influence_range = abs(q_high - q_low)
        rows.append({
            "parameter": parameter,
            "base_value": base_value,
            "low_value": low_value,
            "high_value": high_value,
            "unit": unit,
            f"Qall_low ({cap_unit(output_unit)})": q_low,
            f"Qall_high ({cap_unit(output_unit)})": q_high,
            "sensitivity_range": influence_range,
            "change_from_base_low_%": ((q_low - base_qall) / base_qall * 100.0) if base_qall else 0.0,
            "change_from_base_high_%": ((q_high - base_qall) / base_qall * 100.0) if base_qall else 0.0,
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    total = out["sensitivity_range"].sum()
    out["relative_influence_%"] = out["sensitivity_range"] / total * 100.0 if total else 0.0
    out = out.sort_values("relative_influence_%", ascending=False).reset_index(drop=True)
    out.insert(0, "rank", range(1, len(out) + 1))
    return out

def shallow_sensitivity_table(row, result_unit="ton", equations_df=None):
    """Feature-importance style one-at-a-time sensitivity ranking."""
    base_row = dict(row)
    base_out, _ = calc_shallow_case(base_row, result_unit=result_unit, equations_df=equations_df)
    base_qall = to_float(base_out.get("Qall", 0.0), 0.0)

    B0 = to_float(base_row.get("B_m"), 1.0)
    D0 = to_float(base_row.get("D_m"), 1.0)
    Ox0 = to_float(base_row.get("Ox_m"), 0.0)
    c20 = to_float(base_row.get("c2_t_m2"), 1.0)
    phi20 = to_float(base_row.get("phi2_deg"), 30.0)
    g20 = to_float(base_row.get("g2u_t_m3"), 1.5)
    dw0 = to_float(base_row.get("water_table_depth_m"), 1.0)

    scenarios = [
        ("B", B0, B0 * 0.90, B0 * 1.10, "m", "B_m"),
        ("D", D0, D0 * 0.90, D0 * 1.10, "m", "D_m"),
        ("Ox", Ox0, Ox0 - 0.05, Ox0 + 0.05, "m", "Ox_m"),
        ("c2", c20, c20 * 0.90, c20 * 1.10, "t/m²", "c2_t_m2"),
        ("φ2", phi20, max(phi20 - 2.0, 0.0), phi20 + 2.0, "degree", "phi2_deg"),
        ("γ2", g20, g20 * 0.95, g20 * 1.05, "t/m³", "g2u_t_m3"),
        ("Water table", dw0, max(dw0 - 1.0, 0.0), dw0 + 1.0, "m", "water_table_depth_m"),
    ]

    rows = []
    for parameter, base_value, low_value, high_value, unit, field in scenarios:
        low_row = dict(base_row); low_row[field] = low_value
        high_row = dict(base_row); high_row[field] = high_value
        low_out, _ = calc_shallow_case(low_row, result_unit=result_unit, equations_df=equations_df)
        high_out, _ = calc_shallow_case(high_row, result_unit=result_unit, equations_df=equations_df)
        q_low = to_float(low_out.get("Qall", 0.0), 0.0)
        q_high = to_float(high_out.get("Qall", 0.0), 0.0)
        influence_range = abs(q_high - q_low)
        rows.append({
            "parameter": parameter,
            "base_value": base_value,
            "low_value": low_value,
            "high_value": high_value,
            "unit": unit,
            f"Qall_low ({cap_unit(result_unit)})": q_low,
            f"Qall_high ({cap_unit(result_unit)})": q_high,
            "sensitivity_range": influence_range,
            "change_from_base_low_%": ((q_low - base_qall) / base_qall * 100.0) if base_qall else 0.0,
            "change_from_base_high_%": ((q_high - base_qall) / base_qall * 100.0) if base_qall else 0.0,
        })
    out = pd.DataFrame(rows)
    total = out["sensitivity_range"].sum()
    out["relative_influence_%"] = out["sensitivity_range"] / total * 100.0 if total else 0.0
    out = out.sort_values("relative_influence_%", ascending=False).reset_index(drop=True)
    out.insert(0, "rank", range(1, len(out) + 1))
    return out

def shallow_design_vs_B_table(row, B_values, D_values, result_unit="ton", equations_df=None):
    rows = []
    for B in B_values:
        out = {"B (m)": float(B)}
        for D in D_values:
            rr = dict(row)
            rr["B_m"] = float(B)
            rr["D_m"] = float(D)
            calc, _ = calc_shallow_case(rr, result_unit=result_unit, equations_df=equations_df)
            out[f"D={float(D):g} m Qall ({cap_unit(result_unit)})"] = to_float(calc.get("Qall"), 0.0)
        rows.append(out)
    return pd.DataFrame(rows)

def shallow_design_vs_D_table(row, D_values, B_values, result_unit="ton", equations_df=None):
    rows = []
    for D in D_values:
        out = {"D (m)": float(D)}
        for B in B_values:
            rr = dict(row)
            rr["B_m"] = float(B)
            rr["D_m"] = float(D)
            calc, _ = calc_shallow_case(rr, result_unit=result_unit, equations_df=equations_df)
            out[f"B={float(B):g} m Qall ({cap_unit(result_unit)})"] = to_float(calc.get("Qall"), 0.0)
        rows.append(out)
    return pd.DataFrame(rows)

def plot_shallow_design_table(design_df, x_col, title, result_unit="ton", board_theme="Bright"):
    face = "#0b1220" if str(board_theme).lower().startswith("dark") else "white"
    text = "#e5e7eb" if str(board_theme).lower().startswith("dark") else "#0f172a"
    grid = "#334155" if str(board_theme).lower().startswith("dark") else "#cbd5e1"
    fig, ax = plt.subplots(figsize=(8.2, 5.4), facecolor=face)
    ax.set_facecolor(face)
    x = design_df[x_col]
    for c in design_df.columns:
        if c != x_col:
            ax.plot(x, design_df[c], linewidth=1.8, label=c.split(" Qall")[0])
    ax.set_xlabel(x_col, color=text)
    ax.set_ylabel(f"Allowable load, Qall ({cap_unit(result_unit)})", color=text)
    ax.set_title(title, color=text, fontweight="bold")
    ax.grid(True, color=grid, alpha=.55)
    ax.tick_params(colors=text)
    for sp in ax.spines.values():
        sp.set_color(grid)
    ax.legend(fontsize=8)
    fig.tight_layout()
    return fig

# -----------------------------
# File/history helpers
# -----------------------------
def df_to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8")

def df_to_excel_bytes(sheets):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            display_no_missing(df).to_excel(writer, sheet_name=str(name)[:31], index=False)
    bio.seek(0)
    return bio.getvalue()

def make_zip_bytes(file_paths):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in file_paths:
            p = Path(path)
            if p.exists(): zf.write(p, arcname=p.name)
    bio.seek(0)
    return bio.getvalue()

def save_history(run_name, pile_df, layers_df, eq_df, results_df, selected_table=None):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = "".join(ch if ch.isalnum() or ch in "_-" else "_" for ch in str(run_name))[:45]
    folder = HISTORY_DIR / f"{stamp}_{safe}"
    folder.mkdir(exist_ok=True)
    pile_df.to_csv(folder/"input_pile_cases.csv", index=False)
    layers_df.to_csv(folder/"input_borehole_layers.csv", index=False)
    eq_df.to_csv(folder/"input_equations.csv", index=False)
    results_df.to_csv(folder/"batch_results.csv", index=False)
    if selected_table is not None and not selected_table.empty:
        selected_table.to_csv(folder/"selected_pile_excel_style_calculation.csv", index=False)
    return folder

def list_history():
    return sorted([p for p in HISTORY_DIR.iterdir() if p.is_dir()], reverse=True)



# -----------------------------
# Permanent tutorial storage helpers
# -----------------------------
def _tutorial_secret(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default

def tutorial_storage_config():
    return {
        "token": _tutorial_secret("github_token", ""),
        "repo": _tutorial_secret("github_repo", ""),
        "branch": _tutorial_secret("github_branch", "main"),
        "path": _tutorial_secret("tutorial_links_path", "tutorial_links.csv"),
    }

def _tutorial_empty_df():
    return pd.DataFrame(columns=["title", "url", "description"])

def _clean_tutorial_df(df):
    if df is None or df.empty:
        return _tutorial_empty_df()
    out = df.copy()
    for col in ["title", "url", "description"]:
        if col not in out.columns:
            out[col] = ""
    return out[["title", "url", "description"]].fillna("")

def load_tutorial_links(tutorial_file):
    cfg = tutorial_storage_config()
    if cfg["token"] and cfg["repo"]:
        try:
            api = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}?ref={cfg['branch']}"
            req = urllib.request.Request(api, headers={
                "Authorization": f"Bearer {cfg['token']}",
                "Accept": "application/vnd.github+json",
                "User-Agent": "GeoCapacity"
            })
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            content = base64.b64decode(data.get("content", "")).decode("utf-8")
            return _clean_tutorial_df(pd.read_csv(io.StringIO(content))), "github"
        except Exception:
            pass
    if tutorial_file.exists():
        try:
            return _clean_tutorial_df(pd.read_csv(tutorial_file)), "local"
        except Exception:
            return _tutorial_empty_df(), "local"
    return _tutorial_empty_df(), "local"

def save_tutorial_links(df, tutorial_file):
    df = _clean_tutorial_df(df)
    csv_text = df.to_csv(index=False)
    cfg = tutorial_storage_config()
    if cfg["token"] and cfg["repo"]:
        api = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
        sha = None
        try:
            get_req = urllib.request.Request(f"{api}?ref={cfg['branch']}", headers={
                "Authorization": f"Bearer {cfg['token']}",
                "Accept": "application/vnd.github+json",
                "User-Agent": "GeoCapacity"
            })
            with urllib.request.urlopen(get_req, timeout=10) as resp:
                sha = json.loads(resp.read().decode("utf-8")).get("sha")
        except Exception:
            sha = None
        payload = {
            "message": "Update GeoCapacity tutorial links",
            "content": base64.b64encode(csv_text.encode("utf-8")).decode("utf-8"),
            "branch": cfg["branch"],
        }
        if sha:
            payload["sha"] = sha
        put_req = urllib.request.Request(api, data=json.dumps(payload).encode("utf-8"), method="PUT", headers={
            "Authorization": f"Bearer {cfg['token']}",
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
            "User-Agent": "GeoCapacity"
        })
        with urllib.request.urlopen(put_req, timeout=15) as resp:
            resp.read()
        tutorial_file.write_text(csv_text, encoding="utf-8")
        return "github"
    tutorial_file.write_text(csv_text, encoding="utf-8")
    return "local"


# -----------------------------
# User-assistance / advisor-review helpers
# -----------------------------
def _status_icon(ok):
    return "✅ OK" if ok else "⚠️ Check"

def deep_input_checklist_df(pile_df, layers_df):
    rows = []
    pile_required = ["pile_id", "boring_no", "pile_type", "pile_section", "pile_length_m"]
    layer_required = ["boring_no", "top_depth_m", "bottom_depth_m", "gamma_t_m3", "soil_type"]
    rows.append({
        "status": _status_icon(not pile_df.empty),
        "check": "Pile cases file loaded",
        "details": f"{len(pile_df)} pile case row(s)" if not pile_df.empty else "No pile cases loaded",
        "suggested action": "Upload pile cases CSV or use the bundled sample"
    })
    rows.append({
        "status": _status_icon(not layers_df.empty),
        "check": "Borehole layers file loaded",
        "details": f"{len(layers_df)} borehole layer row(s)" if not layers_df.empty else "No borehole layers loaded",
        "suggested action": "Upload borehole layers CSV or use the bundled sample"
    })
    if not pile_df.empty:
        missing = [c for c in pile_required if c not in pile_df.columns]
        rows.append({
            "status": _status_icon(len(missing) == 0),
            "check": "Required pile columns",
            "details": "All required pile columns found" if not missing else "Missing: " + ", ".join(missing),
            "suggested action": "Use the pile template if columns are missing"
        })
        if "pile_length_m" in pile_df.columns:
            bad = int((pd.to_numeric(pile_df["pile_length_m"], errors="coerce").fillna(0) <= 0).sum())
            rows.append({
                "status": _status_icon(bad == 0),
                "check": "Pile length values",
                "details": "All pile lengths are positive" if bad == 0 else f"{bad} row(s) have zero/blank/negative pile length",
                "suggested action": "Check pile_length_m"
            })
        if "pile_type" in pile_df.columns:
            bad = [str(v) for v in pile_df["pile_type"] if normalize_pile_type_text(v) == ""]
            rows.append({
                "status": _status_icon(len(bad) == 0),
                "check": "Pile type support",
                "details": "All pile types are supported" if not bad else "Unsupported examples: " + ", ".join(bad[:5]),
                "suggested action": "Use B/Bored or D/Driven"
            })
    if not layers_df.empty:
        missing = [c for c in layer_required if c not in layers_df.columns]
        rows.append({
            "status": _status_icon(len(missing) == 0),
            "check": "Required borehole layer columns",
            "details": "All required layer columns found" if not missing else "Missing: " + ", ".join(missing),
            "suggested action": "Use the borehole layer template if columns are missing"
        })
        if {"top_depth_m", "bottom_depth_m"}.issubset(layers_df.columns):
            top = pd.to_numeric(layers_df["top_depth_m"], errors="coerce")
            bottom = pd.to_numeric(layers_df["bottom_depth_m"], errors="coerce")
            bad = int(((bottom <= top) | top.isna() | bottom.isna()).sum())
            rows.append({
                "status": _status_icon(bad == 0),
                "check": "Layer depth order",
                "details": "All layer bottoms are deeper than layer tops" if bad == 0 else f"{bad} layer row(s) have invalid depth order",
                "suggested action": "Check top_depth_m and bottom_depth_m"
            })
        if "gamma_t_m3" in layers_df.columns:
            gamma = pd.to_numeric(layers_df["gamma_t_m3"], errors="coerce")
            bad = int(((gamma <= 0) | gamma.isna()).sum())
            rows.append({
                "status": _status_icon(bad == 0),
                "check": "Layer unit weight",
                "details": "All unit weights are positive" if bad == 0 else f"{bad} layer row(s) have blank/non-positive unit weight",
                "suggested action": "Check gamma_t_m3"
            })
    if not pile_df.empty and not layers_df.empty and {"boring_no"}.issubset(pile_df.columns) and {"boring_no"}.issubset(layers_df.columns):
        missing_bh = sorted(set(pile_df["boring_no"].astype(str)) - set(layers_df["boring_no"].astype(str)))
        rows.append({
            "status": _status_icon(len(missing_bh) == 0),
            "check": "Borehole matching",
            "details": "All pile cases have matching borehole layers" if not missing_bh else "Missing layers for: " + ", ".join(missing_bh[:8]),
            "suggested action": "Make boring_no identical in pile and layer CSVs"
        })
    return pd.DataFrame(rows)

def shallow_input_checklist_df(shallow_df):
    rows = []
    req = required_shallow_columns()
    rows.append({
        "status": _status_icon(shallow_df is not None and not shallow_df.empty),
        "check": "Shallow cases file loaded",
        "details": f"{len(shallow_df)} footing case row(s)" if shallow_df is not None and not shallow_df.empty else "No shallow cases loaded",
        "suggested action": "Upload footing cases CSV or use the bundled sample"
    })
    if shallow_df is not None and not shallow_df.empty:
        optional_cols = ['beta_deg']
        missing = [c for c in req if c not in shallow_df.columns and c not in optional_cols]
        rows.append({
            "status": _status_icon(len(missing) == 0),
            "check": "Required shallow columns",
            "details": "All shallow columns found" if not missing else "Missing: " + ", ".join(missing),
            "suggested action": "Use the footing template if columns are missing"
        })
        for col, label in [("B_m", "Footing width B"), ("D_m", "Foundation depth D"), ("FS", "Factor of safety")]:
            if col in shallow_df.columns:
                vals = pd.to_numeric(shallow_df[col], errors="coerce")
                bad = int(((vals <= 0) | vals.isna()).sum())
                rows.append({
                    "status": _status_icon(bad == 0),
                    "check": label,
                    "details": "Values are positive" if bad == 0 else f"{bad} row(s) have blank/non-positive values",
                    "suggested action": f"Check {col}"
                })
        if "footing_shape" in shallow_df.columns:
            shapes = shallow_df["footing_shape"].astype(str).str.upper().str.strip()
            bad = int((~shapes.isin(["S", "R"])).sum())
            rows.append({
                "status": _status_icon(bad == 0),
                "check": "Footing shape",
                "details": "All shapes are supported" if bad == 0 else f"{bad} row(s) are not S or R",
                "suggested action": "Use S for strip or R for rectangular"
            })
        if "result_unit" in shallow_df.columns:
            units = shallow_df["result_unit"].astype(str).str.lower().str.strip()
            bad = int((~units.isin(["", "t", "ton", "tons", "tonne", "tonnes", "kn"])).sum())
            rows.append({
                "status": _status_icon(bad == 0),
                "check": "Result unit",
                "details": "All result units are readable" if bad == 0 else f"{bad} row(s) have unusual result_unit",
                "suggested action": "Use ton or kN"
            })
    return pd.DataFrame(rows)

def deep_result_explanation_df(summary_row, table_ton, case):
    if table_ton is None or table_ton.empty:
        return pd.DataFrame()
    L = to_float(case.get("pile_length_m"), 0.0)
    idx = (table_ton["Depth from (m)"] - L).abs().idxmin()
    row = table_ton.loc[idx]
    qs = to_float(row.get("Qs cumulative (ton)", 0.0), 0.0)
    qe = to_float(row.get("Qe (ton)", 0.0), 0.0)
    qult = to_float(row.get("Qult (ton)", 0.0), 0.0)
    qall = to_float(row.get("Qall (ton)", 0.0), 0.0)
    req = to_float(row.get("Required load (ton)", 0.0), 0.0)
    return pd.DataFrame([
        {"item": "Selected depth", "value": f"{row.get('Depth from (m)', 0):.3g} m", "meaning": "Nearest calculated row to the selected pile length"},
        {"item": "Controlling soil layer", "value": f"{row.get('Soil name', '')} ({row.get('Soil type', '')})", "meaning": "Soil layer at the selected result depth"},
        {"item": "Shaft resistance share", "value": f"{(qs/qult*100 if qult else 0):.1f}%", "meaning": "Contribution of Qs to Qult"},
        {"item": "End bearing share", "value": f"{(qe/qult*100 if qult else 0):.1f}%", "meaning": "Contribution of Qe to Qult"},
        {"item": "Allowable capacity", "value": f"{qall:.3g} ton", "meaning": "Qult divided by factor of safety"},
        {"item": "Required load check", "value": "OK" if req > 0 and qall >= req else ("No required load" if req <= 0 else "NG"), "meaning": "Compares Qall with required load"},
        {"item": "Concrete check", "value": str(row.get("Concrete status", "")), "meaning": "Compares Qall with concrete allowable capacity"},
    ])

def shallow_result_explanation_df(outrow, factor_table):
    if outrow is None:
        return pd.DataFrame()
    out = dict(outrow)
    q_all = to_float(out.get("Qall", 0.0), 0.0)
    qt = to_float(out.get("Qt", 0.0), 0.0)
    qmax = to_float(out.get("qmax", 0.0), 0.0)
    qall = to_float(out.get("qall", 0.0), 0.0)
    rows = [
        {"item": "Foundation class", "value": clean_text(out.get("foundation_class", "")), "meaning": "Class based on D/B relationship"},
        {"item": "Effective width B'", "value": f"{to_float(out.get('B_eff_m', 0), 0):.3g} m", "meaning": "Width after eccentricity correction"},
        {"item": "Effective area A'", "value": f"{to_float(out.get('A_eff_m2', 0), 0):.3g} m²", "meaning": "Area used for capacity calculation"},
        {"item": "Allowable load check", "value": clean_text(out.get("criterion_Qall_ge_Qt", "")), "meaning": f"Qall = {q_all:.3g}, Qt = {qt:.3g}"},
        {"item": "Bearing pressure check", "value": clean_text(out.get("criterion_qall_ge_qmax", "")), "meaning": f"qall = {qall:.3g}, qmax = {qmax:.3g}"},
        {"item": "Contact pressure condition", "value": clean_text(out.get("criterion_qmin_ge_0", "")), "meaning": "Checks whether qmin is non-negative"},
    ]
    if factor_table is not None and not factor_table.empty and "total_term_t_m2" in factor_table.columns:
        ft = factor_table.copy()
        total = pd.to_numeric(ft["total_term_t_m2"], errors="coerce").fillna(0).sum()
        if total != 0:
            major = ft.assign(abs_term=pd.to_numeric(ft["total_term_t_m2"], errors="coerce").fillna(0).abs()).sort_values("abs_term", ascending=False).iloc[0]
            rows.append({"item": "Largest capacity term", "value": str(major.get("factor", "")), "meaning": f"Approx. {major.get('total_term_t_m2', 0):.3g} t/m²"})
    return pd.DataFrame(rows)

def selected_project_zip_bytes(file_map):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in file_map.items():
            if data is None:
                continue
            if isinstance(data, str):
                data = data.encode("utf-8")
            zf.writestr(name, data)
    bio.seek(0)
    return bio.getvalue()


# -----------------------------
# Admin-uploaded manual PDF helpers
# -----------------------------
def manual_pdf_path():
    return APP_DIR / "admin_manual.pdf"

def save_uploaded_manual_pdf(uploaded_file):
    if uploaded_file is None:
        return None
    data = uploaded_file.getvalue()
    path = manual_pdf_path()
    path.write_bytes(data)
    return path

def manual_pdf_download_bytes():
    path = manual_pdf_path()
    if path.exists():
        return path.read_bytes()
    return None



# -----------------------------
# Admin file and preset management helpers
# -----------------------------
def safe_storage_name(name, default="file"):
    base = clean_text(name, default)
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("._-")
    return base or default

def manual_registry_path():
    return MANUAL_UPLOAD_DIR / "manual_registry.csv"

def load_manual_registry():
    p = manual_registry_path()
    cols = ["title", "file_name", "description", "uploaded_at"]
    if p.exists():
        try:
            df = pd.read_csv(p)
        except Exception:
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols].fillna("")

def save_manual_registry(df):
    cols = ["title", "file_name", "description", "uploaded_at"]
    out = df.copy() if df is not None else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    out = out[cols].fillna("")
    MANUAL_UPLOAD_DIR.mkdir(exist_ok=True)
    out.to_csv(manual_registry_path(), index=False)
    return out

def manual_file_bytes(file_name):
    fn = safe_storage_name(file_name, "")
    p = MANUAL_UPLOAD_DIR / fn
    if p.exists():
        return p.read_bytes()
    return None

def save_manual_upload(uploaded_file, title="", description=""):
    if uploaded_file is None:
        return None
    original = safe_storage_name(uploaded_file.name, "manual.pdf")
    if not original.lower().endswith(".pdf"):
        original += ".pdf"
    stamped = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{original}"
    MANUAL_UPLOAD_DIR.mkdir(exist_ok=True)
    target = MANUAL_UPLOAD_DIR / stamped
    target.write_bytes(uploaded_file.getvalue())

    reg = load_manual_registry()
    row = pd.DataFrame([{
        "title": clean_text(title, Path(original).stem),
        "file_name": stamped,
        "description": clean_text(description, ""),
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }])
    save_manual_registry(pd.concat([reg, row], ignore_index=True))
    return target

def equation_builtin_preset_options():
    return {
        "Bangkok Excel Method": get_default_file("equation_library_bangkok_excel.csv"),
    }

def equation_custom_preset_options():
    EQUATION_PRESET_DIR.mkdir(exist_ok=True)
    opts = {}
    for p in sorted(EQUATION_PRESET_DIR.glob("*.csv")):
        label = "Custom: " + p.stem.replace("_", " ")
        opts[label] = p
    return opts

def equation_all_preset_options():
    opts = {}
    for name, path in equation_builtin_preset_options().items():
        opts[name] = {"path": path, "type": "built-in"}
    for name, path in equation_custom_preset_options().items():
        opts[name] = {"path": path, "type": "custom"}
    return opts

def embedded_bangkok_equations_df():
    return ensure_equation_columns(pd.DataFrame([{'parameter': 'su2_t_m2', 'equation': 'nan if SPT_N_missing > 0.5 else 0.52 * SPT_N', 'latex': 's_{u2}=0.52N\\quad(\\mathrm{blank\\ if\\ SPT\\ is\\ blank})', 'unit': 't/m²', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Excel match: blank SPT gives blank su2, not zero.'}, {'parameter': 'Su_t_m2', 'equation': 'avg(su1_t_m2, su2_t_m2)', 'latex': 'S_u=\\mathrm{avg}(s_{u1},s_{u2})', 'unit': 't/m²', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Average ignores blank/nan'}, {'parameter': 'alpha_clay', 'equation': '0.93 * exp(-0.0536 * Su_t_m2)', 'latex': '\\alpha=0.93e^{-0.0536S_u}', 'unit': '-', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Adhesion factor'}, {'parameter': 'CN', 'equation': '0.77 * log10(20 * 95.76 / max(sigma_v02_t_m2 * g_factor, 0.0001))', 'latex': "C_N=0.77\\log_{10}\\left(\\frac{20\\times95.76}{\\sigma'_{v,mid}g}\\right)", 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Overburden correction'}, {'parameter': 'N_corr', 'equation': 'CN * SPT_N', 'latex': "N'=C_NN", 'unit': 'blows/ft', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Corrected SPT'}, {'parameter': 'phi_deg', 'equation': '27.1 + 0.3 * N_corr - 0.00054 * N_corr**2', 'latex': "\\phi=27.1+0.3N'-0.00054N'^2", 'unit': 'degree', 'soil_type': 'Sand', 'source': 'Das/Wolff style', 'notes': 'Friction angle correlation'}, {'parameter': 'K0', 'equation': '1 - sin(radians(phi_deg))', 'latex': 'K_0=1-\\sin\\phi', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'At-rest coefficient'}, {'parameter': 'Ks_driven', 'equation': '1.5 * K0', 'latex': 'K_{s,D}=1.5K_0', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Driven pile'}, {'parameter': 'Ks_bored', 'equation': '0.85 * K0', 'latex': 'K_{s,B}=0.85K_0', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Bored pile'}, {'parameter': 'Ks', 'equation': 'Ks_driven if pile_type_code > 1.5 else Ks_bored', 'latex': 'K_s=K_{s,D}\\ \\mathrm{or}\\ K_{s,B}', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Bored=1, Driven=2'}, {'parameter': 'delta_deg', 'equation': '0.75 * phi_deg', 'latex': '\\delta=0.75\\phi', 'unit': 'degree', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Interface friction angle'}, {'parameter': 'beta_sand', 'equation': 'Ks * tan(radians(delta_deg))', 'latex': '\\beta=K_s\\tan\\delta', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Beta method'}, {'parameter': 'fs_clay_t_m2', 'equation': 'alpha_clay * Su_t_m2', 'latex': 'f_s=\\alpha S_u', 'unit': 't/m²', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Clay unit shaft friction'}, {'parameter': 'fs_sand_t_m2', 'equation': 'beta_sand * sigma_v02_t_m2', 'latex': "f_s=\\beta\\sigma'_{v,mid}", 'unit': 't/m²', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Sand unit shaft friction'}, {'parameter': 'Nc', 'equation': '9', 'latex': 'N_c=9', 'unit': '-', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Clay bearing factor'}, {'parameter': 'Nq_driven', 'equation': 'exp(pi * tan(radians(phi_deg))) * tan(radians(45 + phi_deg/2))**2 * (1 + tan(radians(phi_deg)))', 'latex': 'N_{q,D}=e^{\\pi\\tan\\phi}\\tan^2(45^\\circ+\\phi/2)(1+\\tan\\phi)', 'unit': '-', 'soil_type': 'Sand', 'source': 'Advisor note', 'notes': 'Driven Nq; phi converted to radians in trig functions'}, {'parameter': 'Nq_bored', 'equation': '0.6934 * exp(0.0974 * phi_deg)', 'latex': 'N_{q,B}=0.6934e^{0.0974\\phi}', 'unit': '-', 'soil_type': 'Sand', 'source': 'Advisor note', 'notes': 'Bored empirical Nq; phi in degrees'}, {'parameter': 'Nq', 'equation': 'Nq_driven if pile_type_code > 1.5 else Nq_bored', 'latex': 'N_q=N_{q,D}\\ \\mathrm{or}\\ N_{q,B}', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Bored=1, Driven=2'}, {'parameter': 'Ngamma', 'equation': '2 * (Nq + 1) * tan(radians(phi_deg))', 'latex': 'N_\\gamma=2(N_q+1)\\tan\\phi', 'unit': '-', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Available for custom equations'}, {'parameter': 'qe_clay_t_m2', 'equation': 'Nc * Su_t_m2', 'latex': 'q_e=N_cS_u', 'unit': 't/m²', 'soil_type': 'Clay', 'source': 'Default v8.8', 'notes': 'Clay tip resistance'}, {'parameter': 'qe_sand_t_m2', 'equation': '(Nq - 1) * sigma_v01_t_m2', 'latex': "q_e=(N_q-1)\\sigma'_{v,bottom}", 'unit': 't/m²', 'soil_type': 'Sand', 'source': 'Default v8.8', 'notes': 'Sand tip resistance'}]))

def load_equation_preset_by_name(preset_name, preset_options):
    preset_info = preset_options.get(preset_name, {})
    preset_path = preset_info.get("path")
    if preset_path and Path(preset_path).exists():
        return ensure_equation_columns(pd.read_csv(preset_path)), None
    if preset_name == "Bangkok Excel Method":
        return embedded_bangkok_equations_df(), None
    return None, "Preset file was not found."

def refresh_equation_editor_widget():
    st.session_state["equation_editor_version"] = int(st.session_state.get("equation_editor_version", 0)) + 1

def save_equation_preset_file(name, eq_df):
    preset_name = safe_storage_name(name, "custom_equation_preset")
    if not preset_name.lower().endswith(".csv"):
        preset_name += ".csv"
    EQUATION_PRESET_DIR.mkdir(exist_ok=True)
    path = EQUATION_PRESET_DIR / preset_name
    ensure_equation_columns(eq_df).to_csv(path, index=False)
    return path

def custom_equation_preset_table():
    rows = []
    for label, path in equation_custom_preset_options().items():
        rows.append({
            "delete": False,
            "preset_name": label.replace("Custom: ", ""),
            "file_name": path.name,
            "path": str(path),
        })
    return pd.DataFrame(rows)

def read_equation_upload(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    name = clean_text(getattr(uploaded_file, "name", ""), "").lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return ensure_equation_columns(pd.read_excel(uploaded_file))
    return ensure_equation_columns(pd.read_csv(uploaded_file))



# -----------------------------
# Shallow equation library helpers
# -----------------------------
def default_shallow_equations_df():
    p = get_default_file("shallow_equation_library_general_bearing_capacity.csv")
    if p and Path(p).exists():
        return ensure_shallow_equation_columns(pd.read_csv(p))
    base = shallow_equation_table().rename(columns={"note": "notes"})
    for c in ["unit", "soil_type", "source"]:
        if c not in base.columns:
            base[c] = ""
    return ensure_shallow_equation_columns(base)

def ensure_shallow_equation_columns(eq_df):
    cols = ["parameter", "equation", "latex", "unit", "soil_type", "source", "notes"]
    if eq_df is None or eq_df.empty or "parameter" not in eq_df.columns or "equation" not in eq_df.columns:
        return pd.DataFrame(columns=cols)
    out = eq_df.copy()
    if "note" in out.columns and "notes" not in out.columns:
        out = out.rename(columns={"note": "notes"})
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    return out[cols].fillna("")

def read_shallow_equation_upload(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    name = clean_text(getattr(uploaded_file, "name", ""), "").lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return ensure_shallow_equation_columns(pd.read_excel(uploaded_file))
    return ensure_shallow_equation_columns(pd.read_csv(uploaded_file))

def shallow_equation_builtin_preset_options():
    return {
        "General bearing capacity method": get_default_file("shallow_equation_library_general_bearing_capacity.csv"),
    }

def shallow_equation_custom_preset_options():
    SHALLOW_EQUATION_PRESET_DIR.mkdir(exist_ok=True)
    opts = {}
    for p in sorted(SHALLOW_EQUATION_PRESET_DIR.glob("*.csv")):
        label = "Custom: " + p.stem.replace("_", " ")
        opts[label] = p
    return opts

def shallow_equation_all_preset_options():
    opts = {}
    for name, path in shallow_equation_builtin_preset_options().items():
        opts[name] = {"path": path, "type": "built-in"}
    for name, path in shallow_equation_custom_preset_options().items():
        opts[name] = {"path": path, "type": "custom"}
    return opts

def load_shallow_equation_preset_by_name(preset_name, preset_options):
    preset_info = preset_options.get(preset_name, {})
    preset_path = preset_info.get("path")
    if preset_path and Path(preset_path).exists():
        return ensure_shallow_equation_columns(pd.read_csv(preset_path)), None
    if preset_name == "General bearing capacity method":
        return default_shallow_equations_df(), None
    return None, "Preset file was not found."

def refresh_shallow_equation_editor_widget():
    st.session_state["shallow_equation_editor_version"] = int(st.session_state.get("shallow_equation_editor_version", 0)) + 1

def save_shallow_equation_preset_file(name, eq_df):
    preset_name = safe_storage_name(name, "custom_shallow_equation_preset")
    if not preset_name.lower().endswith(".csv"):
        preset_name += ".csv"
    SHALLOW_EQUATION_PRESET_DIR.mkdir(exist_ok=True)
    path = SHALLOW_EQUATION_PRESET_DIR / preset_name
    ensure_shallow_equation_columns(eq_df).to_csv(path, index=False)
    return path

def custom_shallow_equation_preset_table():
    rows = []
    for label, path in shallow_equation_custom_preset_options().items():
        rows.append({
            "delete": False,
            "preset_name": label.replace("Custom: ", ""),
            "file_name": path.name,
            "path": str(path),
        })
    return pd.DataFrame(rows)

def shallow_equations_df_to_dict(eq_df):
    eq_df = ensure_shallow_equation_columns(eq_df)
    out = {}
    for _, r in eq_df.iterrows():
        parameter = clean_text(r.get("parameter"), "")
        equation = clean_text(r.get("equation"), "")
        if parameter and equation:
            out[parameter] = equation
    return out

def render_equation_library_used(eq_df, title="LaTeX formulas used from Equation Lab", columns=2, shallow=False):
    eq_df = ensure_shallow_equation_columns(eq_df) if shallow else ensure_equation_columns(eq_df)
    if eq_df is None or eq_df.empty:
        st.info("No equation table is available for this run.")
        return
    st.markdown(f"#### {title}")
    show_cols = [c for c in ["parameter", "equation", "unit", "soil_type", "source", "notes"] if c in eq_df.columns]
    with st.expander("Equation table used by this run", expanded=False):
        st.dataframe(display_no_missing(eq_df[show_cols]), use_container_width=True, hide_index=True)
    eq_cols = st.columns(columns)
    for i, (_, r) in enumerate(eq_df.iterrows()):
        with eq_cols[i % columns]:
            st.markdown(f"**{clean_text(r.get('parameter'), '')}**")
            latex = clean_text(r.get("latex"), "")
            if latex:
                st.latex(latex)
            else:
                st.code(clean_text(r.get("equation"), ""))
            meta = []
            if clean_text(r.get("unit"), ""):
                meta.append(f"Unit: `{r.get('unit')}`")
            if clean_text(r.get("source"), ""):
                meta.append(f"Source: `{r.get('source')}`")
            if meta:
                st.caption(" | ".join(meta))


# -----------------------------
# Interface
# -----------------------------
st.markdown("""
<div class="app-title"><h1>🌍 GeoCapacity</h1><p>Foundation capacity calculator for deep and shallow foundations</p></div>
<div class="top-toolbar"><div class="toolbar-chip">📂 Data <span>upload</span></div><div class="toolbar-chip">🧮 Equations <span>standards</span></div><div class="toolbar-chip">▶️ Run <span>calculate</span></div><div class="toolbar-chip">📊 Results <span>check</span></div><div class="toolbar-chip">📐 Charts <span>design</span></div><div class="toolbar-chip">📄 Reports <span>export</span></div></div>
""", unsafe_allow_html=True)


# Display and unit settings are stored in session so they remain active when navigating pages.
if "ton_to_kn" not in st.session_state:
    st.session_state.ton_to_kn = 9.81
if "display_decimals" not in st.session_state:
    st.session_state.display_decimals = 6
TON_TO_KN = float(st.session_state.ton_to_kn)

with st.sidebar:
    st.markdown("### Navigation")
    nav_pages = ["🏠 Home", "🏗️ Deep Foundation", "🧱 Shallow Foundation", "📥 Data Format", "🧮 Equation Lab", "📚 Equation Guide", "📐 Design Charts", "✅ Excel Check", "🗂️ History", "👤 About", "📘 Manual", "🎬 Tutorial"]
    page = st.radio("Choose page", nav_pages, label_visibility="collapsed")
    ui_theme = "Bright"
    st.divider()
    result_state = "<span class='ok'>loaded</span>" if st.session_state.get("current_results") is not None else "<span class='wait'>empty</span>"
    st.markdown("### Project Explorer")
    st.markdown(f"""<div class='project-tree'>🌍 GeoCapacity<br>├─ 📁 Data files: <span class='ok'>CSV</span><br>├─ 🧮 Design method: <span class='ok'>editable</span><br>├─ ⚙️ Calculation: <span class='ok'>ready</span><br>├─ 📈 Results: {result_state}<br>└─ 🗂️ History: <span class='ok'>{len(list_history())} saved</span></div>""", unsafe_allow_html=True)


# Visible version remains GeoCapacity v1.0; deep calculation core restored from v9.10.2.
APP_VERSION = "geocapacity_v1_0_eccentricity_mode_simplified"
if st.session_state.get("app_version") != APP_VERSION:
    # Reset equations on a new app version so corrected default equations are loaded.
    # Users can still upload or edit their own standard in Equation Lab.
    st.session_state.equations_df = default_equations_df()
    st.session_state.shallow_equations_df = default_shallow_equations_df()
    st.session_state.app_version = APP_VERSION

if "equations_df" not in st.session_state:
    st.session_state.equations_df = default_equations_df()
if "shallow_equations_df" not in st.session_state:
    st.session_state.shallow_equations_df = default_shallow_equations_df()
if "current_results" not in st.session_state:
    st.session_state.current_results = None
if "current_errors" not in st.session_state:
    st.session_state.current_errors = None
if "current_inputs" not in st.session_state:
    st.session_state.current_inputs = None
if "current_selected_table" not in st.session_state:
    st.session_state.current_selected_table = None
if "design_chart_tables" not in st.session_state:
    st.session_state.design_chart_tables = None

if page == "🏠 Home":
    st.subheader("Welcome to GeoCapacity")
    st.markdown("""
    <div class='software-note'>
    GeoCapacity is a foundation-capacity workflow app for deep and shallow foundation checking, input validation,
    design review, Excel checking, sensitivity checks, and report export.
    </div>
    """, unsafe_allow_html=True)
    h1, h2, h3 = st.columns(3)
    with h1:
        st.markdown("### 1. Prepare data")
        st.write("Use the CSV templates for pile cases, borehole layers, and shallow footing cases.")
    with h2:
        st.markdown("### 2. Run calculation")
        st.write("Run Deep Foundation or Shallow Foundation and review the calculation tables.")
    with h3:
        st.markdown("### 3. Verify and export")
        st.write("Use Excel Check, Design Charts, reports, and project ZIP export.")
    st.markdown("### Main pages")
    st.dataframe(pd.DataFrame([
        {"page": "Deep Foundation", "purpose": "Pile input, calculation, soil profile, and capacity result"},
        {"page": "Shallow Foundation", "purpose": "Footing input, bearing-capacity checks, and geometry sketch"},
        {"page": "Excel Check", "purpose": "Compare app output with user Excel/template reference for deep and shallow cases"},
        {"page": "Design Charts", "purpose": "Capacity charts, safe pile options, and sensitivity check"},
        {"page": "Tutorial", "purpose": "Reserved area for future YouTube/video guidance"},
    ]), use_container_width=True, hide_index=True)

elif page == "📥 Data Format":
    st.subheader("Data Format")
    deep_format_tab, shallow_format_tab = st.tabs(["Deep foundation", "Shallow foundation"])
    with deep_format_tab:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Pile cases CSV")
            st.dataframe(pd.DataFrame(columns=required_pile_columns()), use_container_width=True)
            p = get_default_file("pile_cases_template.csv")
            if p: st.download_button("Download pile cases template", p.read_bytes(), "pile_cases_template.csv", "text/csv")
            p2 = get_default_file("pile_cases_from_uploaded_excel_BH1.csv")
            if p2: st.download_button("Download uploaded-Excel BH1 pile case", p2.read_bytes(), p2.name, "text/csv")
        with c2:
            st.markdown("### Borehole layers CSV")
            st.dataframe(pd.DataFrame(columns=required_layer_columns()), use_container_width=True)
            p = get_default_file("borehole_layers_template.csv")
            if p: st.download_button("Download borehole layers template", p.read_bytes(), "borehole_layers_template.csv", "text/csv")
            p2 = get_default_file("borehole_layers_from_uploaded_excel_BH1.csv")
            if p2: st.download_button("Download uploaded-Excel BH1 layers", p2.read_bytes(), p2.name, "text/csv")
        st.markdown("### Pile type guide")
        st.dataframe(pile_type_reference_df(), use_container_width=True, hide_index=True)
        st.markdown("### Pile section guide")
        st.dataframe(pile_section_reference_df(), use_container_width=True, hide_index=True)
        c3, c4 = st.columns(2)
        pt_ref = TEMPLATE_DIR / "pile_type_reference.csv"
        sec_ref = TEMPLATE_DIR / "pile_section_reference.csv"
        if pt_ref.exists():
            with c3: st.download_button("Download pile type guide", pt_ref.read_bytes(), "pile_type_reference.csv", "text/csv")
        if sec_ref.exists():
            with c4: st.download_button("Download pile section guide", sec_ref.read_bytes(), "pile_section_reference.csv", "text/csv")
    with shallow_format_tab:
        st.markdown("### Footing cases CSV")
        st.dataframe(pd.DataFrame(columns=required_shallow_columns()), use_container_width=True)
        st.warning("Strip S: use L=1, Cy=1, Oy=0, Mcx=0, Hy=0. Shape factors Fcs, Fqs, Fγs are forced to 1.0 for strip.")
        sc1, sc2 = st.columns(2)
        with sc1:
            p = get_default_file("footing_cases_template.csv")
            if p: st.download_button("Download footing cases template", p.read_bytes(), "footing_cases_template.csv", "text/csv")
        with sc2:
            p2 = get_default_file("footing_cases_sample_from_excel.csv")
            if p2: st.download_button("Download shallow sample", p2.read_bytes(), "footing_cases_sample_from_excel.csv", "text/csv")
        st.markdown("### Brief shallow input guide")
        st.dataframe(pd.DataFrame([
            {"item": "footing_shape", "how to input": "S, R, or Sq", "meaning": "S = strip, R = rectangular, Sq = square"},
            {"item": "beta_deg", "how to input": "Blank or angle in degrees", "meaning": "Blank = beta calculated from Hx_t and Hy_t; filled = use the input angle directly"},
            {"item": "eccentricity_mode", "how to input": "MOMENT or DIRECT", "meaning": "MOMENT: Ox/Oy are load offsets and ex=Mty/Qt, ey=Mtx/Qt. DIRECT: Ox/Oy are used directly as ex/ey."},
            {"item": "Ox_m / Oy_m", "how to input": "Use the same two columns for both modes", "meaning": "MOMENT = load-offset distances. DIRECT = final eccentricities. For strip, Oy/ey is forced to 0."},
        ]), use_container_width=True, hide_index=True)


elif page == "🧮 Equation Lab":
    st.subheader("Equation Lab")
    eq_deep_tab, eq_shallow_tab = st.tabs(["Deep foundation", "Shallow foundation"])
    with eq_deep_tab:
        preset_options = equation_all_preset_options()

        requested_preset = st.session_state.pop("_equation_preset_select_request", None)
        if requested_preset and requested_preset in preset_options:
            st.session_state["equation_preset_dropdown"] = requested_preset
        elif st.session_state.get("equation_preset_dropdown") not in preset_options:
            st.session_state["equation_preset_dropdown"] = list(preset_options.keys())[0] if preset_options else ""

        preset_col, action_col, delete_col = st.columns([1.35, 0.85, 0.85])
        with preset_col:
            preset_name = st.selectbox("Equation preset", list(preset_options.keys()), key="equation_preset_dropdown")
        if preset_name and st.session_state.get("last_loaded_equation_preset") != preset_name:
            loaded_eq, preset_error = load_equation_preset_by_name(preset_name, preset_options)
            if loaded_eq is not None:
                st.session_state.equations_df = loaded_eq.copy()
                st.session_state["last_loaded_equation_preset"] = preset_name
                st.session_state["equation_source_label"] = preset_name
                refresh_equation_editor_widget()
            elif preset_error:
                st.warning(preset_error)

        with action_col:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Load selected preset"):
                loaded_eq, preset_error = load_equation_preset_by_name(preset_name, preset_options)
                if loaded_eq is not None:
                    st.session_state.equations_df = loaded_eq.copy()
                    st.session_state["last_loaded_equation_preset"] = preset_name
                    st.session_state["equation_source_label"] = preset_name
                    refresh_equation_editor_widget()
                    st.success(f"Loaded: {preset_name}")
                else:
                    st.warning(preset_error or "Preset file was not found.")
        with delete_col:
            st.markdown("<br>", unsafe_allow_html=True)
            preset_info = preset_options.get(preset_name, {})
            if preset_info.get("type") == "custom":
                if st.button("Delete selected custom preset"):
                    p = Path(preset_info.get("path"))
                    if p.exists():
                        p.unlink()
                    st.success("Custom preset deleted.")
                    st.rerun()
            else:
                st.caption("Built-in preset")

        st.markdown("#### Upload equation library from device")
        template_c1, template_c2 = st.columns(2)
        with template_c1:
            eq_template_csv = get_default_file("equation_library_template.csv")
            if eq_template_csv:
                st.download_button(
                    "Download equation template CSV",
                    eq_template_csv.read_bytes(),
                    "equation_library_template.csv",
                    "text/csv",
                    key="download_eq_template_csv"
                )
        with template_c2:
            eq_template_xlsx = get_default_file("equation_library_template.xlsx")
            if eq_template_xlsx:
                st.download_button(
                    "Download equation template Excel",
                    eq_template_xlsx.read_bytes(),
                    "equation_library_template.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_eq_template_xlsx"
                )

        uploaded_eq = st.file_uploader("Upload equation library CSV or Excel", type=["csv", "xlsx", "xls"], key="eq_upload")
        if uploaded_eq is not None:
            try:
                uploaded_signature = f"{uploaded_eq.name}|{uploaded_eq.getbuffer().nbytes}"
            except Exception:
                uploaded_signature = clean_text(getattr(uploaded_eq, "name", ""), "uploaded_equation_file")

            reload_uploaded = st.button("Reload uploaded file into table", key="reload_uploaded_equation_file")

            if st.session_state.get("last_eq_upload_signature") != uploaded_signature or reload_uploaded:
                uploaded_df = read_equation_upload(uploaded_eq)
                st.session_state.equations_df = uploaded_df.copy()
                st.session_state["last_eq_upload_signature"] = uploaded_signature
                st.session_state["equation_source_label"] = f"Uploaded temporary: {uploaded_eq.name}"
                refresh_equation_editor_widget()
                st.success("Uploaded equation library loaded into the table below.")
            else:
                st.info("Uploaded file is available. It will not overwrite the selected preset unless you reload it.")

            default_uploaded_name = Path(uploaded_eq.name).stem
            save_cols = st.columns([1.15, 0.85])
            with save_cols[0]:
                upload_preset_name = st.text_input(
                    "Preset name for this uploaded file",
                    value=default_uploaded_name,
                    key="uploaded_equation_preset_name"
                )
            with save_cols[1]:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Save uploaded file to preset dropdown"):
                    preset_df = read_equation_upload(uploaded_eq)
                    preset_path = save_equation_preset_file(upload_preset_name or default_uploaded_name, preset_df)
                    selected_label = "Custom: " + preset_path.stem.replace("_", " ")
                    st.session_state["_equation_preset_select_request"] = selected_label
                    st.session_state["last_loaded_equation_preset"] = selected_label
                    st.success(f"Saved to dropdown: {selected_label}")
                    st.rerun()

        with st.expander("Advanced: manage saved equation presets"):
            expected_password = ""
            try:
                expected_password = st.secrets.get("tutorial_admin_password", "")
            except Exception:
                expected_password = ""
            eq_admin_password = st.text_input("Admin password", type="password", key="equation_admin_password")
            if expected_password and eq_admin_password == expected_password:
                st.markdown("#### Save current equation table as preset")
                preset_save_name = st.text_input("Preset name", value="custom_equation_preset", key="save_current_equation_preset_name")
                if st.button("Save current table as reusable preset"):
                    save_equation_preset_file(preset_save_name, st.session_state.equations_df)
                    st.success("Preset saved.")
                    st.rerun()

                st.markdown("#### Upload reusable preset CSV")
                reusable_upload = st.file_uploader("Upload reusable preset CSV or Excel", type=["csv", "xlsx", "xls"], key="reusable_equation_preset_upload")
                reusable_name = st.text_input("Reusable preset name", value="", key="reusable_equation_preset_name")
                if st.button("Save uploaded file as reusable preset"):
                    if reusable_upload is None:
                        st.warning("Please upload a CSV or Excel file first.")
                    else:
                        name = reusable_name or Path(reusable_upload.name).stem
                        save_equation_preset_file(name, read_equation_upload(reusable_upload))
                        st.success("Reusable preset saved.")
                        st.rerun()

                st.markdown("#### Manage custom presets")
                custom_table = custom_equation_preset_table()
                if custom_table.empty:
                    st.info("No custom equation presets saved yet.")
                else:
                    edited_custom = st.data_editor(
                        custom_table,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "delete": st.column_config.CheckboxColumn("Delete"),
                            "preset_name": st.column_config.TextColumn("Preset name"),
                            "file_name": st.column_config.TextColumn("File name"),
                            "path": st.column_config.TextColumn("Path", disabled=True),
                        },
                        key="custom_equation_preset_manager"
                    )
                    if st.button("Delete checked custom presets"):
                        for _, rr in edited_custom.iterrows():
                            if bool(rr.get("delete", False)):
                                p = Path(clean_text(rr.get("path"), ""))
                                if p.exists() and p.parent == EQUATION_PRESET_DIR:
                                    p.unlink()
                        st.success("Checked presets deleted.")
                        st.rerun()
            elif not expected_password:
                st.warning("Set `tutorial_admin_password` in Streamlit secrets to enable preset management.")
            elif eq_admin_password:
                st.error("Wrong password.")

        top_left, top_right = st.columns([1.45, 1.0])
        with top_left:
            st.markdown("### Editable equation table")
            edited = st.data_editor(
                ensure_equation_columns(st.session_state.equations_df),
                use_container_width=True,
                num_rows="dynamic",
                key=f"equation_editor_{st.session_state.get('equation_editor_version', 0)}",
                column_config={
                    "parameter": st.column_config.TextColumn("Parameter"),
                    "equation": st.column_config.TextColumn("Simple equation", width="large"),
                    "latex": st.column_config.TextColumn("LaTeX preview", width="large"),
                    "unit": st.column_config.TextColumn("Unit"),
                    "soil_type": st.column_config.TextColumn("Soil type"),
                    "source": st.column_config.TextColumn("Source / standard"),
                    "notes": st.column_config.TextColumn("Notes", width="large"),
                },
            )
            current_eq_df = ensure_equation_columns(edited).copy()
            st.session_state.equations_df = current_eq_df

            c1, c2 = st.columns(2)
            with c1:
                if st.button("Check equations"):
                    ok, check_df = validate_equation_set(current_eq_df)
                    st.session_state["equation_check_df"] = check_df
                    if ok:
                        st.success("Equation check passed.")
                    else:
                        st.error("Some equations need correction.")
            with c2:
                st.download_button("Download equations CSV", df_to_csv_bytes(current_eq_df), "equation_library.csv", "text/csv")

            if "equation_check_df" in st.session_state:
                st.markdown("### Equation check result")
                st.dataframe(display_no_missing(st.session_state["equation_check_df"].round(6)), use_container_width=True)

        with top_right:
            st.markdown("### LaTeX equation table")
            for _, r in current_eq_df.iterrows():
                st.markdown(f"**{r['parameter']}**")
                latex = clean_text(r.get("latex"), "")
                if latex:
                    st.latex(latex)
                else:
                    st.code(clean_text(r.get("equation"), ""))
                meta = []
                if clean_text(r.get("unit"), ""):
                    meta.append(f"Unit: `{r.get('unit')}`")
                if clean_text(r.get("source"), ""):
                    meta.append(f"Source: `{r.get('source')}`")
                if meta:
                    st.caption(" | ".join(meta))
            st.markdown("### Nq quick check")
            st.latex(r"N_{q,D}=e^{\pi\tan\phi}\tan^2(45^\circ+\phi/2)(1+\tan\phi)")
            st.latex(r"N_{q,B}=0.6934e^{0.0974\phi}")
            st.dataframe(nq_formula_check_df().round(4), use_container_width=True, hide_index=True)

    with eq_shallow_tab:
        st.caption("This tab manages shallow foundation equation libraries for display, checking, documentation, and reusable presets. It does not change the calculation engine unless the calculation code is later connected to these libraries.")

        shallow_preset_options = shallow_equation_all_preset_options()

        requested_shallow_preset = st.session_state.pop("_shallow_equation_preset_select_request", None)
        if requested_shallow_preset and requested_shallow_preset in shallow_preset_options:
            st.session_state["shallow_equation_preset_dropdown"] = requested_shallow_preset
        elif st.session_state.get("shallow_equation_preset_dropdown") not in shallow_preset_options:
            st.session_state["shallow_equation_preset_dropdown"] = list(shallow_preset_options.keys())[0] if shallow_preset_options else ""

        sp_col, sa_col, sd_col = st.columns([1.35, 0.85, 0.85])
        with sp_col:
            shallow_preset_name = st.selectbox("Shallow equation preset", list(shallow_preset_options.keys()), key="shallow_equation_preset_dropdown")
        if shallow_preset_name and st.session_state.get("last_loaded_shallow_equation_preset") != shallow_preset_name:
            loaded_shallow_eq, shallow_preset_error = load_shallow_equation_preset_by_name(shallow_preset_name, shallow_preset_options)
            if loaded_shallow_eq is not None:
                st.session_state.shallow_equations_df = loaded_shallow_eq.copy()
                st.session_state["last_loaded_shallow_equation_preset"] = shallow_preset_name
                st.session_state["shallow_equation_source_label"] = shallow_preset_name
                refresh_shallow_equation_editor_widget()
            elif shallow_preset_error:
                st.warning(shallow_preset_error)

        with sa_col:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Load selected shallow preset"):
                loaded_shallow_eq, shallow_preset_error = load_shallow_equation_preset_by_name(shallow_preset_name, shallow_preset_options)
                if loaded_shallow_eq is not None:
                    st.session_state.shallow_equations_df = loaded_shallow_eq.copy()
                    st.session_state["last_loaded_shallow_equation_preset"] = shallow_preset_name
                    st.session_state["shallow_equation_source_label"] = shallow_preset_name
                    refresh_shallow_equation_editor_widget()
                    st.success(f"Loaded: {shallow_preset_name}")
                else:
                    st.warning(shallow_preset_error or "Preset file was not found.")
        with sd_col:
            st.markdown("<br>", unsafe_allow_html=True)
            shallow_preset_info = shallow_preset_options.get(shallow_preset_name, {})
            if shallow_preset_info.get("type") == "custom":
                if st.button("Delete selected shallow custom preset"):
                    p = Path(shallow_preset_info.get("path"))
                    if p.exists():
                        p.unlink()
                    st.success("Custom shallow preset deleted.")
                    st.rerun()
            else:
                st.caption("Built-in preset")

        st.markdown("#### Upload shallow equation library from device")
        stc1, stc2 = st.columns(2)
        with stc1:
            shallow_eq_template_csv = get_default_file("shallow_equation_library_template.csv")
            if shallow_eq_template_csv:
                st.download_button(
                    "Download shallow equation template CSV",
                    shallow_eq_template_csv.read_bytes(),
                    "shallow_equation_library_template.csv",
                    "text/csv",
                    key="download_shallow_eq_template_csv"
                )
        with stc2:
            shallow_eq_template_xlsx = get_default_file("shallow_equation_library_template.xlsx")
            if shallow_eq_template_xlsx:
                st.download_button(
                    "Download shallow equation template Excel",
                    shallow_eq_template_xlsx.read_bytes(),
                    "shallow_equation_library_template.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_shallow_eq_template_xlsx"
                )

        uploaded_shallow_eq = st.file_uploader("Upload shallow equation library CSV or Excel", type=["csv", "xlsx", "xls"], key="shallow_eq_upload")
        if uploaded_shallow_eq is not None:
            try:
                shallow_uploaded_signature = f"{uploaded_shallow_eq.name}|{uploaded_shallow_eq.getbuffer().nbytes}"
            except Exception:
                shallow_uploaded_signature = clean_text(getattr(uploaded_shallow_eq, "name", ""), "uploaded_shallow_equation_file")

            reload_shallow_uploaded = st.button("Reload uploaded shallow file into table", key="reload_uploaded_shallow_equation_file")

            if st.session_state.get("last_shallow_eq_upload_signature") != shallow_uploaded_signature or reload_shallow_uploaded:
                uploaded_shallow_df = read_shallow_equation_upload(uploaded_shallow_eq)
                st.session_state.shallow_equations_df = uploaded_shallow_df.copy()
                st.session_state["last_shallow_eq_upload_signature"] = shallow_uploaded_signature
                st.session_state["shallow_equation_source_label"] = f"Uploaded temporary: {uploaded_shallow_eq.name}"
                refresh_shallow_equation_editor_widget()
                st.success("Uploaded shallow equation library loaded into the table below.")
            else:
                st.info("Uploaded file is available. It will not overwrite the selected preset unless you reload it.")

            default_shallow_uploaded_name = Path(uploaded_shallow_eq.name).stem
            shallow_save_cols = st.columns([1.15, 0.85])
            with shallow_save_cols[0]:
                shallow_upload_preset_name = st.text_input(
                    "Preset name for this uploaded shallow file",
                    value=default_shallow_uploaded_name,
                    key="uploaded_shallow_equation_preset_name"
                )
            with shallow_save_cols[1]:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Save uploaded shallow file to preset dropdown"):
                    shallow_preset_df = read_shallow_equation_upload(uploaded_shallow_eq)
                    shallow_preset_path = save_shallow_equation_preset_file(shallow_upload_preset_name or default_shallow_uploaded_name, shallow_preset_df)
                    shallow_selected_label = "Custom: " + shallow_preset_path.stem.replace("_", " ")
                    st.session_state["_shallow_equation_preset_select_request"] = shallow_selected_label
                    st.session_state["last_loaded_shallow_equation_preset"] = shallow_selected_label
                    st.success(f"Saved to dropdown: {shallow_selected_label}")
                    st.rerun()

        with st.expander("Advanced: manage saved shallow equation presets"):
            expected_password = ""
            try:
                expected_password = st.secrets.get("tutorial_admin_password", "")
            except Exception:
                expected_password = ""

            shallow_admin_password = st.text_input("Admin password", type="password", key="shallow_equation_admin_password")
            if expected_password and shallow_admin_password == expected_password:
                st.markdown("#### Save current shallow equation table as preset")
                shallow_preset_save_name = st.text_input("Preset name", value="custom_shallow_equation_preset", key="save_current_shallow_equation_preset_name")
                if st.button("Save current shallow table as reusable preset"):
                    save_shallow_equation_preset_file(shallow_preset_save_name, st.session_state.shallow_equations_df)
                    st.success("Shallow preset saved.")
                    st.rerun()

                st.markdown("#### Upload reusable shallow preset")
                shallow_reusable_upload = st.file_uploader("Upload reusable shallow preset CSV or Excel", type=["csv", "xlsx", "xls"], key="reusable_shallow_equation_preset_upload")
                shallow_reusable_name = st.text_input("Reusable shallow preset name", value="", key="reusable_shallow_equation_preset_name")
                if st.button("Save uploaded shallow file as reusable preset"):
                    if shallow_reusable_upload is None:
                        st.warning("Please upload a CSV or Excel file first.")
                    else:
                        name = shallow_reusable_name or Path(shallow_reusable_upload.name).stem
                        save_shallow_equation_preset_file(name, read_shallow_equation_upload(shallow_reusable_upload))
                        st.success("Reusable shallow preset saved.")
                        st.rerun()

                st.markdown("#### Manage custom shallow presets")
                custom_shallow_table = custom_shallow_equation_preset_table()
                if custom_shallow_table.empty:
                    st.info("No custom shallow equation presets saved yet.")
                else:
                    edited_custom_shallow = st.data_editor(
                        custom_shallow_table,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "delete": st.column_config.CheckboxColumn("Delete"),
                            "preset_name": st.column_config.TextColumn("Preset name"),
                            "file_name": st.column_config.TextColumn("File name"),
                            "path": st.column_config.TextColumn("Path", disabled=True),
                        },
                        key="custom_shallow_equation_preset_manager"
                    )
                    if st.button("Delete checked custom shallow presets"):
                        for _, rr in edited_custom_shallow.iterrows():
                            if bool(rr.get("delete", False)):
                                p = Path(clean_text(rr.get("path"), ""))
                                if p.exists() and p.parent == SHALLOW_EQUATION_PRESET_DIR:
                                    p.unlink()
                        st.success("Checked shallow presets deleted.")
                        st.rerun()
            elif not expected_password:
                st.warning("Set `tutorial_admin_password` in Streamlit secrets to enable shallow preset management.")
            elif shallow_admin_password:
                st.error("Wrong password.")

        shallow_left, shallow_right = st.columns([1.45, 1.0])
        with shallow_left:
            st.markdown("### Editable shallow equation table")
            edited_shallow = st.data_editor(
                ensure_shallow_equation_columns(st.session_state.shallow_equations_df),
                use_container_width=True,
                num_rows="dynamic",
                key=f"shallow_equation_editor_{st.session_state.get('shallow_equation_editor_version', 0)}",
                column_config={
                    "parameter": st.column_config.TextColumn("Parameter"),
                    "equation": st.column_config.TextColumn("Simple equation / relationship", width="large"),
                    "latex": st.column_config.TextColumn("LaTeX preview", width="large"),
                    "unit": st.column_config.TextColumn("Unit"),
                    "soil_type": st.column_config.TextColumn("Factor group"),
                    "source": st.column_config.TextColumn("Source / standard"),
                    "notes": st.column_config.TextColumn("Notes", width="large"),
                },
            )
            current_shallow_eq_df = ensure_shallow_equation_columns(edited_shallow).copy()
            st.session_state.shallow_equations_df = current_shallow_eq_df

            sc1, sc2 = st.columns(2)
            with sc1:
                if st.button("Check shallow equations"):
                    if current_shallow_eq_df.empty:
                        st.warning("The shallow equation table is empty.")
                    elif "parameter" in current_shallow_eq_df.columns and "equation" in current_shallow_eq_df.columns:
                        st.success("Shallow equation table has the required columns.")
            with sc2:
                st.download_button(
                    "Download shallow equations CSV",
                    df_to_csv_bytes(current_shallow_eq_df),
                    "shallow_equation_library.csv",
                    "text/csv"
                )

        with shallow_right:
            st.markdown("### LaTeX shallow equation table")
            for _, rr in current_shallow_eq_df.iterrows():
                st.markdown(f"**{rr['parameter']}**")
                latex = clean_text(rr.get("latex"), "")
                if latex:
                    st.latex(latex)
                else:
                    st.code(clean_text(rr.get("equation"), ""))
                meta = []
                if clean_text(rr.get("unit"), ""):
                    meta.append(f"Unit: `{rr.get('unit')}`")
                if clean_text(rr.get("source"), ""):
                    meta.append(f"Source: `{rr.get('source')}`")
                if meta:
                    st.caption(" | ".join(meta))

elif page == "📚 Equation Guide":
    st.subheader("Equation Guide")
    guide_deep_tab, guide_shallow_tab = st.tabs(["Deep foundation", "Shallow foundation"])
    with guide_deep_tab:
        st.dataframe(equation_parameter_reference_df(), use_container_width=True, height=420)
        st.dataframe(equation_variable_reference_df(), use_container_width=True, height=320)
        st.dataframe(equation_function_reference_df(), use_container_width=True, height=260)
        st.markdown("### Nq")
        st.latex(r"N_{q,D}=e^{\pi\tan\phi}\tan^2(45^\circ+\phi/2)(1+\tan\phi)")
        st.latex(r"N_{q,B}=0.6934e^{0.0974\phi}")
    with guide_shallow_tab:
        eqtab = shallow_equation_table()
        st.dataframe(eqtab[["parameter", "equation", "note"]], use_container_width=True, hide_index=True)
        cols = st.columns(2)
        for i, (_, rr) in enumerate(eqtab.iterrows()):
            with cols[i % 2]:
                st.markdown(f"**{rr['parameter']}**")
                st.latex(rr["latex"])

elif page == "🏗️ Deep Foundation":
    st.subheader("Deep Foundation")
    st.caption("GeoCapacity v1.0")
    tab1, tab2, tab3 = st.tabs(["📁 Input", "⚙️ Run", "📈 Results"])
    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            pile_upload = st.file_uploader("Pile cases CSV", type=["csv"], key="pile_upload")
            default_pile = get_default_file("pile_cases_from_uploaded_excel_BH1.csv") or get_default_file("pile_cases_sample_excel_BH1.csv")
        with c2:
            layer_upload = st.file_uploader("Borehole layers CSV", type=["csv"], key="layer_upload")
            default_layer = get_default_file("borehole_layers_from_uploaded_excel_BH1.csv") or get_default_file("borehole_layers_sample_excel_BH1.csv")
        pile_df = read_csv_from_upload(pile_upload, default_pile)
        layers_df = read_csv_from_upload(layer_upload, default_layer)
        p1, p2 = st.columns(2)
        with p1:
            st.markdown("#### Pile cases preview")
            st.dataframe(display_no_missing(pile_df.head(20)), use_container_width=True)
        with p2:
            st.markdown("#### Borehole layers preview")
            st.dataframe(display_no_missing(layers_df), use_container_width=True)
        st.markdown("#### Deep foundation templates")
        td1, td2 = st.columns(2)
        with td1:
            pile_template_path = get_default_file("pile_cases_template.csv")
            if pile_template_path:
                st.download_button(
                    "Download pile cases template",
                    pile_template_path.read_bytes(),
                    "pile_cases_template.csv",
                    "text/csv",
                    key="deep_tab_pile_cases_template"
                )
        with td2:
            layer_template_path = get_default_file("borehole_layers_template.csv")
            if layer_template_path:
                st.download_button(
                    "Download borehole layers template",
                    layer_template_path.read_bytes(),
                    "borehole_layers_template.csv",
                    "text/csv",
                    key="deep_tab_borehole_layers_template"
                )
        msgs = validate_inputs(pile_df, layers_df) if not pile_df.empty and not layers_df.empty else []
        if msgs:
            for m in msgs: st.error(m)
        else:
            st.success("Input check passed." if not pile_df.empty and not layers_df.empty else "Upload CSV files or use the bundled Excel sample.")
        st.markdown("### Input validation dashboard")
        st.dataframe(deep_input_checklist_df(pile_df, layers_df), use_container_width=True, hide_index=True)
    with tab2:
        r1, r2, r3, r4, r5 = st.columns(5)
        with r1:
            default_output_unit = st.selectbox("**Default result unit**", ["ton", "kN"], index=0)
        with r2:
            output_unit_mode = st.selectbox("**Result unit rule**", ["Follow load unit in CSV", "Use default unit for all piles"], index=0)
        with r3:
            gravity_factor = st.selectbox("**1 ton force equals**", [9.81, 10.0], index=0 if float(st.session_state.get("ton_to_kn", 9.81)) == 9.81 else 1, format_func=lambda x: f"{x:g} kN")
            st.session_state.ton_to_kn = float(gravity_factor)
            TON_TO_KN = float(gravity_factor)
        with r4:
            precision_choice = st.selectbox("**Table decimals**", [1, 2, 3, 4, "Custom"], index=1)
            if precision_choice == "Custom":
                display_decimals = int(st.number_input("Custom decimals", min_value=0, max_value=10, value=int(st.session_state.get("display_decimals", 2) or 2), step=1, key="deep_custom_table_decimals"))
            else:
                display_decimals = int(precision_choice)
            st.session_state.display_decimals = display_decimals
        with r5:
            save_to_history = st.toggle("Save this run", value=True)
        concrete_input_mode = st.selectbox("**Concrete strength f\'c**", ["Use one value for all piles", "Use value from CSV"], index=0)
        cfc_default = 350.0
        if not pile_df.empty and "pile_fc_ksc" in pile_df.columns:
            cfc_default = to_float(pile_df["pile_fc_ksc"].dropna().iloc[0], 350.0) if not pile_df["pile_fc_ksc"].dropna().empty else 350.0
        cfc1, cfc2 = st.columns([1, 2])
        with cfc1:
            concrete_fc_ksc = st.number_input("f'c value (ksc)", value=float(cfc_default), min_value=1.0, step=10.0)
            qe_cap_t_m2 = st.number_input("q_e cap (t/m²)", value=500.0, min_value=0.0, step=25.0)
        with cfc2:
            run_name = st.text_input("Run name", value="pile_capacity_check")
            st.caption("Changing f'c or q_e cap updates the table and graphs after you click Calculate again. q_e,ave uses one row above, current row, and two rows below.")
        output_follows_input = output_unit_mode == "Follow load unit in CSV"
        if output_follows_input:
            st.caption(f"Results follow `required_load_unit` in the CSV. If the CSV unit is blank, the default unit is used. 1 ton force = {TON_TO_KN:g} kN.")
        else:
            st.caption(f"All result tables, plots, and design charts use `{default_output_unit}`. 1 ton force = {TON_TO_KN:g} kN.")
        if st.button("🚀 Run calculation"):
            if pile_df.empty or layers_df.empty:
                st.error("Upload pile cases and borehole layers first.")
            else:
                run_pile = pile_df.copy()
                if concrete_input_mode == "Use one value for all piles":
                    run_pile["pile_fc_ksc"] = concrete_fc_ksc
                elif "pile_fc_ksc" not in run_pile.columns:
                    run_pile["pile_fc_ksc"] = concrete_fc_ksc
                run_pile["qe_cap_t_m2"] = qe_cap_t_m2
                if output_follows_input:
                    if "required_load_unit" in run_pile.columns:
                        run_pile["display_force_unit"] = run_pile["required_load_unit"].fillna(default_output_unit)
                    else:
                        run_pile["display_force_unit"] = default_output_unit
                else:
                    run_pile["display_force_unit"] = default_output_unit
                results, errors = run_batch(run_pile, layers_df, st.session_state.equations_df)
                st.session_state.current_results = results
                st.session_state.current_errors = errors
                st.session_state.current_inputs = (run_pile, layers_df, st.session_state.equations_df.copy(), ui_theme)
                st.session_state.deep_equation_option_used = st.session_state.get("equation_source_label", st.session_state.get("last_loaded_equation_preset", st.session_state.get("equation_preset_dropdown", "Current edited deep equation table")))
                st.session_state.deep_equations_used = st.session_state.equations_df.copy()
                st.session_state.current_gravity_factor = float(TON_TO_KN)
                st.session_state.current_display_decimals = display_decimals
                if save_to_history and not results.empty:
                    save_history(run_name, run_pile, layers_df, st.session_state.equations_df, results)
                st.success(f"Calculation finished: {len(results)} piles calculated, {len(errors)} errors.")
        if st.button("🧹 Clear current interface results"):
            st.session_state.current_results = None
            st.session_state.current_errors = None
            st.session_state.current_inputs = None
            st.session_state.current_selected_table = None
            st.session_state.design_chart_tables = None
            st.success("Current interface results cleared. Saved history was not deleted.")
    with tab3:
        results = st.session_state.current_results
        if results is None or results.empty:
            st.info("No current result. Run the calculation first.")
        else:
            ok = int((results["load_status"] == "OK").sum()) if "load_status" in results else 0
            total = len(results)
            k1,k2,k3,k4 = st.columns(4)
            k1.metric("Total piles", total); k2.metric("Load OK", ok); k3.metric("Load NG", total-ok); k4.metric("Min safety ratio", f"{results['safety_ratio'].min():.2f}")
            st.markdown("### Equation option used")
            st.info(f"Deep foundation equation option: {st.session_state.get('deep_equation_option_used', 'Current edited deep equation table')}")
            st.markdown("### Batch summary table")
            st.dataframe(round_for_display(results, st.session_state.get("current_display_decimals", st.session_state.get("display_decimals", 6))), use_container_width=True, height=280)
            cdl, cdr = st.columns(2)
            with cdl: st.download_button("Download batch results CSV", df_to_csv_bytes(results), "batch_results.csv", "text/csv")
            with cdr: st.download_button("Download batch results Excel", df_to_excel_bytes({"batch_results": results}), "batch_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            errors = st.session_state.current_errors
            if errors is not None and not errors.empty:
                st.warning("Some rows could not be calculated.")
                st.dataframe(errors, use_container_width=True)
            st.markdown("### Selected pile check")
            selected = st.selectbox("Select pile", results["pile_id"].astype(str).tolist())
            run_pile, run_layers, eq_df, saved_theme = st.session_state.current_inputs
            TON_TO_KN = float(st.session_state.get("current_gravity_factor", st.session_state.get("ton_to_kn", 9.81)))
            display_decimals = st.session_state.get("current_display_decimals", st.session_state.get("display_decimals", 6))
            pile_row = run_pile[run_pile["pile_id"].astype(str) == selected].iloc[0]
            case = normalize_pile_case(pile_row)
            bh = clean_text(case.get("boring_no"))
            selected_layers = run_layers[run_layers["boring_no"].astype(str) == bh].copy()
            eqs = equations_df_to_dict(eq_df)
            table_ton = calculate_excel_style_table(selected_layers, case, eqs)
            out_unit = unit_label(case.get("display_force_unit", case.get("required_load_unit", "ton")))
            disp_table = display_calculation_table(table_ton, out_unit, display_decimals)
            st.markdown("#### Aligned soil profile and capacity graphs")
            deep_fig = plot_aligned_panels(selected_layers, table_ton, case, out_unit, ui_theme)
            st.pyplot(deep_fig)
            st.markdown("#### Excel-style calculation table")
            st.dataframe(disp_table, use_container_width=True, height=420)
            st.session_state.current_selected_table = disp_table
            d1,d2 = st.columns(2)
            with d1: st.download_button("Download selected pile table CSV", df_to_csv_bytes(disp_table), f"{selected}_calculation_table.csv", "text/csv")
            with d2: st.download_button("Download selected pile table Excel", df_to_excel_bytes({"calculation_table": disp_table, "batch_results": results}), f"{selected}_calculation_table.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            summary_row = results[results["pile_id"].astype(str) == str(selected)].iloc[0].to_dict()
            st.markdown("#### Result explanation")
            st.dataframe(deep_result_explanation_df(summary_row, table_ton, case), use_container_width=True, hide_index=True)
            with st.expander("Sensitivity analysis", expanded=False):
                sens_df = deep_sensitivity_table(selected_layers, case, eqs, out_unit)
                st.dataframe(round_for_display(sens_df, display_decimals), use_container_width=True, hide_index=True)
                st.download_button("Download deep sensitivity CSV", df_to_csv_bytes(sens_df), f"{selected}_deep_sensitivity.csv", "text/csv")
            if True:
                advanced_zip = selected_project_zip_bytes({
                    f"{selected}_calculation_table.csv": df_to_csv_bytes(disp_table),
                    "batch_results.csv": df_to_csv_bytes(results),
                    "input_pile_cases.csv": df_to_csv_bytes(run_pile),
                    "input_borehole_layers.csv": df_to_csv_bytes(run_layers),
                    "input_equations.csv": df_to_csv_bytes(eq_df),
                })
                st.download_button("Download selected project ZIP", advanced_zip, f"{selected}_project_package.zip", "application/zip")
            rpdf1, rpdf2 = st.columns(2)
            with rpdf1:
                st.download_button("Download selected pile PDF report", selected_pile_pdf_report_bytes(selected, case, summary_row, disp_table, out_unit, deep_fig), f"{selected}_selected_pile_report.pdf", "application/pdf")
            with rpdf2:
                st.download_button("Download batch summary PDF report", batch_summary_pdf_report_bytes(results, errors, run_name="current_run"), "batch_summary_report.pdf", "application/pdf")
            render_equation_library_used(eq_df, title="LaTeX deep formulas used from Equation Lab", columns=3, shallow=False)

elif page == "🧱 Shallow Foundation":
    st.subheader("Shallow Foundation")
    stab1, stab2, stab3 = st.tabs(["📁 Input", "⚙️ Run", "📈 Results"])
    with stab1:
        st.markdown("### Footing cases CSV")
        shallow_upload = st.file_uploader("Shallow footing cases CSV", type=["csv"], key="shallow_upload")
        default_shallow = get_default_file("footing_cases_sample_from_excel.csv") or get_default_file("footing_cases_template.csv")
        shallow_df = read_csv_from_upload(shallow_upload, default_shallow)
        shallow_df = ensure_shallow_columns(shallow_df)
        st.dataframe(display_no_missing(shallow_df), use_container_width=True, height=240)
        for wm in strip_input_warning_messages(shallow_df):
            st.warning(wm)
        sc1, sc2 = st.columns(2)
        with sc1:
            p = get_default_file("footing_cases_template.csv")
            if p:
                st.download_button("Download footing template", p.read_bytes(), "footing_cases_template.csv", "text/csv")
        with sc2:
            p = get_default_file("footing_cases_sample_from_excel.csv")
            if p:
                st.download_button("Download Excel-based shallow sample", p.read_bytes(), "footing_cases_sample_from_excel.csv", "text/csv")
        msgs = validate_shallow_inputs(shallow_df)
        if msgs:
            for m in msgs:
                st.error(m)
        else:
            st.success("Shallow footing input check passed.")
        st.markdown("### Input validation dashboard")
        st.dataframe(shallow_input_checklist_df(shallow_df), use_container_width=True, hide_index=True)
        st.markdown("### Expected CSV columns")
        st.dataframe(pd.DataFrame(columns=required_shallow_columns()), use_container_width=True)
    with stab2:
        r1, r2, r3, r4, r5 = st.columns(5)
        with r1:
            shallow_default_unit = st.selectbox("**Default result unit**", ["ton", "kN"], index=0, key='shallow_default_unit')
        with r2:
            shallow_unit_mode = st.selectbox("**Result unit rule**", ["Follow result_unit in CSV", "Use default unit for all cases"], index=0, key='shallow_unit_mode')
        with r3:
            gravity_factor_sf = st.selectbox("**1 ton force equals**", [9.81, 10.0], index=0 if float(st.session_state.get('ton_to_kn',9.81))==9.81 else 1, format_func=lambda x: f"{x:g} kN", key='shallow_g')
            st.session_state.ton_to_kn = float(gravity_factor_sf)
            TON_TO_KN = float(gravity_factor_sf)
        with r4:
            shallow_precision_choice = st.selectbox("**Table decimals**", [1, 2, 3, 4, "Custom"], index=1, key='shallow_table_decimals')
            if shallow_precision_choice == "Custom":
                shallow_display_decimals = int(st.number_input("Custom decimals", min_value=0, max_value=10, value=int(st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 2)) or 2), step=1, key='shallow_custom_table_decimals'))
            else:
                shallow_display_decimals = int(shallow_precision_choice)
        with r5:
            shallow_save_history = st.toggle("Save shallow run", value=False, key='shallow_save')
        if shallow_unit_mode == "Follow result_unit in CSV":
            st.caption(f"Results follow `result_unit` in the shallow CSV. If the CSV unit is blank, the default unit `{shallow_default_unit}` is used. 1 ton force = {TON_TO_KN:g} kN.")
        else:
            st.caption(f"All shallow result tables use `{shallow_default_unit}`. 1 ton force = {TON_TO_KN:g} kN.")
        if st.button("🚀 Run shallow foundation calculation"):
            for wm in strip_input_warning_messages(shallow_df):
                st.warning(wm)
            msgs = validate_shallow_inputs(shallow_df)
            if msgs:
                for m in msgs:
                    st.error(m)
            else:
                run_shallow_df = shallow_df.copy()
                if shallow_unit_mode == "Follow result_unit in CSV":
                    if "result_unit" in run_shallow_df.columns:
                        run_shallow_df["result_unit"] = run_shallow_df["result_unit"].replace("", np.nan).fillna(shallow_default_unit)
                    else:
                        run_shallow_df["result_unit"] = shallow_default_unit
                else:
                    run_shallow_df["result_unit"] = shallow_default_unit
                sf_results, sf_details, sf_errors = calc_shallow_batch(run_shallow_df, result_unit=shallow_default_unit, equations_df=st.session_state.get('shallow_equations_df', default_shallow_equations_df()))
                st.session_state.shallow_results = sf_results
                st.session_state.shallow_details = sf_details
                st.session_state.shallow_errors = sf_errors
                st.session_state.shallow_inputs = run_shallow_df.copy()
                st.session_state.shallow_equation_option_used = st.session_state.get("shallow_equation_source_label", st.session_state.get("last_loaded_shallow_equation_preset", st.session_state.get("shallow_equation_preset_dropdown", "Current edited shallow equation table")))
                st.session_state.shallow_equations_used = st.session_state.get('shallow_equations_df', default_shallow_equations_df()).copy()
                st.session_state.shallow_display_decimals = shallow_display_decimals
                st.session_state.shallow_result_unit_rule = shallow_unit_mode
                st.session_state.shallow_default_unit_run = shallow_default_unit
                if shallow_save_history and sf_results is not None and not sf_results.empty:
                    folder = HISTORY_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_shallow_preview"
                    folder.mkdir(exist_ok=True)
                    run_shallow_df.to_csv(folder/'input_footing_cases.csv', index=False)
                    sf_results.to_csv(folder/'shallow_results.csv', index=False)
                    st.session_state.get('shallow_equations_df', default_shallow_equations_df()).to_csv(folder/'input_shallow_equations.csv', index=False)
                st.success(f"Shallow calculation finished: {len(sf_results)} cases, {len(sf_errors)} errors.")
        if st.button("🧹 Clear shallow results"):
            st.session_state.shallow_results = None
            st.session_state.shallow_details = None
            st.session_state.shallow_errors = None
            st.session_state.shallow_inputs = None
            st.success("Shallow current results cleared.")
    with stab3:
        sf_results = st.session_state.get('shallow_results')
        if sf_results is None or sf_results.empty:
            st.info("No shallow result yet. Run the shallow module first.")
        else:
            st.markdown("### Equation option used")
            st.info(f"Shallow foundation equation option: {st.session_state.get('shallow_equation_option_used', 'Current edited shallow equation table')}")
            st.markdown("### Shallow footing summary")
            st.dataframe(round_for_display(sf_results, st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 6))), use_container_width=True, height=280)
            cc1, cc2 = st.columns(2)
            with cc1:
                st.download_button("Download shallow results CSV", df_to_csv_bytes(sf_results), "shallow_results.csv", "text/csv")
            with cc2:
                st.download_button("Download shallow results Excel", df_to_excel_bytes({'shallow_results': sf_results}), "shallow_results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if st.session_state.get('shallow_errors') is not None and not st.session_state.get('shallow_errors').empty:
                st.warning("Some shallow rows could not be calculated.")
                st.dataframe(st.session_state.get('shallow_errors'), use_container_width=True)
            selected_case = st.selectbox("Select shallow case", sf_results['case_id'].astype(str).tolist(), key='selected_shallow_case')
            row = st.session_state.shallow_inputs[st.session_state.shallow_inputs['case_id'].astype(str)==selected_case].iloc[0]
            outrow = sf_results[sf_results['case_id'].astype(str)==selected_case].iloc[0]
            drow = st.session_state.shallow_details.get(str(selected_case), pd.DataFrame())
            st.markdown("#### Geometry sketch")
            shallow_fig = plot_shallow_geometry(row, ui_theme)
            st.pyplot(shallow_fig)
            st.markdown("#### Factor table")
            st.dataframe(round_for_display(drow, st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 6))), use_container_width=True)
            if drow is not None and not drow.empty and 'total_term_t_m2' in drow.columns:
                total_term_val = pd.to_numeric(drow['total_term_t_m2'], errors='coerce').fillna(0).sum()
                total_box = pd.DataFrame([{
                    'Total Σ(total_term_t_m2)': total_term_val,
                    'Qall ≥ Qt': clean_text(outrow.get('criterion_Qall_ge_Qt', ''), ''),
                    'qall ≥ qmax': clean_text(outrow.get('criterion_qall_ge_qmax', ''), ''),
                    'qmin ≥ 0': clean_text(outrow.get('criterion_qmin_ge_0', ''), '')
                }])
                st.dataframe(round_for_display(total_box, st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 6))), use_container_width=True, hide_index=True)
            st.markdown("#### Result explanation")
            st.dataframe(shallow_result_explanation_df(outrow, drow), use_container_width=True, hide_index=True)
            with st.expander("Sensitivity analysis", expanded=False):
                sens_sf = shallow_sensitivity_table(row, clean_text(outrow.get("result_unit"), "ton"), equations_df=st.session_state.get("shallow_equations_used", st.session_state.get("shallow_equations_df", default_shallow_equations_df())))
                st.dataframe(round_for_display(sens_sf, st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 2))), use_container_width=True, hide_index=True)
                st.download_button("Download shallow sensitivity CSV", df_to_csv_bytes(sens_sf), f"{selected_case}_shallow_sensitivity.csv", "text/csv")
            if True:
                shallow_zip = selected_project_zip_bytes({
                    f"{selected_case}_summary.csv": df_to_csv_bytes(pd.DataFrame([outrow])),
                    f"{selected_case}_factor_table.csv": df_to_csv_bytes(drow),
                    "input_footing_cases.csv": df_to_csv_bytes(st.session_state.shallow_inputs),
                    "shallow_results.csv": df_to_csv_bytes(sf_results),
                    "input_shallow_equations.csv": df_to_csv_bytes(st.session_state.get("shallow_equations_used", st.session_state.get("shallow_equations_df", default_shallow_equations_df()))),
                })
                st.download_button("Download selected shallow project ZIP", shallow_zip, f"{selected_case}_shallow_project_package.zip", "application/zip")
            st.markdown("#### Selected case summary")
            st.dataframe(round_for_display(pd.DataFrame([outrow]), st.session_state.get('shallow_display_decimals', st.session_state.get('display_decimals', 6))), use_container_width=True)
            sca, scb, scc = st.columns(3)
            with sca:
                st.download_button("Download selected shallow case CSV", df_to_csv_bytes(pd.DataFrame([outrow])), f"{selected_case}_summary.csv", "text/csv")
            with scb:
                st.download_button("Download selected shallow case Excel", df_to_excel_bytes({'summary': pd.DataFrame([outrow]), 'factors': drow}), f"{selected_case}_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with scc:
                st.download_button("Download selected shallow PDF report", selected_shallow_pdf_report_bytes(selected_case, outrow, drow, shallow_fig), f"{selected_case}_shallow_report.pdf", "application/pdf")
            render_equation_library_used(
                st.session_state.get("shallow_equations_used", st.session_state.get("shallow_equations_df", default_shallow_equations_df())),
                title="LaTeX shallow formulas used from Equation Lab",
                columns=2,
                shallow=True
            )
elif page == "📐 Design Charts":
    st.subheader("Design Charts")
    deep_chart_tab, shallow_chart_tab = st.tabs(["Deep foundation", "Shallow foundation"])

    with deep_chart_tab:
        results = st.session_state.current_results
        if results is None or results.empty or st.session_state.current_inputs is None:
            st.info("Run Deep Foundation first.")
        else:
            run_pile, run_layers, eq_df, saved_theme = st.session_state.current_inputs
            TON_TO_KN = float(st.session_state.get("current_gravity_factor", st.session_state.get("ton_to_kn", 9.81)))
            csel1, csel2, csel3 = st.columns([1.1, 1.1, 1.2])
            with csel1:
                selected = st.selectbox("Reference pile", results["pile_id"].astype(str).tolist(), key="design_chart_pile")
            pile_row = run_pile[run_pile["pile_id"].astype(str) == selected].iloc[0]
            case = normalize_pile_case(pile_row)
            bh = clean_text(case.get("boring_no"))
            selected_layers = run_layers[run_layers["boring_no"].astype(str) == bh].copy()
            eqs = equations_df_to_dict(eq_df)
            out_unit = unit_label(case.get("display_force_unit", case.get("required_load_unit", "ton")))
            max_depth = float(selected_layers["bottom_depth_m"].max()) if not selected_layers.empty else 30.0
            with csel2:
                chart_output_unit = st.selectbox("Chart output unit", [out_unit, "kN" if out_unit == "ton" else "ton"], index=0)
            with csel3:
                st.caption(f"Borehole: {bh} | Pile section: {case.get('pile_section','')} | Current unit: {chart_output_unit}")
            r1, r2, r3, r4 = st.columns(4)
            with r1:
                length_min = st.number_input("Length min (m)", value=1.0, min_value=0.0, step=0.5)
                dia_min = st.number_input("Diameter/size min (m)", value=0.30, min_value=0.05, step=0.05)
            with r2:
                length_max = st.number_input("Length max (m)", value=float(max_depth), min_value=0.5, step=0.5)
                dia_max = st.number_input("Diameter/size max (m)", value=max(1.50, to_float(case.get("pile_size"), 0.6)), min_value=0.05, step=0.05)
            with r3:
                length_step = st.number_input("Length increment (m)", value=1.0, min_value=0.1, step=0.1)
                dia_step = st.number_input("Diameter/size increment (m)", value=0.10, min_value=0.01, step=0.01)
            with r4:
                selected_diameters_text = st.text_input("Diameters for Qall-Length", value="0.30,0.40,0.50,0.80,1.00")
                selected_lengths_text = st.text_input("Lengths for Qall-Diameter", value="10,15,20,30")
            all_lengths = make_float_range(length_min, length_max, length_step)
            all_diameters = make_float_range(dia_min, dia_max, dia_step)
            selected_diameters = parse_number_list(selected_diameters_text, [to_float(case.get("pile_size"), 0.6)])
            selected_lengths = parse_number_list(selected_lengths_text, [to_float(case.get("pile_length_m"), 20.0)])
            with st.expander("Design recommendation engine", expanded=False):
                req_default = force_from_ton(force_to_kn(case.get("required_load_value", 0), case.get("required_load_unit", "ton")) / TON_TO_KN, chart_output_unit)
                req_design = st.number_input(f"Required load for recommendation ({cap_unit(chart_output_unit)})", value=float(req_default), min_value=0.0, step=10.0)
                top_n = st.number_input("Number of safe options to show", value=5, min_value=1, max_value=20, step=1)
                if st.button("Find safe pile options"):
                    rec_rows = []
                    for L0 in all_lengths:
                        for D0 in all_diameters:
                            vals = qall_for_design(selected_layers, case, eqs, D0, L0, chart_output_unit)
                            q_allow = vals["Qall"]
                            if req_design <= 0 or q_allow >= req_design:
                                rec_rows.append({
                                    "Diameter/size (m)": D0,
                                    "Length (m)": L0,
                                    f"Qall ({cap_unit(chart_output_unit)})": q_allow,
                                    "safety ratio": q_allow / req_design if req_design > 0 else 0.0,
                                    f"Qult ({cap_unit(chart_output_unit)})": vals["Qult"],
                                    f"Qs ({cap_unit(chart_output_unit)})": vals["Qs"],
                                    f"Qb ({cap_unit(chart_output_unit)})": vals["Qb"],
                                })
                    rec_df = pd.DataFrame(rec_rows)
                    if rec_df.empty:
                        st.warning("No safe option found in the selected range.")
                    else:
                        rec_df = rec_df.sort_values(["Length (m)", "Diameter/size (m)"]).head(int(top_n))
                        st.dataframe(round_for_display(rec_df, st.session_state.get("display_decimals", 2)), use_container_width=True, hide_index=True)
                        st.download_button("Download safe options CSV", df_to_csv_bytes(rec_df), f"{selected}_safe_pile_options.csv", "text/csv")
            if st.button("Generate deep design charts"):
                with st.spinner("Generating design charts..."):
                    length_table = capacity_vs_length_table(selected_layers, case, eqs, selected_diameters, all_lengths, chart_output_unit)
                    diameter_table = capacity_vs_diameter_table(selected_layers, case, eqs, all_diameters, selected_lengths, chart_output_unit)
                    heat_table = capacity_heatmap_table(selected_layers, case, eqs, all_diameters, all_lengths, chart_output_unit)
                    st.session_state.design_chart_tables = (selected, length_table, diameter_table, heat_table, chart_output_unit, ui_theme)
                st.success("Design charts generated.")
            if "design_chart_tables" in st.session_state and st.session_state.design_chart_tables is not None:
                selected_saved, length_table, diameter_table, heat_table, chart_output_unit, theme_saved = st.session_state.design_chart_tables
                st.markdown(f"### Charts for {selected_saved}")
                g1, g2 = st.columns(2)
                with g1:
                    st.pyplot(plot_capacity_vs_length(length_table, chart_output_unit, ui_theme))
                with g2:
                    st.pyplot(plot_capacity_vs_diameter(diameter_table, chart_output_unit, ui_theme))
                st.pyplot(plot_capacity_heatmap(heat_table, chart_output_unit, ui_theme))
                t1, t2, t3 = st.tabs(["Qall vs length table", "Qall vs diameter table", "Heatmap data"])
                with t1:
                    st.dataframe(length_table, use_container_width=True, height=300)
                    st.download_button("Download Qall vs length CSV", df_to_csv_bytes(length_table), f"{selected_saved}_qall_vs_length.csv", "text/csv")
                with t2:
                    st.dataframe(diameter_table, use_container_width=True, height=300)
                    st.download_button("Download Qall vs diameter CSV", df_to_csv_bytes(diameter_table), f"{selected_saved}_qall_vs_diameter.csv", "text/csv")
                with t3:
                    st.dataframe(heat_table, use_container_width=True, height=300)
                    st.download_button("Download heatmap data CSV", df_to_csv_bytes(heat_table), f"{selected_saved}_capacity_heatmap_data.csv", "text/csv")
                st.download_button("Download all design chart tables as Excel", df_to_excel_bytes({"qall_vs_length": length_table, "qall_vs_diameter": diameter_table, "heatmap_data": heat_table}), f"{selected_saved}_design_charts.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with shallow_chart_tab:
        sf_results = st.session_state.get("shallow_results")
        if sf_results is None or sf_results.empty or st.session_state.get("shallow_inputs") is None:
            st.info("Run Shallow Foundation first.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                selected_case = st.selectbox("Reference shallow case", sf_results["case_id"].astype(str).tolist(), key="shallow_design_case")
            row = st.session_state.shallow_inputs[st.session_state.shallow_inputs["case_id"].astype(str) == selected_case].iloc[0]
            outrow = sf_results[sf_results["case_id"].astype(str) == selected_case].iloc[0]
            with c2:
                shallow_unit = st.selectbox("Chart output unit", [clean_text(outrow.get("result_unit"), "ton"), "kN" if clean_text(outrow.get("result_unit"), "ton") == "ton" else "ton"], key="shallow_design_unit")
            r1, r2, r3 = st.columns(3)
            with r1:
                B_min = st.number_input("B min (m)", value=max(to_float(row.get("B_m"), 1.5)*0.5, 0.1), min_value=0.1, step=0.1)
                D_min = st.number_input("D min (m)", value=max(to_float(row.get("D_m"), 1.0)*0.5, 0.1), min_value=0.1, step=0.1)
            with r2:
                B_max = st.number_input("B max (m)", value=max(to_float(row.get("B_m"), 1.5)*1.5, 0.2), min_value=0.1, step=0.1)
                D_max = st.number_input("D max (m)", value=max(to_float(row.get("D_m"), 1.0)*1.5, 0.2), min_value=0.1, step=0.1)
            with r3:
                B_step = st.number_input("B increment (m)", value=0.1, min_value=0.01, step=0.05)
                D_step = st.number_input("D increment (m)", value=0.1, min_value=0.01, step=0.05)
            B_values = make_float_range(B_min, B_max, B_step)
            D_values = make_float_range(D_min, D_max, D_step)
            selected_B_values = parse_number_list(st.text_input("B values for Qall-D", value=f"{to_float(row.get('B_m'), 1.5):g}"), [to_float(row.get("B_m"), 1.5)])
            selected_D_values = parse_number_list(st.text_input("D values for Qall-B", value=f"{to_float(row.get('D_m'), 1.0):g}"), [to_float(row.get("D_m"), 1.0)])
            if st.button("Generate shallow design charts"):
                B_table = shallow_design_vs_B_table(row, B_values, selected_D_values, shallow_unit)
                D_table = shallow_design_vs_D_table(row, D_values, selected_B_values, shallow_unit)
                st.session_state.shallow_design_tables = (selected_case, B_table, D_table, shallow_unit)
            if st.session_state.get("shallow_design_tables") is not None:
                case_saved, B_table, D_table, unit_saved = st.session_state.shallow_design_tables
                g1, g2 = st.columns(2)
                with g1:
                    st.pyplot(plot_shallow_design_table(B_table, "B (m)", "Qall vs footing width B", unit_saved, ui_theme))
                    st.dataframe(round_for_display(B_table, st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2))), use_container_width=True)
                with g2:
                    st.pyplot(plot_shallow_design_table(D_table, "D (m)", "Qall vs foundation depth D", unit_saved, ui_theme))
                    st.dataframe(round_for_display(D_table, st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2))), use_container_width=True)
                st.download_button("Download shallow design charts Excel", df_to_excel_bytes({"qall_vs_B": B_table, "qall_vs_D": D_table}), f"{case_saved}_shallow_design_charts.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "✅ Excel Check":
    st.subheader("Excel Check")
    st.caption("Compare GeoCapacity output with a user Excel-ready calculation template or reference table.")
    deep_check_tab, shallow_check_tab = st.tabs(["Deep foundation", "Shallow foundation"])

    with deep_check_tab:
        results = st.session_state.current_results
        if results is None or results.empty or st.session_state.current_inputs is None:
            st.info("Run Deep Foundation first, then return here for deep Excel checking.")
            summary_path = APP_DIR / "deep_excel_verification_summary.csv"
            if summary_path.exists():
                st.markdown("### Bundled deep-foundation Excel verification summary")
                st.dataframe(display_no_missing(pd.read_csv(summary_path)), use_container_width=True, hide_index=True)
        else:
            run_pile, run_layers, eq_df, saved_theme = st.session_state.current_inputs
            selected_verify = st.selectbox("Select pile for deep verification", results["pile_id"].astype(str).tolist(), key="excel_verify_pile")
            pile_row = run_pile[run_pile["pile_id"].astype(str) == selected_verify].iloc[0]
            case = normalize_pile_case(pile_row)
            bh = clean_text(case.get("boring_no"))
            selected_layers = run_layers[run_layers["boring_no"].astype(str) == bh].copy()
            eqs = equations_df_to_dict(eq_df)
            table_ton = calculate_excel_style_table(selected_layers, case, eqs)
            out_unit = unit_label(case.get("display_force_unit", case.get("required_load_unit", "ton")))
            decimals_for_verify = st.session_state.get("current_display_decimals", st.session_state.get("display_decimals", 2))
            app_table = display_calculation_table(table_ton, out_unit, decimals_for_verify)
            st.markdown("### Current app deep calculation table")
            st.dataframe(app_table, use_container_width=True, height=280)
            deep_formula_template = get_default_file("deep_excel_check_template.xlsx")
            if deep_formula_template:
                st.download_button("Download deep Excel template with formulas", deep_formula_template.read_bytes(), "deep_excel_check_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("Download deep reference table template", df_to_excel_bytes({"deep_reference_table": app_table.iloc[0:0]}), "deep_reference_table_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            ref_upload = st.file_uploader("Upload deep Excel/reference table", type=["csv", "xlsx", "xls"], key="deep_reference_upload")
            ref_df = read_reference_table(ref_upload) if ref_upload is not None else None
            if ref_df is None or ref_df.empty:
                st.info("Upload your Excel-ready calculation output to compare with the current app table.")
            else:
                summary, detail = generic_reference_comparison(app_table, ref_df, decimals_for_verify)
                st.markdown("### Difference summary")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                st.markdown("### Row-by-row difference")
                st.dataframe(detail, use_container_width=True, height=360)
                st.download_button("Download deep verification summary CSV", df_to_csv_bytes(summary), f"{selected_verify}_deep_excel_verification_summary.csv", "text/csv")
                st.download_button("Download deep row-by-row difference CSV", df_to_csv_bytes(detail), f"{selected_verify}_deep_excel_row_difference.csv", "text/csv")

    with shallow_check_tab:
        sf_results = st.session_state.get("shallow_results")
        if sf_results is None or sf_results.empty:
            st.info("Run Shallow Foundation first, then return here for shallow Excel checking.")
        else:
            selected_case = st.selectbox("Select shallow case for verification", sf_results["case_id"].astype(str).tolist(), key="excel_verify_shallow_case")
            outrow = sf_results[sf_results["case_id"].astype(str) == selected_case].iloc[0]
            drow = st.session_state.shallow_details.get(str(selected_case), pd.DataFrame())
            app_summary = round_for_display(pd.DataFrame([outrow]), st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2)))
            app_factor = round_for_display(drow, st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2)))
            st.markdown("### Current app shallow summary")
            st.dataframe(app_summary, use_container_width=True)
            st.markdown("### Current app shallow factor table")
            st.dataframe(app_factor, use_container_width=True, height=220)
            shallow_formula_template = get_default_file("shallow_excel_check_template.xlsx")
            if shallow_formula_template:
                st.download_button("Download shallow Excel template with formulas", shallow_formula_template.read_bytes(), "shallow_excel_check_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("Download shallow reference template", df_to_excel_bytes({"shallow_summary": app_summary.iloc[0:0], "shallow_factor_table": app_factor.iloc[0:0]}), "shallow_reference_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            shallow_ref_upload = st.file_uploader("Upload shallow Excel/reference table", type=["csv", "xlsx", "xls"], key="shallow_reference_upload")
            shallow_ref_df = read_reference_table(shallow_ref_upload) if shallow_ref_upload is not None else None
            if shallow_ref_df is None or shallow_ref_df.empty:
                st.info("Upload your shallow Excel-ready calculation output to compare with the app output.")
            else:
                summary, detail = generic_reference_comparison(app_summary, shallow_ref_df, st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2)))
                if summary.empty:
                    summary, detail = generic_reference_comparison(app_factor, shallow_ref_df, st.session_state.get("shallow_display_decimals", st.session_state.get("display_decimals", 2)))
                st.markdown("### Difference summary")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                st.markdown("### Row-by-row difference")
                st.dataframe(detail, use_container_width=True, height=360)
                st.download_button("Download shallow verification summary CSV", df_to_csv_bytes(summary), f"{selected_case}_shallow_excel_verification_summary.csv", "text/csv")
                st.download_button("Download shallow row-by-row difference CSV", df_to_csv_bytes(detail), f"{selected_case}_shallow_excel_row_difference.csv", "text/csv")

elif page == "🗂️ History":
    st.subheader("History")
    histories = list_history()
    if not histories:
        st.info("No saved history yet.")
    else:
        selected = st.selectbox("Saved runs", [p.name for p in histories])
        folder = HISTORY_DIR / selected
        files = sorted(folder.glob("*.csv"))
        for f in files:
            st.download_button(f"Download {f.name}", f.read_bytes(), f.name, "text/csv")
        st.download_button("Download selected history ZIP", make_zip_bytes(files), f"{selected}.zip", "application/zip")
        if st.button("Delete selected history"):
            for f in files: f.unlink(missing_ok=True)
            try: folder.rmdir()
            except Exception: pass
            st.success("Deleted selected history.")

elif page == "👤 About":
    st.subheader("About GeoCapacity")
    st.markdown("""
    ### 🌍 GeoCapacity

    **GeoCapacity** is a geotechnical foundation capacity calculator developed by **David An**.  
    The app has expanded from a deep-foundation pile calculator into a broader **foundation-capacity platform** that can check both **deep foundations** and **shallow foundations** in one workflow.

    ### Developer

    **David An**  
    Master of Engineering in Geotechnical Engineering, Chulalongkorn University  
    Bachelor of Science in Civil Engineering, Paragon International University, Cambodia

    ### Why the scope was expanded

    The original version focused on **bored and driven pile capacity** using borehole-layer data and editable equations.  
    The current version adds a **shallow foundation module** so the same app can also calculate footing bearing capacity for early-stage foundation comparison, teaching, checking, and future AI-ready design workflows.

    ### Current scope

    **1. Deep foundation module**
    - Bored pile and driven pile capacity checking
    - Borehole-layer based calculation
    - Supported pile sections: circular, square, hexagonal, hollow circular, hollow square, I-section, and custom geometry
    - Automatic or user-defined pile area and perimeter
    - Editable equation library with LaTeX display
    - Soil profile, shaft resistance, end bearing, capacity graphs, calculation table, and PDF report export

    **2. Shallow foundation module**
    - Strip footing and rectangular footing capacity checking
    - Total stress and effective stress analysis options
    - Water-table influence on unit weight above and below the footing base
    - Load, moment, eccentricity, contact stress, effective footing dimensions, and bearing-capacity factor checks
    - Excel-style factor table, geometry sketch, summary table, CSV/Excel export, and PDF report export

    ### Intended users

    GeoCapacity is intended for students, researchers, and engineers who need a clear calculation workflow for comparing pile and footing capacity, preparing reports, checking assumptions, and organizing foundation design data.

    ### Long-term vision

    GeoCapacity can be expanded into a broader geotechnical platform for foundation capacity, borehole interpretation, 3D subsoil modeling, reliability analysis, and AI-assisted geotechnical decision support.
    """)

elif page == "📘 Manual":
    st.subheader("GeoCapacity Manual")
    st.markdown("""
    ## 🌍 GeoCapacity v1.0

    GeoCapacity is a Streamlit-based geotechnical calculation app for **foundation bearing/capacity checking**.  
    The app now covers two main design modules:

    - **Deep Foundation**: bored and driven pile capacity
    - **Shallow Foundation**: strip and rectangular footing bearing capacity

    The purpose of this upgrade is to make GeoCapacity a more complete foundation-capacity tool instead of a pile-only calculator.

    ---

    ## 1. Overall workflow

    1. Prepare the required CSV input files.
    2. Open the suitable module: **Deep Foundation** or **Shallow Foundation**.
    3. Upload the CSV files or use the bundled sample files.
    4. Check the preview table and input warnings.
    5. Run the calculation.
    6. Review tables, plots, factors, and status checks.
    7. Export CSV, Excel, or PDF reports.

    ---

    ## 2. Deep Foundation module

    The Deep Foundation module calculates pile capacity using borehole-layer data.

    ### Required deep-foundation files

    **Pile cases CSV**
    - pile ID
    - borehole number
    - pile type
    - pile section
    - pile size
    - inner size when needed
    - pile length
    - load and unit
    - water table
    - concrete strength
    - optional area/perimeter override

    **Borehole layers CSV**
    - borehole number
    - top and bottom depth
    - total unit weight
    - soil type/name
    - undrained shear strength, `su1`
    - SPT-N value

    ### Supported pile types

    | Code | Meaning |
    |---|---|
    | `B` | Bored pile / cast-in-place pile |
    | `D` | Driven pile |

    ### Supported pile sections

    | Code | Section | Important input |
    |---|---|---|
    | `C` | Solid circular | `pile_size` = diameter |
    | `S` | Solid square | `pile_size` = width |
    | `H`, `HEX`, `HEXAGON` | Solid hexagonal | `pile_size` = across flats |
    | `SC` | Hollow spun circular | `pile_size` + `inner_size` |
    | `SS` | Hollow square | `pile_size` + `inner_size` |
    | `I` | I-section | use exact `pile_perimeter` and `pile_cross_section_area` if available |
    | `CUSTOM` | User-defined section | provide `pile_perimeter` and `pile_cross_section_area` |

    ### Deep-foundation outputs

    The app produces:

    - batch summary table
    - selected pile calculation table
    - shaft resistance `Qs`
    - end bearing `Qe`
    - ultimate capacity `Qult`
    - allowable capacity `Qall`
    - concrete capacity check
    - required load check
    - aligned soil profile and capacity plots
    - selected pile PDF report
    - batch PDF report

    ---

    ## 3. Shallow Foundation module

    The Shallow Foundation module was added to expand GeoCapacity beyond pile design.  
    It calculates footing bearing capacity using an Excel-style workflow for shallow foundations.

    ### Supported shallow foundation types

    | Code | Meaning |
    |---|---|
    | `S` | Strip footing |
    | `R` | Rectangular footing |

    ### Supported analysis types

    | Code | Meaning |
    |---|---|
    | `T` | Total stress analysis |
    | `E` | Effective stress analysis |

    ### Main shallow-foundation inputs

    | Input | Meaning |
    |---|---|
    | `B_m` | footing width |
    | `L_m` | footing length |
    | `D_m` | footing base depth |
    | `t_m` | footing thickness |
    | `Cx_m`, `Cy_m` | column size |
    | `Ox_m`, `Oy_m` | column/load offset |
    | `water_table_depth_m` | water table depth from ground surface |
    | `Qc_t` | vertical applied column load |
    | `Mcx_tm`, `Mcy_tm` | applied moments |
    | `Hx_t`, `Hy_t` | horizontal loads |
    | `g1u_t_m3`, `g1sat_t_m3` | soil unit weight above footing base |
    | `c1_t_m2`, `phi1_deg` | strength parameters above footing base |
    | `g2u_t_m3`, `g2sat_t_m3` | soil unit weight below footing base |
    | `c2_t_m2`, `phi2_deg` | strength parameters below footing base |
    | `gc_t_m3`, `gw_t_m3` | concrete and water unit weights |
    | `FS` | factor of safety |
    | `result_unit` | `ton` or `kN` |

    ### Strip footing note

    For Strip `S`, the app follows the Excel-style strip footing convention:

    - use `L = 1 m`
    - use `Cy = 1 m`
    - use `Oy = 0`
    - use `Mcx = 0`
    - use `Hy = 0`

    The app displays this warning when strip footing input is detected.

    ### Shallow-foundation calculations

    The shallow module checks:

    - footing self-weight and soil weight above footing
    - total vertical load `Qt`
    - eccentricity `eB`, `eL`
    - effective dimensions `B'`, `L'`
    - contact stress `qmax`, `qmin`
    - bearing factors `Nc`, `Nq`, `Nγ`
    - shape factors
    - depth factors
    - inclination factors
    - ultimate bearing pressure `qult`
    - allowable bearing pressure `qall`
    - ultimate load `Qult`
    - allowable load `Qall`
    - stability/status criteria

    ### Shallow-foundation outputs

    The app produces:

    - shallow footing summary table
    - selected case geometry sketch
    - bearing-capacity factor table
    - selected case summary table
    - CSV, Excel, and PDF report export

    ---

    ## 4. Data Format page

    The **Data Format** page is divided into:

    - Deep foundation CSV format
    - Shallow foundation CSV format

    Use this page to download templates before preparing new inputs.

    ---

    ## 5. Equation Lab and Equation Guide

    **Equation Lab** and **Equation Guide** are divided into:

    - Deep foundation equations
    - Shallow foundation equations

    Deep-foundation equations are editable. Shallow-foundation equations are displayed as a guide to explain the footing bearing-capacity calculations.

    ---

    ## 6. GitHub / Streamlit deployment

    For GitHub or Streamlit Community Cloud deployment, keep these files/folders in the repository root:

    - `streamlit_app.py`
    - `requirements.txt`
    - `README.md`
    - `templates/`
    - `samples/`

    Then deploy with main file:

    ```text
    streamlit_app.py
    ```

    Do not upload confidential borehole or client data to a public repository.
    """)


    st.markdown("---")
    st.markdown("### Uploaded PDF manuals")
    manual_df = load_manual_registry()

    if manual_df.empty:
        st.info("No admin-uploaded PDF manual yet.")
    else:
        for i, r in manual_df.iterrows():
            title = clean_text(r.get("title"), f"Manual {i+1}")
            desc = clean_text(r.get("description"), "")
            file_name = clean_text(r.get("file_name"), "")
            st.markdown(f"**{title}**")
            if desc:
                st.write(desc)
            data = manual_file_bytes(file_name)
            if data:
                st.download_button(
                    "Download PDF",
                    data,
                    file_name if file_name else "manual.pdf",
                    "application/pdf",
                    key=f"manual_pdf_download_{i}"
                )
            else:
                st.warning(f"File not found: {file_name}")

    with st.expander("Admin manual upload"):
        expected_password = ""
        try:
            expected_password = st.secrets.get("tutorial_admin_password", "")
        except Exception:
            expected_password = ""

        manual_admin_password = st.text_input("Admin password", type="password", key="manual_admin_password")
        if expected_password and manual_admin_password == expected_password:
            st.markdown("#### Add / replace manual PDF")
            mt1, mt2 = st.columns(2)
            with mt1:
                manual_title = st.text_input("Manual title", key="manual_upload_title")
            with mt2:
                uploaded_manual = st.file_uploader("Upload PDF manual", type=["pdf"], key="manual_pdf_upload")
            manual_description = st.text_area("Manual description", key="manual_upload_description")

            if st.button("Add PDF manual"):
                if uploaded_manual is None:
                    st.warning("Please upload a PDF file first.")
                else:
                    save_manual_upload(uploaded_manual, manual_title, manual_description)
                    st.success("Manual PDF uploaded.")
                    st.rerun()

            st.markdown("#### Manage uploaded manuals")
            manage_manual_df = load_manual_registry()
            if manage_manual_df.empty:
                st.info("No manual file to manage yet.")
            else:
                manage_manual_df = manage_manual_df.copy()
                manage_manual_df.insert(0, "delete", False)
                edited_manuals = st.data_editor(
                    manage_manual_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "delete": st.column_config.CheckboxColumn("Delete"),
                        "title": st.column_config.TextColumn("Title"),
                        "file_name": st.column_config.TextColumn("File name", disabled=True),
                        "description": st.column_config.TextColumn("Description", width="large"),
                        "uploaded_at": st.column_config.TextColumn("Uploaded at", disabled=True),
                    },
                    key="manual_admin_table"
                )

                mc1, mc2 = st.columns(2)
                with mc1:
                    if st.button("Save manual table changes"):
                        save_df = edited_manuals.copy()
                        if "delete" in save_df.columns:
                            for _, rr in save_df[save_df["delete"] == True].iterrows():
                                p = MANUAL_UPLOAD_DIR / safe_storage_name(rr.get("file_name"), "")
                                if p.exists():
                                    p.unlink()
                            save_df = save_df[save_df["delete"] != True].drop(columns=["delete"])
                        save_manual_registry(save_df)
                        st.success("Manual table saved.")
                        st.rerun()
                with mc2:
                    if st.button("Delete all uploaded manuals"):
                        for _, rr in manage_manual_df.iterrows():
                            p = MANUAL_UPLOAD_DIR / safe_storage_name(rr.get("file_name"), "")
                            if p.exists():
                                p.unlink()
                        save_manual_registry(pd.DataFrame(columns=["title", "file_name", "description", "uploaded_at"]))
                        st.success("All uploaded manuals deleted.")
                        st.rerun()
        elif not expected_password:
            st.warning("Set `tutorial_admin_password` in Streamlit secrets to enable admin upload.")
        elif manual_admin_password:
            st.error("Wrong password.")


elif page == "🎬 Tutorial":
    st.subheader("Tutorial")
    tutorial_file = APP_DIR / "tutorial_links.csv"

    tutorial_df, tutorial_storage = load_tutorial_links(tutorial_file)

    if not tutorial_df.empty:
        for _, r in tutorial_df.iterrows():
            title = clean_text(r.get("title"), "Tutorial")
            url = clean_text(r.get("url"), "")
            desc = clean_text(r.get("description"), "")
            if url:
                st.markdown(f"**[{title}]({url})**")
            else:
                st.markdown(f"**{title}**")
            if desc:
                st.write(desc)
    else:
        st.info("No tutorial link added yet.")

    with st.expander("Admin upload"):
        expected_password = ""
        try:
            expected_password = st.secrets.get("tutorial_admin_password", "")
        except Exception:
            expected_password = ""

        admin_password = st.text_input("Admin password", type="password")

        if expected_password and admin_password == expected_password:
            st.caption(f"Storage: {tutorial_storage}")
            st.markdown("### Add tutorial link")
            t1, t2 = st.columns(2)
            with t1:
                new_title = st.text_input("Title", key="tutorial_new_title")
            with t2:
                new_url = st.text_input("URL", key="tutorial_new_url")
            new_description = st.text_area("Description", key="tutorial_new_description")

            if st.button("Add tutorial"):
                new_row = pd.DataFrame([{"title": new_title, "url": new_url, "description": new_description}])
                tutorial_df = pd.concat([tutorial_df, new_row], ignore_index=True)
                save_tutorial_links(tutorial_df, tutorial_file)
                st.rerun()

            st.markdown("### Manage tutorial links")
            manage_df = tutorial_df.copy()
            manage_df.insert(0, "delete", False)
            edited_tutorials = st.data_editor(
                manage_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "delete": st.column_config.CheckboxColumn("Delete"),
                    "title": st.column_config.TextColumn("Title"),
                    "url": st.column_config.LinkColumn("URL"),
                    "description": st.column_config.TextColumn("Description", width="large"),
                },
                key="tutorial_admin_table"
            )

            c1, c2 = st.columns(2)
            with c1:
                if st.button("Save changes"):
                    save_df = edited_tutorials.copy()
                    if "delete" in save_df.columns:
                        save_df = save_df[save_df["delete"] != True].drop(columns=["delete"])
                    save_tutorial_links(save_df, tutorial_file)
                    st.rerun()
            with c2:
                if st.button("Delete all tutorial links"):
                    save_tutorial_links(_tutorial_empty_df(), tutorial_file)
                    st.rerun()

        elif not expected_password:
            st.warning("Set `tutorial_admin_password` in Streamlit secrets to enable admin upload.")
        elif admin_password:
            st.error("Wrong password.")

