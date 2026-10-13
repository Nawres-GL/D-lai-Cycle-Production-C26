from __future__ import annotations

import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Campagne 26 · Délais & Cycle Production", page_icon="🚚", layout="wide", initial_sidebar_state="expanded")


# ==================== STYLE CSS PREMIUM - NEXT-LEVEL DESIGN ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* ===== ANIMATIONS ===== */
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    @keyframes floatCard {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-8px); }
    }
    
    @keyframes glowPulse {
        0%, 100% { box-shadow: 0 0 20px rgba(6, 182, 212, 0.3); }
        50% { box-shadow: 0 0 40px rgba(6, 182, 212, 0.6); }
    }
    
    @keyframes slideInLeft {
        from { opacity: 0; transform: translateX(-20px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    @keyframes slideInRight {
        from { opacity: 0; transform: translateX(20px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    
    @keyframes shimmer {
        0% { background-position: -1000px 0; }
        100% { background-position: 1000px 0; }
    }
    
    @keyframes borderGlow {
        0% { border-color: rgba(6, 182, 212, 0.5); }
        50% { border-color: rgba(16, 185, 129, 0.8); }
        100% { border-color: rgba(6, 182, 212, 0.5); }
    }
    
    /* ===== GLOBAL ===== */
    .stApp {
        background: linear-gradient(-45deg, #0f172a, #1a2d4d, #1e293b, #0f172a);
        background-size: 400% 400%;
        animation: gradientShift 15s ease infinite;
    }
    
    html {
        scroll-behavior: smooth;
    }
    
    /* ===== SIDEBAR ===== */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.95) 0%, rgba(26, 45, 77, 0.9) 100%);
        border-right: 2px solid;
        border-image: linear-gradient(180deg, rgba(6, 182, 212, 0.5), rgba(16, 185, 129, 0.5)) 1;
        backdrop-filter: blur(20px);
    }
    
    /* ===== HEADERS WITH ANIMATED GRADIENTS ===== */
    h1 {
        background: linear-gradient(135deg, #06b6d4 0%, #10b981 50%, #06b6d4 100%);
        background-size: 200% 200%;
        animation: gradientShift 8s ease infinite;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 800;
        letter-spacing: -1px;
        font-size: 2.8rem !important;
        margin-bottom: 0.5rem;
        animation: slideInLeft 0.6s ease-out;
    }
    
    h2 {
        background: linear-gradient(135deg, #06b6d4 0%, #10b981 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 700;
        font-size: 2rem !important;
        margin-top: 1.5rem;
        animation: slideInLeft 0.6s ease-out 0.1s both;
    }
    
    h3 {
        background: linear-gradient(135deg, #22d3ee 0%, #34d399 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 600;
        animation: slideInLeft 0.6s ease-out 0.2s both;
    }
    
    /* ===== PREMIUM ALERT CARDS ===== */
    .alert-critical {
        background: linear-gradient(135deg, rgba(220, 38, 38, 0.2) 0%, rgba(185, 28, 28, 0.1) 100%);
        backdrop-filter: blur(15px);
        border: 1.5px solid rgba(220, 38, 38, 0.4);
        border-left: 4px solid #dc2626;
        color: #fecaca;
        padding: 20px;
        border-radius: 14px;
        margin: 12px 0;
        font-weight: 500;
        box-shadow: 0 10px 40px rgba(220, 38, 38, 0.15), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        animation: slideInRight 0.5s ease-out;
        position: relative;
        overflow: hidden;
    }
    
    .alert-critical::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.1), transparent);
        animation: shimmer 3s infinite;
    }
    
    .alert-critical:hover {
        background: linear-gradient(135deg, rgba(220, 38, 38, 0.3) 0%, rgba(185, 28, 28, 0.15) 100%);
        box-shadow: 0 20px 60px rgba(220, 38, 38, 0.2), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        transform: translateX(6px) translateY(-2px);
        border-color: rgba(220, 38, 38, 0.6);
    }
    
    .alert-warning {
        background: linear-gradient(135deg, rgba(245, 158, 11, 0.2) 0%, rgba(217, 119, 6, 0.1) 100%);
        backdrop-filter: blur(15px);
        border: 1.5px solid rgba(245, 158, 11, 0.4);
        border-left: 4px solid #f59e0b;
        color: #fed7aa;
        padding: 20px;
        border-radius: 14px;
        margin: 12px 0;
        font-weight: 500;
        box-shadow: 0 10px 40px rgba(245, 158, 11, 0.15), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        animation: slideInRight 0.5s ease-out;
        position: relative;
        overflow: hidden;
    }
    
    .alert-warning::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.1), transparent);
        animation: shimmer 3s infinite;
    }
    
    .alert-warning:hover {
        background: linear-gradient(135deg, rgba(245, 158, 11, 0.3) 0%, rgba(217, 119, 6, 0.15) 100%);
        box-shadow: 0 20px 60px rgba(245, 158, 11, 0.2), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        transform: translateX(6px) translateY(-2px);
        border-color: rgba(245, 158, 11, 0.6);
    }
    
    .alert-success {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.2) 0%, rgba(5, 150, 105, 0.1) 100%);
        backdrop-filter: blur(15px);
        border: 1.5px solid rgba(16, 185, 129, 0.4);
        border-left: 4px solid #10b981;
        color: #a7f3d0;
        padding: 20px;
        border-radius: 14px;
        margin: 12px 0;
        font-weight: 500;
        box-shadow: 0 10px 40px rgba(16, 185, 129, 0.15), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        animation: slideInRight 0.5s ease-out;
        position: relative;
        overflow: hidden;
    }
    
    .alert-success::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.1), transparent);
        animation: shimmer 3s infinite;
    }
    
    .alert-success:hover {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.3) 0%, rgba(5, 150, 105, 0.15) 100%);
        box-shadow: 0 20px 60px rgba(16, 185, 129, 0.2), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        transform: translateX(6px) translateY(-2px);
        border-color: rgba(16, 185, 129, 0.6);
    }
    
    /* ===== RECOMMENDATION CARDS ===== */
    .recommendation-box {
        background: linear-gradient(135deg, rgba(6, 182, 212, 0.2) 0%, rgba(8, 145, 178, 0.1) 100%);
        backdrop-filter: blur(15px);
        border: 1.5px solid rgba(6, 182, 212, 0.4);
        border-left: 4px solid #06b6d4;
        color: #a5f3fc;
        padding: 20px;
        border-radius: 14px;
        margin: 12px 0;
        font-weight: 500;
        box-shadow: 0 10px 40px rgba(6, 182, 212, 0.15), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        animation: floatCard 3s ease-in-out infinite, slideInRight 0.5s ease-out;
        position: relative;
        overflow: hidden;
    }
    
    .recommendation-box::after {
        content: '';
        position: absolute;
        top: -50%;
        right: -50%;
        width: 100%;
        height: 100%;
        background: radial-gradient(circle, rgba(6, 182, 212, 0.2) 0%, transparent 70%);
        animation: glowPulse 3s ease-in-out infinite;
    }
    
    .recommendation-box:hover {
        background: linear-gradient(135deg, rgba(6, 182, 212, 0.3) 0%, rgba(8, 145, 178, 0.15) 100%);
        box-shadow: 0 20px 60px rgba(6, 182, 212, 0.25), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        transform: translateY(-6px) translateX(4px);
        border-color: rgba(6, 182, 212, 0.7);
    }
    
    /* ===== EXPLANATION CARDS ===== */
    .explanation-text {
        background: linear-gradient(135deg, rgba(79, 172, 254, 0.12) 0%, rgba(59, 130, 246, 0.06) 100%);
        backdrop-filter: blur(15px);
        border: 1.5px solid rgba(79, 172, 254, 0.25);
        border-left: 3px solid #4facfe;
        color: #bfdbfe;
        padding: 18px;
        border-radius: 12px;
        margin: 12px 0;
        font-size: 0.95rem;
        line-height: 1.7;
        box-shadow: 0 8px 32px rgba(79, 172, 254, 0.1), inset 0 1px 0 rgba(255, 255, 255, 0.05);
        transition: all 0.3s ease;
        animation: slideInLeft 0.5s ease-out;
    }
    
    .explanation-text:hover {
        background: linear-gradient(135deg, rgba(79, 172, 254, 0.18) 0%, rgba(59, 130, 246, 0.08) 100%);
        border-color: rgba(79, 172, 254, 0.4);
        transform: translateX(4px);
    }
    
    /* ===== METRIC CARDS ===== */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.9) 0%, rgba(30, 41, 59, 0.7) 100%);
        backdrop-filter: blur(15px);
        border: 1px solid rgba(51, 65, 85, 0.6);
        border-top: 2px solid transparent;
        border-image: linear-gradient(135deg, #06b6d4, #10b981) 1;
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        animation: floatCard 4s ease-in-out infinite;
    }
    
    [data-testid="stMetric"]:hover {
        transform: translateY(-8px);
        box-shadow: 0 20px 60px rgba(6, 182, 212, 0.2), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        border-top-color: #06b6d4;
    }
    
    /* ===== DATA FRAMES ===== */
    [data-testid="stDataFrame"] {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.6) 0%, rgba(30, 41, 59, 0.4) 100%) !important;
        border: 1px solid rgba(51, 65, 85, 0.4) !important;
        border-radius: 14px !important;
        overflow: hidden !important;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2) !important;
        backdrop-filter: blur(10px) !important;
    }
    
    /* ===== DIVIDERS ===== */
    hr {
        background: linear-gradient(90deg, transparent, rgba(6, 182, 212, 0.4), rgba(16, 185, 129, 0.4), transparent);
        border: none;
        height: 2px;
        margin: 2.5rem 0;
        animation: borderGlow 4s ease-in-out infinite;
    }
    
    /* ===== TEXT STYLING ===== */
    p, span, label {
        color: #e2e8f0;
        font-weight: 400;
        transition: color 0.3s ease;
    }
    
    strong {
        color: #f1f5f9;
        font-weight: 700;
        background: linear-gradient(135deg, #22d3ee, #34d399);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    /* ===== BUTTONS ===== */
    button {
        background: linear-gradient(135deg, #06b6d4 0%, #10b981 100%) !important;
        background-size: 200% 200%;
        border: none !important;
        color: white !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 8px 24px rgba(6, 182, 212, 0.3) !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.5px !important;
        cursor: pointer !important;
    }
    
    button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 40px rgba(6, 182, 212, 0.5) !important;
        animation: gradientShift 2s ease infinite !important;
    }
    
    button:active {
        transform: translateY(-1px) !important;
    }
    
    /* ===== SCROLL STYLING ===== */
    ::-webkit-scrollbar {
        width: 12px;
        height: 12px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(15, 23, 42, 0.5);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #06b6d4, #10b981);
        border-radius: 10px;
        transition: all 0.3s ease;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(180deg, #10b981, #06b6d4);
        box-shadow: 0 0 10px rgba(6, 182, 212, 0.5);
    }
</style>
""", unsafe_allow_html=True)

def dataframe_to_xlsx_bytes(dataframe: pd.DataFrame) -> bytes:
    """Create a professionally styled Royal Plants Excel workbook."""

    output = io.BytesIO()

    # Royal Plants color palette
    DARK_GREEN = "164B35"
    FOREST_GREEN = "216747"
    LIME_GREEN = "A3C76D"
    LIGHT_GREEN = "EAF3E5"
    WHITE = "FFFFFF"
    TEXT_DARK = "20352A"
    BORDER_GREEN = "C8DCCB"
    AMBER = "FFF0C2"
    RED = "FCE0DC"
    GREEN_OK = "D7F0DF"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Timeline",
            startrow=2
        )

        workbook = writer.book
        worksheet = writer.sheets["Timeline"]

        # --------------------------------------------------
        # 1. Professional title section
        # --------------------------------------------------
        last_col = max(dataframe.shape[1], 1)
        last_letter = get_column_letter(last_col)

        worksheet.merge_cells(f"A1:{last_letter}1")
        title = worksheet["A1"]
        title.value = "ROYAL PLANTS | SUIVI CHRONOLOGIQUE"
        title.font = Font(
            name="Aptos Display",
            size=16,
            bold=True,
            color=WHITE
        )
        title.fill = PatternFill(
            "solid",
            fgColor=DARK_GREEN
        )
        title.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        worksheet.row_dimensions[1].height = 36

        # Subtitle
        worksheet.merge_cells(f"A2:{last_letter}2")
        subtitle = worksheet["A2"]
        subtitle.value = (
            f"Suivi des lots et des retards | "
            f"{len(dataframe)} enregistrements"
        )
        subtitle.font = Font(
            name="Aptos",
            size=10,
            italic=True,
            color=WHITE
        )
        subtitle.fill = PatternFill(
            "solid",
            fgColor=FOREST_GREEN
        )
        subtitle.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )
        worksheet.row_dimensions[2].height = 24

        # --------------------------------------------------
        # 2. Header styling
        # --------------------------------------------------
        header_row = 3
        data_start = header_row + 1
        data_end = header_row + len(dataframe)

        for cell in worksheet[header_row]:
            cell.fill = PatternFill(
                "solid",
                fgColor=DARK_GREEN
            )
            cell.font = Font(
                name="Aptos",
                bold=True,
                color=WHITE,
                size=10
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = Border(
                bottom=Side(
                    style="medium",
                    color=LIME_GREEN
                )
            )

        worksheet.row_dimensions[header_row].height = 34

        # --------------------------------------------------
        # 3. Body styling and alternating row colors
        # --------------------------------------------------
        for row_idx in range(data_start, data_end + 1):
            for cell in worksheet[row_idx]:
                cell.fill = PatternFill(
                    "solid",
                    fgColor=(
                        WHITE
                        if row_idx % 2 == 0
                        else LIGHT_GREEN
                    )
                )
                cell.font = Font(
                    name="Aptos",
                    size=10,
                    color=TEXT_DARK
                )
                cell.alignment = Alignment(
                    vertical="center",
                    horizontal="left"
                )
                cell.border = Border(
                    bottom=Side(
                        style="hair",
                        color=BORDER_GREEN
                    )
                )

            worksheet.row_dimensions[row_idx].height = 23

        # --------------------------------------------------
        # 4. Column widths and date formatting
        # --------------------------------------------------
        for col_idx, column_name in enumerate(
            dataframe.columns, start=1
        ):
            letter = get_column_letter(col_idx)

            values = [
                str(column_name),
                *[
                    str(value)
                    for value in dataframe[column_name].dropna().head(100)
                ]
            ]

            max_length = max(
                (len(value) for value in values),
                default=12
            )

            worksheet.column_dimensions[letter].width = min(
                max(max_length + 3, 14),
                30
            )

            name_lower = str(column_name).lower()

            for row_idx in range(data_start, data_end + 1):
                cell = worksheet.cell(row_idx, col_idx)

                # Real Excel dates
                if (
                    pd.api.types.is_datetime64_any_dtype(
                        dataframe[column_name]
                    )
                    or "date" in name_lower
                    or "ds. réel" in name_lower
                ):
                    cell.number_format = "dd/mm/yyyy"
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

                # Duration and delay columns
                if (
                    "jours" in name_lower
                    or "(j)" in name_lower
                    or "cycle" in name_lower
                ):
                    cell.number_format = "0"
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

        # --------------------------------------------------
        # 5. Highlight delay values automatically
        # --------------------------------------------------
        for col_idx, column_name in enumerate(
            dataframe.columns, start=1
        ):
            name_lower = str(column_name).lower()

            if "retard" not in name_lower:
                continue

            letter = get_column_letter(col_idx)
            cell_range = f"{letter}{data_start}:{letter}{data_end}"

            if data_end < data_start:
                continue

            # No delay: green
            worksheet.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="lessThanOrEqual",
                    formula=["0"],
                    fill=PatternFill("solid", fgColor=GREEN_OK),
                    font=Font(color="176B3A", bold=True)
                )
            )

            # Moderate delay: amber
            worksheet.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="between",
                    formula=["1", "3"],
                    fill=PatternFill("solid", fgColor=AMBER),
                    font=Font(color="8A5700", bold=True)
                )
            )

            # Significant delay: red
            worksheet.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="greaterThan",
                    formula=["3"],
                    fill=PatternFill("solid", fgColor=RED),
                    font=Font(color="A61B1B", bold=True)
                )
            )

        # --------------------------------------------------
        # 6. Navigation and print settings
        # --------------------------------------------------
        worksheet.freeze_panes = "A4"

        # Filter the actual data, excluding title rows
        if not dataframe.empty and len(dataframe.columns) > 0:
            worksheet.auto_filter.ref = (
                f"A{header_row}:{last_letter}{data_end}"
            )

        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.page_setup.orientation = "landscape"
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25

        worksheet.print_title_rows = "1:3"

        # Active sheet
        worksheet.sheet_view.zoomScale = 85
        worksheet.sheet_properties.tabColor = LIME_GREEN

    output.seek(0)
    return output.getvalue()

st.set_page_config(
    page_title="Tableau de Bord Production Royal Plants",
    page_icon="rp_logo.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

def format_number(value):
    """Format number with thousand separators"""
    if pd.isna(value):
        return "N/A"
    if isinstance(value, (int, float)):
        return "{:,.0f}".format(value) if isinstance(value, float) and value == int(value) else "{:,.2f}".format(value)
    return str(value)

def format_dataframe(df, numeric_cols=None):
    """Format numeric columns in dataframe with thousand separators"""
    df_formatted = df.copy()
    if numeric_cols is None:
        numeric_cols = df_formatted.select_dtypes(include=['number']).columns
    
    for col in numeric_cols:
        if col in df_formatted.columns:
            df_formatted[col] = df_formatted[col].apply(lambda x: format_number(x) if pd.notna(x) else "N/A")
    
    return df_formatted
def get_unique_lot_records(df):
    """
    Creates one record per production lot without summing repeated
    production rows.

    The original dataframe is not modified.
    Quantity fields are taken once per lot.
    """

    result = df.copy()

    if "LOT" not in result.columns:
        return result.copy()

    # Nettoyage uniquement du numéro de lot
    result["LOT"] = (
        result["LOT"]
        .astype(str)
        .str.strip()
    )

    # Colonnes qui identifient le lot
    keys = ["LOT"]

    for col in ["Espèce", "Article", "Variete"]:
        if col in result.columns:
            keys.append(col)

    # IMPORTANT :
    # On garde UNE ligne représentative de chaque lot.
    # On ne somme PAS les lignes répétées.
    result = (
        result
        .drop_duplicates(
            subset=keys,
            keep="first"
        )
        .copy()
    )

    return result

def render_metric_card(label, value, unit="", color_gradient="cyan-to-green"):
    """Render a professional metric card with gradient"""
    color_map = {
        "cyan-to-green": "linear-gradient(135deg, #06b6d4 0%, #10b981 100%)",
        "red-to-orange": "linear-gradient(135deg, #ef4444 0%, #f97316 100%)",
        "blue-to-purple": "linear-gradient(135deg, #3b82f6 0%, #a855f7 100%)",
        "amber-to-orange": "linear-gradient(135deg, #f59e0b 0%, #fb923c 100%)"
    }
    
    gradient = color_map.get(color_gradient, color_map["cyan-to-green"])
    formatted_value = format_number(value) if isinstance(value, (int, float)) else str(value)
    
    st.markdown(f"""
    <div style='
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8) 0%, rgba(30, 41, 59, 0.6) 100%);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(51, 65, 85, 0.5);
        border-top: 2px solid;
        border-image: {gradient} 1;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
        transition: all 0.3s ease;
    ' onmouseover="this.style.transform='translateY(-4px)'; this.style.boxShadow='0 12px 48px rgba(0, 0, 0, 0.3)';" 
      onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 8px 32px rgba(0, 0, 0, 0.2)';">
        <p style='color: #94a3b8; font-size: 0.9rem; margin: 0; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;'>{label}</p>
        <p style='color: #f1f5f9; font-size: 2.2rem; margin: 8px 0 0 0; font-weight: 700; background: {gradient}; -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;'>{formatted_value}</p>
        <p style='color: #64748b; font-size: 0.85rem; margin: 6px 0 0 0;'>{unit}</p>
    </div>
    """, unsafe_allow_html=True)
    
    
def extract_lot_number(lot_str):
    """Extrait le N° Lot du format 'RO-13362/25' -> '13362'"""
    if pd.isna(lot_str):
        return None
    lot_str = str(lot_str).strip()
    # Format: RO-13362/25 ou similaire
    parts = lot_str.split('-')
    if len(parts) >= 2:
        number_part = parts[-1].split('/')[0]  # Prend la partie après le '-' et avant le '/'
        try:
            return int(number_part)
        except ValueError:
            return None
    return None

    

def load_delivery_data():
    """Charge les données de livraison réelle depuis Livraison_C25.xlsx"""
    try:
        # Les headers commencent à la ligne 4 (index 3 en pandas)
        delivery_raw = pd.read_excel('data/Livraison_26.xlsx')
        delivery_df = delivery_raw.copy()
        delivery_df = delivery_df.dropna(axis=1, how='all')
        
        # Convertir les dates
        if 'Date Effet' in delivery_df.columns:
            delivery_df['Date Effet'] = pd.to_datetime(delivery_df['Date Effet'], errors='coerce')
        
        # Convertir N° Lot en entier pour la jointure
        if 'N° Lot' in delivery_df.columns:
            delivery_df['Lot_Number'] = pd.to_numeric(delivery_df['N° Lot'], errors='coerce')
        
        return delivery_df
    except Exception as e:
        st.warning(f"Impossible de charger Livraison_C25.xlsx: {str(e)}")
        return None

@st.cache_data
def load_and_clean_data():
    """Charge et nettoie les données avec gestion rigoureuse"""
    try:
        franc_raw = pd.read_excel('data/Franc_26.xlsx')
        
        
        # ===== NETTOYAGE FRANC =====
        franc_df = franc_raw.copy()
        franc_df = franc_df.dropna(axis=1, how='all')
        franc_df = franc_df.drop_duplicates()
        
        numeric_cols_franc = franc_df.select_dtypes(include=['number']).columns.tolist()
        for col in numeric_cols_franc:
            franc_df[col] = pd.to_numeric(franc_df[col], errors='coerce')
        
        if 'Taux_Germination' in franc_df.columns:
            franc_df = franc_df.dropna(subset=['Taux_Germination'])
            franc_df = franc_df[(franc_df['Taux_Germination'] >= 0) & (franc_df['Taux_Germination'] <= 100)]
        
        if 'Article' in franc_df.columns:
            franc_df = franc_df.dropna(subset=['Article'])
        
        # Ensure Variete doesn't have NaN values for page 11 (prediction page)
        if 'Variete' in franc_df.columns:
            franc_df['Variete'] = franc_df['Variete'].fillna('Unknown')
        
        # Extraire le N° Lot pour jointure avec Livraison_C25
        if 'LOT' in franc_df.columns:
            franc_df['Lot_Number'] = franc_df['LOT'].apply(extract_lot_number)
        
       
        return franc_df
    except Exception as e:
        st.error(f"Erreur chargement: {str(e)}")
        return None, None

franc_df = load_and_clean_data()

if franc_df is None:
    st.stop()
    
# Charger les données de livraison réelle
delivery_df = load_delivery_data()

# Joindre les données de livraison avec franc_df si disponibles
if delivery_df is not None and 'Lot_Number' in franc_df.columns and 'Lot_Number' in delivery_df.columns:
    try:
        # Sélectionner uniquement les colonnes qui existent
        delivery_merge_cols = ['Lot_Number']
        if 'Date Effet' in delivery_df.columns:
            delivery_merge_cols.append('Date Effet')
        
        franc_df = franc_df.merge(
            delivery_df[delivery_merge_cols],
            on='Lot_Number',
            how='left'
        )
    except Exception as e:
        st.warning(f"Erreur fusion données livraison: {str(e)}")
        
        
logo_path = Path("rp_logo.png")

col_logo, col_title = st.columns([0.07, 0.88], vertical_alignment="center")

with col_logo:
    if logo_path.exists():
        st.image(str(logo_path), width=350)

with col_title:
    st.header("Analyse Délais & Cycle de Production")
    
    # ==================== DESCRIPTION ====================
st.markdown("""
    <div class='explanation-text' style='background: linear-gradient(135deg, rgba(6, 182, 212, 0.15) 0%, rgba(16, 185, 129, 0.1) 100%);'>
    <strong>Évaluation Complète Supply Chain & Production</strong><br>
    Cette page vous permet d'analyser le respect des délais, la qualité de la planification, le cycle de production réel et les ajustements réalisés. 
    Filtrez par Espèce, Article et Lot (optionnel) pour une analyse détaillée et décisionnelle de votre production.
    </div>
    """, unsafe_allow_html=True)
    
st.divider()
    
    # ==================== PRÉPARATION DES DONNÉES ====================
analysis_data = franc_df.copy()
    
    # Convertir les colonnes de dates
date_columns = ['DS. Réel', 'Date Liv Demander', 'Date Liv Prevu', 'Date Liv. Ajuster']
for col in date_columns:
    if col in analysis_data.columns:
        analysis_data[col] = pd.to_datetime(analysis_data[col], errors='coerce')
    
    # ==================== FILTRES HIÉRARCHIQUES ====================
st.subheader("Filtres Hiérarchiques")
    
filter_col1, filter_col2, filter_col3 = st.columns(3)
    
    # FILTRE 1: ESPÈCE
with filter_col1:
    if 'Espèce' in analysis_data.columns:
        species_list = sorted(analysis_data['Espèce'].dropna().unique().tolist())
        selected_species = st.selectbox("Espèce", species_list, key="delivery_species")
        data_by_species = analysis_data[analysis_data['Espèce'] == selected_species].copy()
    else:
        st.error("Colonne 'Espèce' non trouvée")
        data_by_species = analysis_data.copy()
        selected_species = "N/A"
    
    # FILTRE 2: ARTICLE (cascade depuis espèce)
with filter_col2:
    if 'Article' in data_by_species.columns:
        article_list = sorted(data_by_species['Article'].dropna().unique().tolist())
        if article_list:
            selected_article = st.selectbox("Article", article_list, key="delivery_article")
            data_by_article = data_by_species[data_by_species['Article'] == selected_article].copy()
        else:
            st.warning("Aucun article pour cette espèce")
            selected_article = "N/A"
            data_by_article = data_by_species.copy()
    else:
        st.error("Colonne 'Article' non trouvée")
        selected_article = "N/A"
        data_by_article = data_by_species.copy()
    
    # FILTRE 3: LOT (optionnel, cascade depuis article)
with filter_col3:
    if 'LOT' in data_by_article.columns:
        lot_list = sorted(data_by_article['LOT'].dropna().unique().tolist())
        lot_options = ["📊 Tous les lots"] + lot_list
        selected_lot_display = st.selectbox("LOT (optionnel)", lot_options, key="delivery_lot")
            
        if selected_lot_display == "📊 Tous les lots":
            filtered_data = data_by_article.copy()
            is_single_lot = False
        else:
            filtered_data = data_by_article[data_by_article['LOT'] == selected_lot_display].copy()
            is_single_lot = True
    else:
        filtered_data = data_by_article.copy()
        selected_lot_display = "N/A"
        is_single_lot = False
    
st.divider()
    
if len(filtered_data) == 0:
        
        st.error(f"Aucune donnée disponible pour: {selected_species} - {selected_article}")
else:
        # ==================== CALCULS DATES & CYCLES ====================
    calc_data = filtered_data.copy()
        
        # Convertir Date Effet en datetime si elle existe
    if 'Date Effet' in calc_data.columns:
        calc_data['Date Effet'] = pd.to_datetime(calc_data['Date Effet'], errors='coerce')
        
        # Cycle prévu: DS Réel → Liv Prévue
    if 'DS. Réel' in calc_data.columns and 'Date Liv Prevu' in calc_data.columns:
        calc_data['Cycle_Prevu_Jours'] = (calc_data['Date Liv Prevu'] - calc_data['DS. Réel']).dt.days
    else:
        calc_data['Cycle_Prevu_Jours'] = np.nan
        
        # Cycle réel: DS Réel → Liv Ajustée (ou Prévu si pas ajustée)
    if 'DS. Réel' in calc_data.columns:
        calc_data['Cycle_Reel_Jours'] = calc_data.apply(
            lambda row: (row['Date Effet'] - row['DS. Réel']).days 
            if pd.notna(row['Date Effet']) 
            else (row['Date Effet'] - row['DS. Réel']).days if pd.notna(row['Date Effet']) else np.nan,
            axis=1
            )
    else:
        calc_data['Cycle_Reel_Jours'] = np.nan
        
        # Retard Semis: Liv Prévue - Liv Demandée (délai entre demande et prévision)
    if 'Date Liv Demander' in calc_data.columns and 'Date Liv Prevu' in calc_data.columns:
        calc_data['Retard_Semis_Jours'] = (calc_data['Date Liv Prevu'] - calc_data['Date Liv Demander']).dt.days
    else:
        calc_data['Retard_Semis_Jours'] = np.nan
        
        # Retard Élevage: Liv Ajustée - Liv Prévue (délai supplémentaire de la phase d'élevage)
    if 'Date Liv. Ajuster' in calc_data.columns and 'Date Liv Prevu' in calc_data.columns:
        calc_data['Retard_Élevage_Jours'] = (calc_data['Date Liv. Ajuster'] - calc_data['Date Liv Prevu']).dt.days
    else:
        calc_data['Retard_Élevage_Jours'] = np.nan
        
        # Retard Ajusté-Réel: Date Effet - Liv Ajustée (réalité vs prévision ajustée)
    if 'Date Effet' in calc_data.columns and 'Date Liv. Ajuster' in calc_data.columns:
        calc_data['Retard_Ajuste_Reel_Jours'] = (calc_data['Date Effet'] - calc_data['Date Liv. Ajuster']).dt.days
    else:
        calc_data['Retard_Ajuste_Reel_Jours'] = np.nan
        
        # Retard Total: Somme de tous les retards (Semis + Élevage + Ajusté-Réel)
    calc_data['Retard_Total_Jours'] = (
            calc_data['Retard_Semis_Jours'].fillna(0) +
            calc_data['Retard_Élevage_Jours'].fillna(0) +
            calc_data['Retard_Ajuste_Reel_Jours'].fillna(0)
        )
        
        # Satisfaction Client: Date Effet - Date Liv Demandée
    if 'Date Effet' in calc_data.columns and 'Date Liv Demander' in calc_data.columns:
        calc_data['Satisfaction_Client_Jours'] = (calc_data['Date Effet'] - calc_data['Date Liv Demander']).dt.days
        calc_data['Client_Satisfait'] = (calc_data['Satisfaction_Client_Jours'] <= 0).astype(int)
    else:
        calc_data['Satisfaction_Client_Jours'] = np.nan
        calc_data['Client_Satisfait'] = 0
        
        # Ajustement: Liv Ajustée existe ou non
    if 'Date Liv. Ajuster' in calc_data.columns:
        calc_data['Est_Ajuste'] = calc_data['Date Liv. Ajuster'].notna().astype(int)
    else:
        calc_data['Est_Ajuste'] = 0
        
        # Respect du planning: Satisfaction Client ≤ 0 (livrée à temps ou en avance)
    calc_data['Respecte_Planning'] = (calc_data['Satisfaction_Client_Jours'] <= 0).astype(int)
        
        # ==================== KPI CARDS ====================
    st.subheader("KPIs Clés - Analyse Délais & Cycles")
        
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        
    nb_lots = len(calc_data)
    cycle_moyen = calc_data['Cycle_Prevu_Jours'].mean()
    retard_total_moyen = calc_data['Retard_Total_Jours'].mean()
    satisfaction_client_pct = (calc_data['Client_Satisfait'].sum() / len(calc_data) * 100) if len(calc_data) > 0 else 0
        
    with kpi_col1:
        render_metric_card(
                "Nombre de Lots",
                f"{nb_lots}",
                color_gradient="cyan-to-green"
            )
        
    with kpi_col2:
        render_metric_card(
                "Cycle Moyen",
                f"{cycle_moyen:.1f} j" if pd.notna(cycle_moyen) else "N/A",
                color_gradient="blue-to-purple"
            )
        
    with kpi_col3:
        render_metric_card(
                "Retard Total Moyen",
                f"{retard_total_moyen:.1f} j" if pd.notna(retard_total_moyen) else "N/A",
                color_gradient="amber-to-orange"
            )
        
    with kpi_col4:
        render_metric_card(
                "Satisfaction Client",
                f"{satisfaction_client_pct:.1f}%",
                color_gradient="cyan-to-green"
            )
        
        # KPI Explanations
    st.markdown("""
        <div class='explanation-text'>
        <strong>📊 Explications des KPIs:</strong><br>
        • <strong>Nombre de Lots:</strong> Nombre total de lots analysés après filtrage par espèce, article et lot<br>
        • <strong>Cycle Moyen:</strong> Durée moyenne entre le Démarrage Réel (DS) et la Livraison Prévue<br>
        • <strong>Retard Total Moyen:</strong> Somme moyenne de tous les retards (Semis + Élevage + Ajusté-Réel)<br>
        • <strong>Satisfaction Client:</strong> Pourcentage de lots livrés à la date demandée ou en avance (Date Effet ≤ Date Demandée)
        </div>
        """, unsafe_allow_html=True)
        
    st.divider()
        
        # ==================== INDICE DE SATISFACTION CLIENT ====================
    client_satisfaction_pct = satisfaction_client_pct
        
    st.markdown(f"""
        <div style='
            background: linear-gradient(135deg, rgba(16, 185, 129, 0.2) 0%, rgba(5, 150, 105, 0.1) 100%);
            border: 2px solid rgba(16, 185, 129, 0.4);
            padding: 24px;
            border-radius: 12px;
            text-align: center;
            animation: floatCard 4s ease-in-out infinite;
        '>
            <p style='color: #a7f3d0; font-size: 0.95rem; margin: 0; font-weight: 600;'>INDICE DE SATISFACTION CLIENT</p>
            <p style='color: #10b981; font-size: 3rem; margin: 10px 0 0 0; font-weight: 800;'>{client_satisfaction_pct:.1f}%</p>
            <p style='color: #86efac; font-size: 0.9rem; margin: 6px 0 0 0;'>{int(calc_data['Client_Satisfait'].sum())} / {len(calc_data)} lots livrés à la date demandée ou en avance</p>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("""
        <div class='explanation-text'>
        <strong>🎯 Calcul de la Satisfaction Client:</strong><br>
        Basé sur la comparaison entre <strong>Date Livraison Réelle (Date Effet)</strong> et <strong>Date Livraison Demandée</strong>.
        Un client est satisfait si la livraison réelle ≤ livraison demandée (0 retard ou avance).<br>
        L'Indice mesure le <strong>pourcentage de lots livrés sans retard</strong>. Un lot est considéré comme conforme si son <strong>Retard Total (Liv Ajustée - Liv Demandée) ≤ 0 jours</strong>. 
        Cet indicateur est <strong>critique pour évaluer la performance globale</strong> de votre chaîne d'approvisionnement et la satisfaction client.
        </div>
        """, unsafe_allow_html=True)
        
    st.divider()
        
        # ==================== EXPLICATIONS DÉTAILLÉES DES KPIs ====================
    with st.expander("📚 Guide Complet des KPIs et Métriques", expanded=False):
        st.markdown("""
            ## Comprendre Page 11: Analyse Délais & Cycle de Production
            
            ### 🎯 Les 4 KPIs Principaux
            
            **1. Nombre de Lots**
            - Représente le nombre total de lots analysés après filtrage
            - Permet de contextualiser les statistiques (petit n vs grand n)
            - Important pour la fiabilité statistique
            
            **2. Cycle Moyen (DS Réel → Liv Prévue)**
            - Durée moyenne en jours du cycle de production complet
            - Calcul: (Date Livraison Prévue - Date Démarrage Réel)
            - Mesure l'efficacité opérationnelle de votre processus de production
            - Benchmark: Comparez avec vos standards de performance
            
            **3. Retard Semis Moyen (Phase 1)**
            - Délai entre la Date de Livraison DEMANDÉE et la Date de Livraison PRÉVUE
            - Calcul: (Date Liv Prévue - Date Liv Demandée)
            - Représente les imprécisions de la phase initiale de planning/semis
            - Négatif = livraison anticipée / Positif = délai anticipé
            - Importance: Critique pour l'optimisation des ressources initiales
            
            **4. Retard Élevage Moyen (Phase 2)**
            - Délai supplémentaire entre la Date de Livraison PRÉVUE et l'AJUSTÉE
            - Calcul: (Date Liv Ajustée - Date Liv Prévue)
            - Capture les imprévus et ajustements pendant la phase d'élevage
            - Négatif = rattrapage anticipé / Positif = délai supplémentaire
            - Importance: Indicateur de contrôle de qualité et stabilité opérationnelle
            
            ### 📊 Indicateurs Secondaires
            
            **INDICE DE RESPECT DU PLANNING**
            - Pourcentage de lots livrés sans retard total
            - Calcul: (Nombre de lots avec Retard Total ≤ 0) / (Total lots) × 100
            - Formule exacte: (Date Liv Ajustée OU Prévue - Date Liv Demandée) ≤ 0
            - Performance:
              - ✅ 80%+ : Excellente (respect exceptionnaire)
              - ⚠️ 60-80% : Acceptable (amélioration recommandée)
              - ❌ <60% : Critique (action urgente requise)
            
            **Lots Ajustés**
            - Nombre et pourcentage de lots ayant eu une Date Liv Ajust��e
            - Indicateur: Plus de 20% d'ajustements = problèmes de prévision
            
            ### 🔄 Les 3 Types de Retards
            
            **Retard Semis (Demandé → Prévu)**
            - Erreurs/délais dans la phase de demande et prévision initiale
            - Sources possibles: Mauvaise estimation client, délais admin, planning initial
            - Amélioration: Meilleur dialogue avec clients, méthodes de prévision
            
            **Retard Élevage (Prévu → Ajusté)**
            - Imprévus pendant la croissance des plants
            - Sources possibles: Qualité des plants, conditions climatiques, problèmes sanitaires
            - Amélioration: Contrôle qualité, conditions optimales, prévention sanitaire
            
            **Retard Total (Demandé → Ajusté/Réel)**
            - Cumul des deux retards précédents
            - Représente l'expérience client réelle
            - C'est ce qui compte vraiment pour la satisfaction client
            
            ### 📈 Statistiques par Métrique
            
            Pour chaque KPI, vous voyez:
            - **Moyenne**: Tendance centrale (attention aux outliers!)
            - **Médiane**: Valeur du milieu (plus robuste que la moyenne)
            - **Min/Max**: Amplitude de variation (écart entre le meilleur et le pire cas)
            
            ### 🎓 Interprétation Avancée
            
            **Cycle Prévu vs Cycle Réel (Visualisations - Analyse Graphique)**
            - **Cycle Prévu**: (Date Livraison Prévue - Date Démarrage Réel) = la durée planifiée
            - **Cycle Réel**: (Date Livraison Ajustée OU Prévue - Date Démarrage Réel) = la durée réelle exécutée
            - **La différence** révèle votre capacité de prévision et de contrôle opérationnel
            - **Petit écart** = bonne maîtrise du processus et prévisions précises
            - **Grand écart** = problèmes de stabilité, imprévus importants ou prévisions inexactes
            - **Comparaison**: Plus le cycle réel se rapproche du prévu, meilleure est votre performance
            
            **Corrélations Importantes**
            - Retard Semis élevé + Retard Élevage faible = Bonne adaptation
            - Retard Semis faible + Retard Élevage élevé = Problèmes opérationnels
            - Les deux élevés = Système instable, action globale requise
            
            ### ✅ Bonnes Pratiques
            
            1. **Analysez les TOP 10** pour identifier les meilleures pratiques
            2. **Comparez Cycle Prévu vs Cycle Réel** pour évaluer votre précision de planning
            3. **Segmentez par espèce/article** pour identifier les problèmes spécifiques
            4. **Suivez les tendances** sur plusieurs périodes (tendance = prédictif)
            5. **Agissez sur le Retard Semis d'abord** (phase amont, impact majeur)
            """)
        
    st.divider()
        
        # ==================== STOCKAGE POUR PARTIES SUIVANTES ====================
    st.session_state['delivery_calc_data'] = calc_data
    st.session_state['delivery_selected_species'] = selected_species
    st.session_state['delivery_selected_article'] = selected_article
    st.session_state['delivery_selected_lot'] = selected_lot_display
    st.session_state['delivery_is_single_lot'] = is_single_lot
        
        # ==================== PARTIE 2: TABLEAU COMPARATIF ====================
    st.subheader("Tableau Comparatif - Dates & Cycles")
        
        # Créer une copie pour affichage
    table_data = calc_data.copy()
        
        # Construire le tableau d'affichage
    display_cols_order = [
            'LOT', 'Article',
            'DS. Réel', 'Date Liv Demander', 'Date Liv Prevu', 'Date Liv. Ajuster', 'Date Effet',
            'Cycle_Prevu_Jours', 'Cycle_Reel_Jours', 'Retard_Semis_Jours', 'Retard_Élevage_Jours', 
            'Retard_Ajuste_Reel_Jours', 'Retard_Total_Jours', 'Satisfaction_Client_Jours', 'Statut'
        ]
        
        # Formater les dates
    for date_col in ['DS. Réel', 'Date Liv Demander', 'Date Liv Prevu', 'Date Liv. Ajuster', 'Date Effet']:
        if date_col in table_data.columns:
            table_data[date_col] = table_data[date_col].dt.strftime('%d/%m/%Y')
        
        # Ajouter indicateur visuel de statut (basé sur Satisfaction Client)
    def get_status_icon(satisfaction_jours):
        if pd.isna(satisfaction_jours):
            return "⚪"
        elif satisfaction_jours <= 0:
            return "✅"
        elif satisfaction_jours <= 3:
             return "⚠️"
        else:
            return "❌"
        
    table_data['Statut'] = calc_data['Satisfaction_Client_Jours'].apply(get_status_icon)
        
        # Afficher le tableau avec les colonnes disponibles
    display_cols_available = [col for col in display_cols_order if col in table_data.columns]
    table_display = table_data[display_cols_available].copy()
        
    st.dataframe(
    table_display,
    use_container_width=True,
    height=400
)

# 📥 Téléchargement Excel
    table_display_xlsx = dataframe_to_xlsx_bytes(table_display)

    st.download_button(
    label="📥 Télécharger ce tableau — Excel",
    data=table_display_xlsx,
    file_name="tableau_detail.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_table_display"
)
        
    st.divider()
        
        # ==================== PARTIE 2B: GRAPHIQUES ====================
    st.subheader("Visualisations - Analyse Graphique")
        
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Distribution Cycles",
            "Retards par Lot",
            "Box Plot Cycles",
            "Comparaison Prévu/Réel",
            "Retard Ajusté-Réel"
        ])
        
    with tab1:
        if 'Cycle_Prevu_Jours' in calc_data.columns:
            fig_dist = px.histogram(
                    calc_data.dropna(subset=['Cycle_Prevu_Jours']),
                    x='Cycle_Prevu_Jours',
                    nbins=20,
                    title="Distribution du Cycle de Production (jours)",
                    labels={'Cycle_Prevu_Jours': 'Jours'},
                    color_discrete_sequence=['#06b6d4']
                )
            fig_dist.update_layout(
                    plot_bgcolor='rgba(15,23,42,0.5)',
                    paper_bgcolor='rgba(15,23,42,0.3)',
                    font=dict(color='#e2e8f0'),
                    title_font_size=16,
                    showlegend=False,
                    height=500
                )
            st.plotly_chart(fig_dist, use_container_width=True)
        else:
            st.warning("Colonne 'Cycle_Prevu_Jours' non disponible")
        
    with tab2:
        if 'Retard_Semis_Jours' in calc_data.columns and 'Retard_Élevage_Jours' in calc_data.columns and 'LOT' in calc_data.columns:
            retard_data = calc_data.dropna(subset=['Retard_Semis_Jours', 'Retard_Élevage_Jours']).reset_index(drop=True)
                
            if len(retard_data) > 0:
                fig_retard = go.Figure()
                    
                fig_retard.add_trace(go.Bar(
                        x=retard_data.index,
                        y=retard_data['Retard_Semis_Jours'],
                        name='Retard Semis (Demandé→Prévu)',
                        marker_color='#f59e0b',
                        text=retard_data['Retard_Semis_Jours'].round(1),
                        textposition='auto'
                    ))
                    
                fig_retard.add_trace(go.Bar(
                        x=retard_data.index,
                        y=retard_data['Retard_Élevage_Jours'],
                        name='Retard Élevage (Prévu→Ajusté)',
                        marker_color='#ef4444',
                        text=retard_data['Retard_Élevage_Jours'].round(1),
                        textposition='auto'
                    ))
                    
                fig_retard.update_layout(
                        title="Analyse des Deux Types de Retards par Lot",
                        barmode='group',
                        plot_bgcolor='rgba(15,23,42,0.5)',
                        paper_bgcolor='rgba(15,23,42,0.3)',
                        font=dict(color='#e2e8f0'),
                        title_font_size=16,
                        height=500,
                        hovermode='x unified',
                        xaxis_title='Lots',
                        yaxis_title='Retard (jours)',
                        legend=dict(
                            orientation='h',
                            yanchor='bottom',
                            y=1.02,
                            xanchor='right',
                            x=1
                        )
                    )
                fig_retard.add_hline(y=0, line_dash="dash", line_color="#10b981", annotation_text="Planning OK")
                st.plotly_chart(fig_retard, use_container_width=True)
        
    with tab3:
        if 'Cycle_Prevu_Jours' in calc_data.columns:
            cycle_clean = calc_data.dropna(subset=['Cycle_Prevu_Jours'])
                
            if len(cycle_clean) > 0:
                fig_box = go.Figure()
                    
                fig_box.add_trace(go.Box(
                        y=cycle_clean['Cycle_Prevu_Jours'],
                        name='Cycle Prévu',
                        marker_color='#06b6d4',
                        boxmean='sd'
                    ))
                    
                if 'Cycle_Reel_Jours' in calc_data.columns:
                    cycle_reel_clean = calc_data.dropna(subset=['Cycle_Reel_Jours'])
                    fig_box.add_trace(go.Box(
                            y=cycle_reel_clean['Cycle_Reel_Jours'],
                            name='Cycle Réel',
                            marker_color='#8b5cf6',
                            boxmean='sd'
                        ))
                    
                fig_box.update_layout(
                        title="Distribution Cycles - Box Plot Statistique",
                        plot_bgcolor='rgba(15,23,42,0.5)',
                        paper_bgcolor='rgba(15,23,42,0.3)',
                        font=dict(color='#e2e8f0'),
                        title_font_size=16,
                        height=500,
                        showlegend=True
                    )
                st.plotly_chart(fig_box, use_container_width=True)
        
    with tab4:
        if 'Cycle_Prevu_Jours' in calc_data.columns and 'Cycle_Reel_Jours' in calc_data.columns:
            compare_data = calc_data.dropna(subset=['Cycle_Prevu_Jours', 'Cycle_Reel_Jours']).reset_index(drop=True)
                
            if len(compare_data) > 0:
                fig_compare = go.Figure()
                    
                fig_compare.add_trace(go.Scatter(
                        x=compare_data.index,
                        y=compare_data['Cycle_Prevu_Jours'],
                        name='Cycle Prévu',
                        mode='lines+markers',
                        line=dict(color='#06b6d4', width=3),
                        marker=dict(size=8)
                    ))
                    
                fig_compare.add_trace(go.Scatter(
                        x=compare_data.index,
                        y=compare_data['Cycle_Reel_Jours'],
                        name='Cycle Réel',
                        mode='lines+markers',
                        line=dict(color='#8b5cf6', width=3),
                        marker=dict(size=8)
                    ))
                    
                fig_compare.update_layout(
                        title="Comparaison Cycle Prévu vs Cycle Réel",
                        plot_bgcolor='rgba(15,23,42,0.5)',
                        paper_bgcolor='rgba(15,23,42,0.3)',
                        font=dict(color='#e2e8f0'),
                        title_font_size=16,
                        height=500,
                        hovermode='x unified',
                        xaxis_title='Lots',
                        yaxis_title='Jours'
                    )
                st.plotly_chart(fig_compare, use_container_width=True)
        
    with tab5:
        if 'Retard_Ajuste_Reel_Jours' in calc_data.columns and 'Date Effet' in calc_data.columns:
            retard_ajuste_reel_clean = calc_data.dropna(subset=['Retard_Ajuste_Reel_Jours'])
                
            if len(retard_ajuste_reel_clean) > 0:
                fig_ajuste_reel = px.bar(
                        retard_ajuste_reel_clean.reset_index(),
                        x=retard_ajuste_reel_clean.index,
                        y='Retard_Ajuste_Reel_Jours',
                        color='Retard_Ajuste_Reel_Jours',
                        title="Retard entre Livraison Ajustée et Livraison Réelle (Date Effet)",
                        labels={'index': 'Lot', 'Retard_Ajuste_Reel_Jours': 'Retard (jours)'},
                        color_continuous_scale=['#22c55e', '#f59e0b', '#ef4444'],
                        range_color=[-5, 15]
                    )
                fig_ajuste_reel.update_layout(
                        plot_bgcolor='rgba(15,23,42,0.5)',
                        paper_bgcolor='rgba(15,23,42,0.3)',
                        font=dict(color='#e2e8f0'),
                        title_font_size=16,
                        height=500,
                        hovermode='x unified'
                    )
                fig_ajuste_reel.add_hline(y=0, line_dash="dash", line_color="#10b981", annotation_text="Liv Ajustée = Liv Réelle")
                st.plotly_chart(fig_ajuste_reel, use_container_width=True)
            else:
                st.info("Données de livraison réelle (Date Effet) non disponibles")
        else:
            st.info("Données de livraison réelle (Date Effet) non disponibles")
        
    st.divider()
        
        # ==================== PARTIE 3: STATISTIQUES DÉTAILLÉES ====================
    st.subheader("Statistiques Détaillées & Analyse")
        
    stat_col1, stat_col2 = st.columns(2)
        
    with stat_col1:
        st.markdown("**Cycle de Production (Prévu)**")
        if 'Cycle_Prevu_Jours' in calc_data.columns:
            cycle_prevu = calc_data['Cycle_Prevu_Jours'].dropna()
            if len(cycle_prevu) > 0:
                stats_prevu = {
                        "Moyenne": f"{cycle_prevu.mean():.1f} j",
                        "Médiane": f"{cycle_prevu.median():.1f} j",
                        "Minimum": f"{cycle_prevu.min():.0f} j",
                        "Maximum": f"{cycle_prevu.max():.0f} j",
                        "Écart-type": f"{cycle_prevu.std():.1f} j",
                        "Q1 (25%)": f"{cycle_prevu.quantile(0.25):.1f} j",
                        "Q3 (75%)": f"{cycle_prevu.quantile(0.75):.1f} j"
                    }
                for key, value in stats_prevu.items():
                    st.write(f"• {key}: **{value}**")
        
    with stat_col2:
        st.markdown("**Analyse des Retards**")
            
            # Retard Semis (Demandé → Prévu)
        if 'Retard_Semis_Jours' in calc_data.columns:
            retards_semis = calc_data['Retard_Semis_Jours'].dropna()
            if len(retards_semis) > 0:
                st.write("**Retard Semis (Phase de Demande → Prévision):**")
                stats_retard_semis = {
                        "Moyenne": f"{retards_semis.mean():.1f} j",
                        "Médiane": f"{retards_semis.median():.1f} j",
                        "Min": f"{retards_semis.min():.0f} j",
                        "Max": f"{retards_semis.max():.0f} j"
                    }
                for key, value in stats_retard_semis.items():
                    st.write(f"  • {key}: **{value}**")
            
            # Retard Élevage (Prévu → Ajusté)
        if 'Retard_Élevage_Jours' in calc_data.columns:
            retards_elevation = calc_data['Retard_Élevage_Jours'].dropna()
            if len(retards_elevation) > 0:
                st.write("**Retard Élevage (Phase de Prévision → Ajustement):**")
                stats_retard_elev = {
                        "Moyenne": f"{retards_elevation.mean():.1f} j",
                        "Médiane": f"{retards_elevation.median():.1f} j",
                        "Min": f"{retards_elevation.min():.0f} j",
                        "Max": f"{retards_elevation.max():.0f} j"
                    }
                for key, value in stats_retard_elev.items():
                    st.write(f"  • {key}: **{value}**")
            
            # Retard Total
        if 'Retard_Total_Jours' in calc_data.columns:
            retards_total = calc_data['Retard_Total_Jours'].dropna()
            if len(retards_total) > 0:
                st.write("**Retard Total Cumulé (Demandé → Ajusté):**")
                stats_retard_total = {
                        "Moyenne": f"{retards_total.mean():.1f} j",
                        "Médiane": f"{retards_total.median():.1f} j",
                        "Min": f"{retards_total.min():.0f} j",
                        "Max": f"{retards_total.max():.0f} j",
                        "Conformité": f"{(retards_total <= 0).sum()} / {len(retards_total)}"
                    }
                for key, value in stats_retard_total.items():
                    st.write(f"  • {key}: **{value}**")
            
            # Retard Ajusté-Réel (Ajusté → Livraison Réelle)
        if 'Retard_Ajuste_Reel_Jours' in calc_data.columns:
            retards_ajuste_reel = calc_data['Retard_Ajuste_Reel_Jours'].dropna()
            if len(retards_ajuste_reel) > 0:
                st.write("**Retard Ajusté-Réel (Phase d'Ajustement → Livraison Réelle):**")
                stats_retard_ajuste_reel = {
                        "Moyenne": f"{retards_ajuste_reel.mean():.1f} j",
                        "Médiane": f"{retards_ajuste_reel.median():.1f} j",
                        "Min": f"{retards_ajuste_reel.min():.0f} j",
                        "Max": f"{retards_ajuste_reel.max():.0f} j"
                    }
                for key, value in stats_retard_ajuste_reel.items():
                    st.write(f"  • {key}: **{value}**")
        
    st.divider()
        
        # ==================== PARTIE 3B: TOP 10 ====================
    st.subheader("Analyse Comparative - Top 10")
        
    top_col1, top_col2 = st.columns(2)
        
    with top_col1:
        st.markdown("**✅ Meilleurs Lots (Cycle le Plus Court)**")
        if 'Cycle_Prevu_Jours' in calc_data.columns:
            best_lots = calc_data.nsmallest(10, 'Cycle_Prevu_Jours')[
                    ['LOT', 'Article', 'Cycle_Prevu_Jours', 'Retard_Semis_Jours', 'Retard_Élevage_Jours', 'Respecte_Planning']
                ].copy()
            best_lots.columns = ['LOT', 'Article', 'Cycle (j)', 'Retard Semis (j)', 'Retard Élev (j)', 'OK']
            st.dataframe(best_lots, use_container_width=True, hide_index=True)
            best_lots_xlsx = dataframe_to_xlsx_bytes(best_lots)
            
            st.download_button(
                label="📥 Télécharger ce tableau — Excel",
                data=best_lots_xlsx,
                file_name="best_lots.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_table_display2"
            )
        
    with top_col2:
        st.markdown("**❌ Pires Lots (Cycle le Plus Long)**")
        if 'Cycle_Prevu_Jours' in calc_data.columns:
            worst_lots = calc_data.nlargest(10, 'Cycle_Prevu_Jours')[
                    ['LOT', 'Article', 'Cycle_Prevu_Jours', 'Retard_Semis_Jours', 'Retard_Élevage_Jours', 'Respecte_Planning']
                ].copy()
            worst_lots.columns = ['LOT', 'Article', 'Cycle (j)', 'Retard Semis (j)', 'Retard Élev (j)', 'OK']
            st.dataframe(worst_lots, use_container_width=True, hide_index=True)
            worst_lots_xlsx = dataframe_to_xlsx_bytes(worst_lots)
                        
            st.download_button(
                            label="📥 Télécharger ce tableau — Excel",
                            data=worst_lots_xlsx,
                            file_name="worst_lots.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_table_display3"
                        )
        
    st.divider()
        
        # ==================== PARTIE 4: TIMELINE & RÉSUMÉ ====================
    st.subheader("Timeline - Progression Chronologique")
        
        # Créer une vue timeline simplifiée
    if 'LOT' in calc_data.columns and 'DS. Réel' in calc_data.columns:
        timeline_cols = [
                'LOT', 'DS. Réel', 'Date Liv Demander', 'Date Liv Prevu', 'Date Liv. Ajuster', 'Date Effet',
                'Cycle_Prevu_Jours', 'Retard_Semis_Jours', 'Retard_Élevage_Jours', 'Retard_Ajuste_Reel_Jours', 'Retard_Total_Jours'
            ]
            # Sélectionner uniquement les colonnes qui existent
        timeline_cols = [col for col in timeline_cols if col in calc_data.columns]
        timeline_data = calc_data[timeline_cols].dropna(subset=['DS. Réel']).sort_values('DS. Réel').head(20).copy()
            
        timeline_html = '<div style="overflow-x: auto;">'
        timeline_html += '<table style="width:100%; border-collapse: collapse; color: #e2e8f0;">'
        timeline_html += '<tr style="background: rgba(6, 182, 212, 0.2); border-bottom: 2px solid #06b6d4;">'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Lot</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">DS Réel</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Liv Demandée</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Liv Prévue</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Liv Ajustée</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Liv Réelle</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Cycle (j)</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Retard Semis (j)</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Retard Élev (j)</th>'
        timeline_html += '<th style="padding: 10px; text-align: left; border-right: 1px solid #475569;">Retard Ajusté-Réel (j)</th>'
        timeline_html += '<th style="padding: 10px; text-align: left;">Retard Total (j)</th>'
        timeline_html += '</tr>'
            
        for idx, row in timeline_data.iterrows():
            timeline_html += '<tr style="border-bottom: 1px solid #475569;">'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{row["LOT"]}</td>'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{row["DS. Réel"].strftime("%d/%m/%Y") if pd.notna(row["DS. Réel"]) else ""}</td>'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{row["Date Liv Demander"].strftime("%d/%m/%Y") if pd.notna(row["Date Liv Demander"]) else ""}</td>'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{row["Date Liv Prevu"].strftime("%d/%m/%Y") if pd.notna(row["Date Liv Prevu"]) else ""}</td>'
            liv_ajuster_text = row["Date Liv. Ajuster"].strftime("%d/%m/%Y") if pd.notna(row["Date Liv. Ajuster"]) else "��"
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{liv_ajuster_text}</td>'
                
                # Date Effet (Livraison Réelle) - vérifier si la colonne existe
            if "Date Effet" in timeline_data.columns:
                date_effet_text = row["Date Effet"].strftime("%d/%m/%Y") if pd.notna(row["Date Effet"]) else "—"
            else:
                date_effet_text = "—"
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569;">{date_effet_text}</td>'
                
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569; text-align: center;">{row["Cycle_Prevu_Jours"]:.0f}</td>'
                
                # Retard Semis
            retard_semis_text = f"{row['Retard_Semis_Jours']:.0f}" if pd.notna(row["Retard_Semis_Jours"]) else "—"
            retard_semis_color = '#10b981' if pd.isna(row["Retard_Semis_Jours"]) or row["Retard_Semis_Jours"] <= 0 else '#f59e0b' if row["Retard_Semis_Jours"] <= 3 else '#ef4444'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569; text-align: center; color: {retard_semis_color}; font-weight: bold;">{retard_semis_text}</td>'
                
                # Retard Élevage
            retard_elev_text = f"{row['Retard_Élevage_Jours']:.0f}" if pd.notna(row["Retard_Élevage_Jours"]) else "—"
            retard_elev_color = '#10b981' if pd.isna(row["Retard_Élevage_Jours"]) or row["Retard_Élevage_Jours"] <= 0 else '#f59e0b' if row["Retard_Élevage_Jours"] <= 3 else '#ef4444'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569; text-align: center; color: {retard_elev_color}; font-weight: bold;">{retard_elev_text}</td>'
                
                # Retard Ajusté-Réel - vérifier si la colonne existe
            if "Retard_Ajuste_Reel_Jours" in timeline_data.columns:
                retard_ajuste_reel_text = f"{row['Retard_Ajuste_Reel_Jours']:.0f}" if pd.notna(row["Retard_Ajuste_Reel_Jours"]) else "—"
                retard_ajuste_reel_color = '#10b981' if pd.isna(row["Retard_Ajuste_Reel_Jours"]) or row["Retard_Ajuste_Reel_Jours"] <= 0 else '#f59e0b' if row["Retard_Ajuste_Reel_Jours"] <= 3 else '#ef4444'
            else:
                retard_ajuste_reel_text = "—"
                retard_ajuste_reel_color = '#475569'
            timeline_html += f'<td style="padding: 10px; border-right: 1px solid #475569; text-align: center; color: {retard_ajuste_reel_color}; font-weight: bold;">{retard_ajuste_reel_text}</td>'
                
                # Retard Total
            retard_total_text = f"{row['Retard_Total_Jours']:.0f}" if pd.notna(row["Retard_Total_Jours"]) else "—"
            retard_total_color = '#10b981' if pd.isna(row["Retard_Total_Jours"]) or row["Retard_Total_Jours"] <= 0 else '#f59e0b' if row["Retard_Total_Jours"] <= 3 else '#ef4444'
            timeline_html += f'<td style="padding: 10px; text-align: center; color: {retard_total_color}; font-weight: bold;">{retard_total_text}</td>'
            timeline_html += '</tr>'
            
        timeline_html += '</table></div>'
        st.markdown(timeline_html, unsafe_allow_html=True)
        timeline_xlsx = dataframe_to_xlsx_bytes(timeline_data)

        st.download_button(
    label="📥 Télécharger ce tableau — Excel",
    data=timeline_xlsx,
    file_name="timeline_progression.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_table_display4"
)
        
    st.divider()
        
        # ==================== PARTIE 4B: RÉSUMÉ ANALYTIQUE AUTOMATIQUE ====================
    st.subheader("Résumé Analytique & Recommandations")
        
        # Construire le résumé
    summary_sections = []
        
        # Section 1: Performance générale
    if 'Cycle_Prevu_Jours' in calc_data.columns:
        cycle_mean = calc_data['Cycle_Prevu_Jours'].mean()
        cycle_std = calc_data['Cycle_Prevu_Jours'].std()
        summary_sections.append(f"Le cycle de production moyen est de **{cycle_mean:.1f} jours** avec un écart-type de **{cycle_std:.1f} jours**.")
        
        # Section 2: Respect du planning
    if 'Respecte_Planning' in calc_data.columns:
        respect_pct_summary = (calc_data['Respecte_Planning'].sum() / len(calc_data) * 100)
        respect_count = int(calc_data['Respecte_Planning'].sum())
        if respect_pct_summary >= 80:
            summary_sections.append(f"✅ **{respect_pct_summary:.0f}%** des lots ({respect_count}/{len(calc_data)}) sont livrés en conformité avec le planning. Excellente performance!")
        elif respect_pct_summary >= 60:
            summary_sections.append(f"⚠️ **{respect_pct_summary:.0f}%** des lots ({respect_count}/{len(calc_data)}) sont livrés en conformité. Des améliorations sont recommandées.")
        else:
            summary_sections.append(f"❌ Seulement **{respect_pct_summary:.0f}%** des lots sont en conformité. Action corrective urgente requise.")
        
        # Section 3: Ajustements
    if 'Est_Ajuste' in calc_data.columns:
        ajuste_pct = (calc_data['Est_Ajuste'].sum() / len(calc_data) * 100)
        ajuste_count = int(calc_data['Est_Ajuste'].sum())
        summary_sections.append(f"**{ajuste_pct:.0f}%** des lots ({ajuste_count}/{len(calc_data)}) ont nécessité un ajustement de date de livraison.")
        
        # Section 4: Retards détaillés
    if 'Retard_Semis_Jours' in calc_data.columns:
        retard_semis_mean = calc_data['Retard_Semis_Jours'].mean()
        summary_sections.append(f"Retard Semis moyen: **{retard_semis_mean:.1f} jours** (phase demande → prévision)")
        
    if 'Retard_Élevage_Jours' in calc_data.columns:
        retard_elev_mean = calc_data['Retard_Élevage_Jours'].mean()
        summary_sections.append(f"Retard Élevage moyen: **{retard_elev_mean:.1f} jours** (phase prévision → ajustement)")
        
    if 'Retard_Ajuste_Reel_Jours' in calc_data.columns:
        retard_ajuste_reel_mean = calc_data['Retard_Ajuste_Reel_Jours'].mean()
        summary_sections.append(f"Retard Ajusté-Réel moyen: **{retard_ajuste_reel_mean:.1f} jours** (écart entre livraison ajustée et livraison réelle)")
        
    if 'Retard_Total_Jours' in calc_data.columns:
        retard_total_mean = calc_data['Retard_Total_Jours'].mean()
        if retard_total_mean > 0:
            summary_sections.append(f"Retard Total moyen: **{retard_total_mean:.1f} jours**. Une optimisation du planning est recommandée.")
        else:
            summary_sections.append(f"Retard Total moyen: **{retard_total_mean:.1f} jours** (avance). Bon respect du planning!")
        
        # Afficher le résumé
    summary_text = "\n\n".join(summary_sections)
    st.markdown(f"""
        <div class='recommendation-box' style='
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.15) 0%, rgba(168, 85, 247, 0.1) 100%);
            border: 2px solid rgba(59, 130, 246, 0.3);
            padding: 20px;
            border-radius: 12px;
            margin-top: 20px;
        '>
        {summary_text}
        </div>
        """, unsafe_allow_html=True)
        
    st.divider()
        
        # Section: Recommandations spécifiques
    st.subheader("Recommandations Actionnables")
        
    recommendations = []
        
        # Recommandation 1: Cycles trop longs
    if 'Cycle_Prevu_Jours' in calc_data.columns:
        long_cycles = len(calc_data[calc_data['Cycle_Prevu_Jours'] > calc_data['Cycle_Prevu_Jours'].quantile(0.75)])
        if long_cycles > 0:
            recommendations.append(f"• {long_cycles} lot(s) ont un cycle > Q3. Analyser les goulets d'étranglement et optimiser le processus.")
        
        # Recommandation 2: Retards semis fréquents
    if 'Retard_Semis_Jours' in calc_data.columns:
        late_lots_semis = len(calc_data[calc_data['Retard_Semis_Jours'] > 3])
        if late_lots_semis > 0:
            recommendations.append(f"• {late_lots_semis} lot(s) ont un Retard Semis > 3 jours. Améliorer la phase de demande/prévision.")
        
        # Recommandation 3: Retards élevage fréquents
    if 'Retard_Élevage_Jours' in calc_data.columns:
        late_lots_elev = len(calc_data[calc_data['Retard_Élevage_Jours'] > 3])
        if late_lots_elev > 0:
            recommendations.append(f"• {late_lots_elev} lot(s) ont un Retard Élevage > 3 jours. Optimiser les ressources d'élevage et la qualité.")
        
        # Recommandation 3: Ajustements fréquents
    if 'Est_Ajuste' in calc_data.columns:
        adjusted = calc_data['Est_Ajuste'].sum()
        if adjusted / len(calc_data) > 0.3:
            recommendations.append(f"• {int(adjusted)} lot(s) ajustés ({adjusted/len(calc_data)*100:.0f}%). Revoir la stratégie de planification.")
        
    if recommendations:
        for rec in recommendations:
            st.write(rec)
    else:
        st.success("✅ Aucun problème majeur détecté. La chaîne d'approvisionnement fonctionne bien.")