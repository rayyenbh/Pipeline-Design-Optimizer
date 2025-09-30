import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import numpy as np
import matplotlib
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import os, sys
import tempfile
import threading


# Fix scaling issue on Windows (DPI awareness)
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1 and later
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # Windows 7
    except Exception:
        pass

    

# Enhanced Pipeline Report Generator with Professional Styling
# ------------------------------------------------------------------
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak, FrameBreak, KeepTogether, HRFlowable
)
from reportlab.lib.units import inch, cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime

# Custom color palette for ReportLab
CORPORATE_BLUE = colors.Color(0.1, 0.2, 0.4)  # Dark blue
ACCENT_BLUE = colors.Color(0.2, 0.4, 0.7)     # Medium blue
LIGHT_BLUE = colors.Color(0.9, 0.95, 1.0)     # Very light blue
GRAY_HEADER = colors.Color(0.3, 0.3, 0.3)     # Dark gray
LIGHT_GRAY = colors.Color(0.95, 0.95, 0.95)   # Light gray

# Matplotlib-compatible colors (RGB tuples)
MPL_CORPORATE_BLUE = (0.1, 0.2, 0.4)
MPL_ACCENT_BLUE = (0.2, 0.4, 0.7)
MPL_GRAY_HEADER = (0.3, 0.3, 0.3)





def resource_path(relative_path: str):
    """ Get absolute path to resource for dev and for PyInstaller """
    if getattr(sys, 'frozen', False):  # exe mode
        return os.path.join(os.path.dirname(sys.executable), relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)


def _cheapest_material(prices: dict) -> str:
    """Return the material key that has the lowest total_cost."""
    if not prices:
        return "N/A"
    return min(prices, key=lambda m: prices[m]["total_cost"])


def create_custom_styles():
    """Create professional custom styles"""
    styles = getSampleStyleSheet()
    
    # Check if custom styles already exist, if not add them
    custom_style_names = ['MainTitle', 'SubTitle', 'SectionHeader', 'BodyText', 'KeyValue']
    
    for style_name in custom_style_names:
        if style_name in styles:
            continue  # Skip if already exists
    
    # Main title style
    if 'MainTitle' not in styles:
        styles.add(ParagraphStyle(
            name='MainTitle',
            parent=styles['Title'],
            fontSize=24,
            textColor=CORPORATE_BLUE,
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
    
    # Subtitle style
    if 'SubTitle' not in styles:
        styles.add(ParagraphStyle(
            name='SubTitle',
            parent=styles['Heading2'],
            fontSize=16,
            textColor=ACCENT_BLUE,
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))
    
    # Section heading style
    if 'SectionHeader' not in styles:
        styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=CORPORATE_BLUE,
            spaceBefore=15,
            spaceAfter=8,
            fontName='Helvetica-Bold',
            borderWidth=0,
            borderColor=ACCENT_BLUE,
            borderPadding=5,
            backColor=LIGHT_BLUE
        ))
    
    # Body text with better spacing
    if 'BodyText' not in styles:
        styles.add(ParagraphStyle(
            name='BodyText',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            leading=14,
            alignment=TA_JUSTIFY
        ))
    
    # Key-value pair style
    if 'KeyValue' not in styles:
        styles.add(ParagraphStyle(
            name='KeyValue',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=4,
            leftIndent=10,
            leading=13
        ))
    
    return styles

def create_side_by_side_plots(img_folder):
    """Create side-by-side plots layout"""
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec
    import matplotlib.image as mpimg
    
    # Check if individual plots exist
    vel_img_path = img_folder / "velocity_vs_diameter.png"
    dp_img_path = img_folder / "pressure_drop_vs_diameter.png"
    
    if vel_img_path.exists() and dp_img_path.exists():
        # Create a figure with side-by-side subplots
        fig = plt.figure(figsize=(12, 5))
        gs = GridSpec(1, 2, figure=fig, hspace=0.1, wspace=0.3)
        
        # Load and display the individual plots
        ax1 = fig.add_subplot(gs[0, 0])
        ax2 = fig.add_subplot(gs[0, 1])
        
        # Load the saved images
        vel_img = mpimg.imread(vel_img_path)
        dp_img = mpimg.imread(dp_img_path)
        
        # Display the images
        ax1.imshow(vel_img)
        ax1.axis('off')
        ax1.set_title('Velocity vs Diameter', fontweight='bold', color=MPL_CORPORATE_BLUE, pad=10)
        
        ax2.imshow(dp_img)
        ax2.axis('off')
        ax2.set_title('Pressure Drop vs Diameter', fontweight='bold', color=MPL_CORPORATE_BLUE, pad=10)
        
        # Adjust layout (remove tight_layout to avoid warning with image subplots)
        plt.subplots_adjust(left=0.05, right=0.95, top=0.85, bottom=0.1, wspace=0.3)
        
        # Save combined plot
        combined_img = img_folder / "combined_plots.png"
        plt.savefig(combined_img, dpi=200, bbox_inches="tight", facecolor='white')
        plt.close()
        
        return combined_img
    else:
        # If individual plots don't exist, create a placeholder
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        ax1.text(0.5, 0.5, 'Velocity vs Diameter\n(Plot not available)', 
                ha='center', va='center', transform=ax1.transAxes,
                fontsize=12, color=MPL_GRAY_HEADER)
        ax1.set_title('Velocity vs Diameter', fontweight='bold', color=MPL_CORPORATE_BLUE)
        
        ax2.text(0.5, 0.5, 'Pressure Drop vs Diameter\n(Plot not available)', 
                ha='center', va='center', transform=ax2.transAxes,
                fontsize=12, color=MPL_GRAY_HEADER)
        ax2.set_title('Pressure Drop vs Diameter', fontweight='bold', color=MPL_CORPORATE_BLUE)
        
        # Style the axes
        for ax in [ax1, ax2]:
            ax.grid(True, alpha=0.3)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color(MPL_GRAY_HEADER)
            ax.spines['bottom'].set_color(MPL_GRAY_HEADER)
        
        plt.subplots_adjust(left=0.1, right=0.9, top=0.85, bottom=0.15, wspace=0.3)
        
        # Save combined plot
        combined_img = img_folder / "combined_plots.png"
        plt.savefig(combined_img, dpi=200, bbox_inches="tight", facecolor='white')
        plt.close()
        
        return combined_img

def build_pipeline_report(inputs, compatible, plots, thickness_results, results, fittings, prices, file_path: Path):
    # ------------------------------------------------------------------
    # 0. Setup and save plots
    # ------------------------------------------------------------------
    img_folder = Path(__file__).parent / "report_imgs"
    img_folder.mkdir(exist_ok=True)

    # Create side-by-side plots
    combined_plots_img = create_side_by_side_plots(img_folder)

    # ------------------------------------------------------------------
    # 1. Document setup with better margins
    # ------------------------------------------------------------------
    doc = SimpleDocTemplate(
        str(file_path), 
        pagesize=A4, 
        topMargin=0.8*inch,
        bottomMargin=0.8*inch,
        leftMargin=0.8*inch,
        rightMargin=0.8*inch
    )

    # Enhanced header with company info
    def header_footer(canvas, doc):
        canvas.saveState()
        
        # Header
        logo_path = Path(__file__).with_name("eppm.png")
        if logo_path.exists():
            # Left logo
            canvas.drawImage(str(logo_path), 40, A4[1]-60, width=40, height=40, preserveAspectRatio=True)
            # Right logo
            canvas.drawImage(str(logo_path), A4[0]-80, A4[1]-60, width=40, height=40, preserveAspectRatio=True)
        
        # Header line
        canvas.setStrokeColor(ACCENT_BLUE)
        canvas.setLineWidth(2)
        canvas.line(40, A4[1]-70, A4[0]-40, A4[1]-70)
        
        # Footer
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(GRAY_HEADER)
        
        # Date and page number
        date_str = datetime.now().strftime("%B %d, %Y")
        canvas.drawString(40, 40, f"Generated on: {date_str}")
        canvas.drawRightString(A4[0]-40, 40, f"Page {doc.page}")
        
        # Footer line
        canvas.setStrokeColor(ACCENT_BLUE)
        canvas.setLineWidth(1)
        canvas.line(40, 50, A4[0]-40, 50)
        
        canvas.restoreState()

    # ------------------------------------------------------------------
    # 2. Content
    # ------------------------------------------------------------------
    story = []
    styles = create_custom_styles()

    # Title section with better formatting
    story.append(Spacer(1, 20))
    story.append(Paragraph("Pipeline Design Optimization Report", styles["MainTitle"]))
    story.append(Paragraph(inputs.get("project_name", "Untitled Project"), styles["SubTitle"]))
    
    # Add decorative line
    story.append(HRFlowable(width="100%", thickness=2, color=ACCENT_BLUE, spaceBefore=10, spaceAfter=20))

    # Executive Summary (new section)
    story.append(Paragraph("Executive Summary", styles["SectionHeader"]))
    summary_text = f"""
    This report presents the results of pipeline design optimization for the {inputs.get('project_name', 'project')} 
    with a flowrate of {inputs.get('flowrate_m3h', '')} m³/h over {inputs.get('pipe_length_m', '')} meters. 
    The analysis determined an optimal diameter of {plots['chosen_d']:.4f} m based on velocity and pressure drop constraints, 
    resulting in a velocity of {results['V']:.3f} m/s and total pressure drop of {(results['dp_linear'] + results['dp_singular'])/1e5:.4f} bar.
    """
    story.append(Paragraph(summary_text, styles["BodyText"]))
    story.append(Spacer(1, 15))

    # User Input section 
    story.append(Paragraph("Design Parameters", styles["SectionHeader"]))
    
    # Create input table
    input_data = [
        ["Parameter", "Value", "Unit"],
        ["Project Name", inputs.get('project_name', ''), ""],
        ["Flowrate", f"{inputs.get('flowrate_m3h', '')}", "m³/h"],
        ["Pipe Length", f"{inputs.get('pipe_length_m', '')}", "m"],
        ["Phase", inputs.get('phase', ''), ""],
        ["Fluid", inputs.get('fluid', ''), ""],
        ["Maximum Velocity", f"{inputs.get('max_velocity_mps', '')}", "m/s"],
        ["Temperature", f"{inputs.get('temperature_c', '')}", "°C"],
        ["Operating Pressure", f"{inputs.get('operating_pressure_bar', '')}", "bar"],
        ["Design Pressure", f"{inputs.get('design_pressure_bar', '')}", "bar"],
        ["Location Type", inputs.get('location_type', ''), ""],
    ]
    
    input_table = Table(input_data, colWidths=[2.5*inch, 1.5*inch, 0.8*inch])
    input_table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0,0), (-1,0), CORPORATE_BLUE),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 11),
        
        # Body styling
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 10),
        ('BACKGROUND', (0,1), (-1,-1), colors.white),
        ('ALTERNATEROWCOLOR', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        
        # Grid and alignment
        ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
    ]))
    story.append(input_table)
    story.append(Spacer(1, 10))

    # Fittings subsection
    if fittings:
        story.append(Paragraph("System Fittings", styles["SectionHeader"]))
        fitting_data = [["Quantity", "Fitting Type", "Loss Coefficient (K)"]]
        for ft, n, k in fittings:
            fitting_data.append([str(n), ft, f"{k:.3f}"])
        
        fitting_table = Table(fitting_data, colWidths=[1*inch, 3*inch, 1.5*inch])
        fitting_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), ACCENT_BLUE),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ]))
        story.append(fitting_table)
        story.append(Spacer(1, 15))

    # Material Compatibility section
    story.append(Paragraph("Material Compatibility Analysis", styles["SectionHeader"]))
    if compatible:
        material_data = [["Material Grade", "Temperature Range (°C)", "Pressure Range (bar)"]]
        for m, tmin, tmax, pmin, pmax in compatible:
            temp_range = f"{tmin:.0f} to {tmax:.0f}"
            pressure_range = f"{pmin:.0f} to {pmax:.0f}"
            material_data.append([m, temp_range, pressure_range])
        
        material_table = Table(material_data, colWidths=[2.5*inch, 1.8*inch, 1.5*inch])
        material_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), CORPORATE_BLUE),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ]))
        story.append(material_table)
    else:
        story.append(Paragraph("⚠ No compatible materials found for the specified conditions.", styles["BodyText"]))
    story.append(Spacer(1, 15))

    # Design Analysis section
    story.append(Paragraph("Design Analysis & Optimization", styles["SectionHeader"]))
    
    # Key design metrics in a professional box
    design_data = [
        ["Design Metric", "Value", "Unit", "Status"],
        ["Critical Diameter (Velocity)", f"{plots['dcrit_velocity']:.4f}", "m", "✓ Calculated"],
        ["Critical Diameter (Pressure)", f"{plots['dcrit_pressure']:.4f}", "m", "✓ Calculated"],
        ["Optimized Diameter", f"{plots['chosen_d']:.4f}", "m", "✓ Selected"],
        ["Actual Velocity", f"{results['V']:.3f}", "m/s", "✓ Within limits"],
        ["Reynolds Number", f"{results['Re']:.0f}", "-", "✓ Acceptable"],
        ["Friction Factor", f"{results['lambda']:.6f}", "-", "✓ Calculated"],
    ]
    
    design_table = Table(design_data, colWidths=[2.2*inch, 1.2*inch, 0.8*inch, 1.3*inch])
    design_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), CORPORATE_BLUE),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
    ]))
    story.append(design_table)
    story.append(Spacer(1, 15))

    # Optimization Charts (side by side)
    story.append(Paragraph("Optimization Charts", styles["SectionHeader"]))
    if combined_plots_img.exists():
        story.append(RLImage(str(combined_plots_img), width=7*inch, height=3*inch))
    story.append(Spacer(1, 15))

    # Pressure Analysis section
    story.append(Paragraph("Hydraulic Analysis Results", styles["SectionHeader"]))
    
    pressure_data = [
        ["Pressure Component", "Value", "Unit", "Percentage"],
        ["Linear Pressure Drop", f"{results['dp_linear']/1e5:.4f}", "bar", f"{results['dp_linear']/(results['dp_linear'] + results['dp_singular'])*100:.1f}%"],
        ["Singular Pressure Drop", f"{results['dp_singular']/1e5:.4f}", "bar", f"{results['dp_singular']/(results['dp_linear'] + results['dp_singular'])*100:.1f}%"],
        ["Total Pressure Drop", f"{(results['dp_linear'] + results['dp_singular'])/1e5:.4f}", "bar", "100.0%"],
        ["Velocity Head", f"{results['H']:.4f}", "m", "N/A"],
    ]
    
    pressure_table = Table(pressure_data, colWidths=[2.2*inch, 1.2*inch, 0.8*inch, 1.3*inch])
    pressure_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), ACCENT_BLUE),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        # Highlight total row
        ('BACKGROUND', (0,-1), (-1,-1), LIGHT_BLUE),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ]))
    story.append(pressure_table)
    story.append(Spacer(1, 15))

    # ------------------------------------------------------------------
    # Wall Thickness Design  ––  one table per material
    # ------------------------------------------------------------------
    story.append(Paragraph("Wall Thickness Design", styles["SectionHeader"]))

    if thickness_results:               # thickness_results is the list you already have
        for r in thickness_results:
            mat_name = r["Material"]
            story.append(Paragraph(f"<b>{mat_name}</b>", styles["BodyText"]))
            tbl_data = [
                ["Parameter", "Value", "Unit"],
                ["Required Thickness",   f"{r['t_required_mm']:.2f}", "mm"],
                ["Outside Diameter",     f"{r['OD_norm_mm']:.2f}",    "mm"],
                ["Selected Thickness",   f"{r['t_norm_mm']:.2f}",     "mm"],
                ["Nominal Pipe Size",    f"{r['NPS']}",               ""],
                ["API Specification",    f"{r['API']}",               ""]
            ]
            t = Table(tbl_data, colWidths=[2.5*inch, 1.5*inch, 1*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), CORPORATE_BLUE),
                ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
                ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE',   (0,0), (-1,0), 10),
                ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
                ('FONTSIZE',   (0,1), (-1,-1), 9),
                ('GRID',       (0,0), (-1,-1), 0.5, GRAY_HEADER),
                ('ALIGN',      (0,0), (-1,-1), 'LEFT'),
                ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))
    else:
        story.append(Paragraph("No thickness data available.", styles["BodyText"]))
    story.append(Spacer(1, 15))

    
    # Cost Analysis section
    story.append(Paragraph("Cost Analysis", styles["SectionHeader"]))
    if prices:
        price_data = [["Material", "Mass (kg)", "Material Cost ($)", "Fittings Cost ($)", "Total Cost ($)"]]
        for mat, info in prices.items():
            if info:
                price_data.append([
                    mat,
                    f"{info['mass']:.1f}",
                    f"${info['material_cost']:.2f}",
                    f"${info['fittings_cost']:.2f}",
                    f"${info['total_cost']:.2f}"
                ])

        price_table = Table(price_data, colWidths=[2.2*inch, 1.2*inch, 1.3*inch, 1.3*inch, 1.3*inch])
        price_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), CORPORATE_BLUE),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.5, GRAY_HEADER),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ]))
        story.append(price_table)
    else:
        story.append(Paragraph("No cost data available.", styles["BodyText"]))
    # ------------------------------------------------------------------
    # Conclusions – automatically pick the cheapest material
    # ------------------------------------------------------------------
    story.append(Paragraph("Design Recommendations", styles["SectionHeader"]))

    cheapest_mat = _cheapest_material(prices)

    # find the thickness record that belongs to the cheapest material
    thickness_rec = next((t for t in thickness_results if t["Material"] == cheapest_mat), {})

    recommendations = f"""
    Based on the comprehensive analysis performed, the following design specifications are recommended:

    • <b>Material:</b> {cheapest_mat} (lowest total cost: ${prices.get(cheapest_mat, {}).get('total_cost', 0):.2f})
    • <b>Pipe Diameter:</b> {thickness_rec.get('OD_norm_mm', 'N/A')} mm (NPS {thickness_rec.get('NPS', 'N/A')})
    • <b>Wall Thickness:</b> {thickness_rec.get('t_norm_mm', 'N/A')} mm per {thickness_rec.get('API', 'N/A')} specification
    • <b>Operating Velocity:</b> {results.get('V', 'N/A'):.3f} m/s (within acceptable limits)
    • <b>Total System Pressure Drop:</b> {(results.get('dp_linear', 0) + results.get('dp_singular', 0)) / 1e5:.4f} bar

    The design meets all specified constraints and provides optimal performance for the given operating conditions.
    """
    story.append(Paragraph(recommendations, styles["BodyText"]))

    # Build the document
    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
# ------------------------------------------------------------------
# 2.  Original data loading
# ------------------------------------------------------------------
liquid_df   = pd.read_excel('liquid_properties.xlsx')
gas_df      = pd.read_excel('gas_properties.xlsx')

# MATERIALS  – 10 columns + Price
material_df = pd.read_excel('material_properties.xlsx')
material_df.columns = [
    "Material", "Specification", "Weld Joint Factor (E)", "SMYS (MPa)",
    "Roughness (mm)", "Pressure Min", "Pressure Max",
    "Temperature Min", "Temperature Max", "Price"
]

# FITTINGS  – 6 columns (Method + Type + K1 + K∞ + Kd + Price)
fittings_df = pd.read_excel('fittings.xlsx')
fittings_df.columns = ["Method", "Fitting Type", "K1", "K∞", "Kd", "Price"]
fittings_df = fittings_df[["Fitting Type", "K1", "K∞", "Kd", "Price"]]  # drop the unused Method column
fittings_df.columns = fittings_df.columns.str.strip()
fittings_df.dropna(subset=["Fitting Type"], inplace=True)

# SCHEDULE
try:
    sched_df = pd.read_excel("schedule.xlsx")
    sched_df["Outside diameter (mm)"] = pd.to_numeric(sched_df["Outside diameter (mm)"], errors="coerce")
    sched_df["Wall thickness (mm)"] = pd.to_numeric(sched_df["Wall thickness (mm)"], errors="coerce")
    sched_df.dropna(subset=["Outside diameter (mm)", "Wall thickness (mm)"], inplace=True)
except Exception as e:
    print(f"❌ Failed to load schedule file: {e}")

# Sanity-check the fittings columns we actually need
if not {'Fitting Type', 'K1', 'K∞', 'Kd', 'Price'}.issubset(fittings_df.columns):
    raise ValueError("Fittings file missing required columns.")

# ------------------------------------------------------------------
# 3.  Helper
# ------------------------------------------------------------------
def get_compatible_materials(temp, pressure):
    out = []
    for _, r in material_df.iterrows():
        try:
            if (float(r["Temperature Min"]) <= temp <= float(r["Temperature Max"]) and
                    float(r["Pressure Min"]) <= pressure <= float(r["Pressure Max"])):
                out.append(r["Material"])
        except Exception:
            pass
    return out


# ------------------------------------------------------------------
# 4.  Main application
# ------------------------------------------------------------------
class PipeDesignOptimizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pipe Design Optimizer")
        self.root.state("zoomed")
        # 1. Build the paths
        ico_file = resource_path("Pipe Design Optimizer.ico")
        png_file = resource_path("Pipe Design Optimizer.png")

        # 2. Debug prints
        print("Looking for icon at:", ico_file)
        print("Looking for png at:", png_file)

        # 3. Try setting the icon
        if os.path.isfile(ico_file):
            try:
                self.root.iconbitmap(ico_file)   # Windows taskbar & title bar
            except Exception as e:
                print("Failed to load .ico:", e)
        elif os.path.isfile(png_file):
            try:
                icon = tk.PhotoImage(file=png_file)
                self.root.iconphoto(True, icon)  # Fallback
            except Exception as e:
                print("Failed to load .png:", e)
        else:
            print("No icon file found!")




        self.style = tb.Style("cyborg")
        self.compatible_materials = []
        self.fitting_widgets = []
        self.calculation_results = {}
        self.pressure_drop_results = {}
        self.project_name = ""

        # File paths
        self.liquid_file_path = 'liquid_properties.xlsx'
        self.gas_file_path = 'gas_properties.xlsx'

        self.create_first_page()

    # ----------------------------------------------------------
    # First page – ask project name
    # ----------------------------------------------------------
    def create_first_page(self):
        self.clear_root()
        bg_file = resource_path("bg_eppm.png")
        self.set_bg(bg_file)

        
        # Main container with modern styling - remove fixed dimensions
        main_container = tb.Frame(self.root, bootstyle="light", padding=40)
        main_container.place(relx=0.5, rely=0.5, anchor="center")
        
        # Header section with modern styling
        header_frame = tb.Frame(main_container, bootstyle="primary", padding=20)
        header_frame.pack(fill="x", pady=(0, 20))
        
        # Modern title with better typography
        title_label = tb.Label(
            header_frame, 
            text="Create New Project", 
            font=("Arial", 20, "bold"), 
            bootstyle="inverse-primary"
        )
        title_label.pack()
        
        subtitle_label = tb.Label(
            header_frame,
            text="Enter a name to get started",
            font=("Arial", 10),
            bootstyle="inverse-primary"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Content area 
        content_frame = tb.Frame(main_container, padding=20)
        content_frame.pack(fill="both", expand=True)
        
        #input section
        tb.Label(
            content_frame, 
            text="Project Name:", 
            font=("Arial", 12, "bold"), 
            bootstyle="primary"
        ).pack(anchor="w", pady=(0, 5))
        
        #entry field
        self.project_entry = tb.Entry(
            content_frame, 
            font=("Arial", 12), 
            bootstyle="primary",
            width=30
        )
        self.project_entry.pack(fill="x", pady=(0, 20), ipady=5)
        
        # Modern button
        start_button = tb.Button(
            content_frame, 
            text="Create Project", 
            bootstyle="primary",
            command=self.validate_project_name,
            width=20
        )
        start_button.pack(pady=20)
        
        # Add enter key binding and focus (with error handling)
        def safe_enter_handler(event):
            try:
                if hasattr(self, 'project_entry') and self.project_entry.winfo_exists():
                    self.validate_project_name()
            except:
                pass
        
        self.root.bind('<Return>', safe_enter_handler)
        self.project_entry.focus_set()

    def validate_project_name(self):
        """Enhanced validation with better UX"""
        name = self.project_entry.get().strip()
        
        # Check if empty
        if not name:
            messagebox.showerror("Error", "Please enter a project name.")
            self.project_entry.focus_set()
            return
        
        self.project_name = name
        self.create_second_page()
    # ----------------------------------------------------------
    # Second page – input form
    # ----------------------------------------------------------
    def create_second_page(self):
        self.clear_root()
        bg_file = resource_path("bg_eppm.png")
        print("Looking for background at:", bg_file)

        if os.path.isfile(bg_file):
            self.set_bg(bg_file)
        else:
            print("Background image not found!")



        frm = self.create_scrollable_frame(self.root)

        # Title
        tb.Label(frm, text="Pipeline Design Input Form", font=("Helvetica", 20, "bold"),
                 bootstyle="primary").grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Section 1 – Basic Parameters
        section1 = tb.Labelframe(frm, text="Basic Parameters", bootstyle="info", padding=15)
        section1.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        section1.grid_columnconfigure(1, weight=1)

        self.entries = {}
        basic_labels = [("Flowrate (m³/h):", "flowrate"), ("Pipe Length (m):", "length")]
        for i, (label, key) in enumerate(basic_labels):
            tb.Label(section1, text=label, font=("Helvetica", 14)).grid(row=i, column=0, sticky="w", pady=5, padx=5)
            self.entries[key] = tb.Entry(section1, font=("Helvetica", 13), width=20)
            self.entries[key].grid(row=i, column=1, pady=5, sticky="ew")

        # Section 2 – Fluid Details
        section2 = tb.Labelframe(frm, text="Fluid Details", bootstyle="info", padding=15)
        section2.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        section2.grid_columnconfigure(1, weight=1)

        # Phase + Add Button
        tb.Label(section2, text="Phase:", font=("Helvetica", 14)).grid(row=0, column=0, sticky="w", pady=5)
        phase_frame = tb.Frame(section2)
        phase_frame.grid(row=0, column=1, sticky="ew")
        self.phase_var = tb.StringVar(value="Liquid")
        self.phase_cb = tb.Combobox(phase_frame, textvariable=self.phase_var, font=("Helvetica", 13),
                                    width=20, state="readonly", values=["Liquid", "Gas"])
        self.phase_cb.grid(row=0, column=0, padx=(0, 5), sticky="w")
        tb.Button(phase_frame, text="+", bootstyle="success", width=3,
                  command=self.open_add_fluid_window).grid(row=0, column=1)
        self.phase_cb.bind("<<ComboboxSelected>>", self.update_fluid_list)

        # Fluid Type
        tb.Label(section2, text="Fluid Type:", font=("Helvetica", 14)).grid(row=1, column=0, sticky="w", pady=5)
        self.fluid_var = tb.StringVar()
        self.fluid_cb = tb.Combobox(section2, textvariable=self.fluid_var, font=("Helvetica", 13),
                                    width=25, state="readonly")
        self.fluid_cb.grid(row=1, column=1, pady=5, sticky="ew")
        self.update_fluid_list()

        # Location
        tb.Label(section2, text="Location Type:", font=("Helvetica", 14)).grid(row=2, column=0, sticky="w", pady=5)
        self.location_var = tb.StringVar()
        locs = ["<10 buildings", "<46 buildings", ">46 buildings", "High-density/traffic"]
        tb.Combobox(section2, textvariable=self.location_var, font=("Helvetica", 13),
                    width=25, values=locs, state="readonly").grid(row=2, column=1, pady=5, sticky="ew")
        self.location_var.set(locs[0])

        # Section 3 – Operating Conditions
        section3 = tb.Labelframe(frm, text="Operating Conditions", bootstyle="info", padding=15)
        section3.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)
        section3.grid_columnconfigure(1, weight=1)

        entry_keys = [
            ("Max Velocity (m/s):", "velocity"),
            ("Temperature (°C):", "temperature"),
            ("Operating Pressure (bar):", "pressure"),
            ("Design Pressure (bar):", "design_pressure"),
            ("Corrosion Allowance (mm):", "corrosion_allowance")
        ]
        for i, (label, key) in enumerate(entry_keys):
            tb.Label(section3, text=label, font=("Helvetica", 14)).grid(row=i, column=0, sticky="w", pady=5, padx=5)
            self.entries[key] = tb.Entry(section3, font=("Helvetica", 13), width=20)
            self.entries[key].grid(row=i, column=1, pady=5, sticky="ew")

        # Results
        self.result_label = tb.Label(frm, text="", font=("Helvetica", 12), bootstyle="secondary")
        self.result_label.grid(row=4, column=0, columnspan=2, pady=10)

        # Buttons
        btn_frame = tb.Frame(frm)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=15)

        tb.Button(btn_frame, text="Submit", bootstyle="success-outline", width=20,
                  command=self.process_input).pack(side="left", padx=10)
        self.pd_btn = tb.Button(btn_frame, text="Pressure Drop Calculation", bootstyle="warning-outline",
                                width=25, command=self.create_pressure_drop_page, state="disabled")
        self.pd_btn.pack(side="left", padx=10)

    # ----------------------------------------------------------
    # Add fluid window 
    # ----------------------------------------------------------
    def open_add_fluid_window(self):
        phase = self.phase_var.get()
        add_window = tb.Toplevel(self.root)
        add_window.title(f"Add New {phase}")
        add_window.geometry("500x600")
        add_window.configure(bg='#2b2b2b')
        add_window.transient(self.root)
        add_window.grab_set()

        # Scrollable container
        container = tb.Frame(add_window)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, bg="#2b2b2b", highlightthickness=0)
        scrollbar = tb.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tb.Frame(canvas)  # This is where all entry widgets will go

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Title
        tb.Label(scroll_frame, text=f"Add New {phase}", font=("Helvetica", 18, "bold"),
                 bootstyle="primary").pack(pady=(0, 20))

        # Get columns and sample data
        columns = liquid_df.columns.tolist() if phase == "Liquid" else gas_df.columns.tolist()
        sample_data = liquid_df.iloc[0] if phase == "Liquid" else gas_df.iloc[0]

        # Create entry fields
        entries = {}
        for i, col in enumerate(columns):
            tb.Label(scroll_frame, text=f"{col}:", font=("Helvetica", 12), bootstyle="info").pack(anchor="w", pady=(10, 5))
            entry = tb.Entry(scroll_frame, font=("Helvetica", 12), width=40)
            entry.pack(anchor="w", padx=(20, 0))
            entries[col] = entry

            # Placeholder with example value
            if sample_data is not None:
                val = str(sample_data[col])
                if val and val != 'nan':
                    entry.insert(0, f"e.g., {val}")
                    entry.config(foreground='gray')

                    def on_in(e=entry):
                        if e.cget('foreground') == 'gray':
                            e.delete(0, 'end')
                            e.config(foreground='white')

                    def on_out(e=entry, ph=f"e.g., {val}"):
                        if not e.get():
                            e.insert(0, ph)
                            e.config(foreground='gray')

                    entry.bind('<FocusIn>', lambda *_: on_in())
                    entry.bind('<FocusOut>', lambda *_: on_out())

    # Add fluid logic
    def add_fluid():
        new_data = {}
        for col, entry in entries.items():
            val = entry.get().strip()
            if val.startswith("e.g.,") or not val:
                if col == columns[0]:  # Name column is mandatory
                    messagebox.showerror("Error", f"{col} is required!")
                    return
                new_data[col] = ""
            else:
                new_data[col] = val

        if not new_data.get(columns[0]):
            messagebox.showerror("Error", f"{columns[0]} is required!")
            return

        if phase == "Liquid":
            if new_data[columns[0]].lower() in liquid_df[columns[0]].str.lower().tolist():
                messagebox.showerror("Error", "Liquid already exists!")
                return
            liquid_df.loc[len(liquid_df)] = new_data
            liquid_df.to_excel(self.liquid_file_path, index=False)
        else:
            if new_data[columns[0]].lower() in gas_df[columns[0]].str.lower().tolist():
                messagebox.showerror("Error", "Gas already exists!")
                return
            gas_df.loc[len(gas_df)] = new_data
            gas_df.to_excel(self.gas_file_path, index=False)

        self.update_fluid_list()
        self.fluid_var.set(new_data[columns[0]])
        messagebox.showinfo("Success", f"{phase} '{new_data[columns[0]]}' added!")
        add_window.destroy()

        # Action buttons
        btn_frame = tb.Frame(scroll_frame)
        btn_frame.pack(pady=30)
        tb.Button(btn_frame, text="Add Fluid", bootstyle="success-outline",
                  command=add_fluid, width=15).pack(side="left", padx=(0, 10))
        tb.Button(btn_frame, text="Cancel", bootstyle="secondary-outline",
                  command=add_window.destroy, width=15).pack(side="left")
    # ----------------------------------------------------------
    # Process inputs & velocity plot
    # ----------------------------------------------------------
    def process_input(self):
        try:
            Q = float(self.entries["flowrate"].get())
            L = float(self.entries["length"].get())
            self.selected_fluid = self.fluid_var.get()
            self.selected_phase = self.phase_var.get()
            vmax = float(self.entries["velocity"].get())
            temp = float(self.entries["temperature"].get())
            pres = float(self.entries["pressure"].get())
            self.design_pressure = float(self.entries["design_pressure"].get()) * 1e5
            self.corrosion_allowance = float(self.entries["corrosion_allowance"].get())
        except ValueError:
            self.result_label.config(text="⚠️ Please enter valid numbers.")
            return

        # Fluid properties
        try:
            if self.selected_phase == "Liquid":
                props = liquid_df[liquid_df["Liquid"] == self.selected_fluid].iloc[0]
                self.rho = float(str(props["Density (kg/mÂ³)"]).replace("Â", ""))
                self.mu = float(str(props["Viscosity (mPaÂ·s)"]).replace("Â", "")) * 1e-3
            else:
                props = gas_df[gas_df["Gas"] == self.selected_fluid].iloc[0]
                self.rho = float(str(props["Density (kg/m³)"]).replace("Â", ""))
                self.mu = float(str(props["Viscosity (μPa·s)"])) * 1e-6
        except Exception as e:
            messagebox.showerror("Fluid Error", str(e))
            return

        # Compatible materials
        self.compatible_materials = get_compatible_materials(temp, pres)
        if not self.compatible_materials:
            self.result_label.config(text="❌ No compatible materials.")
            return

        self.result_label.config(
            text="✔ Materials:\n" + "\n".join(f"- {m}" for m in self.compatible_materials)
        )
        self.pd_btn.config(state="normal")

        # Critical diameter for velocity
        Q_m3s = Q / 3600
        self.dcrit_velocity = (4 * Q_m3s / (np.pi * vmax)) ** 0.5

        # Velocity vs diameter plot
        diameters = np.linspace(0.008, 2.05, 1000)
        velocities = Q_m3s / ((np.pi * diameters ** 2) / 4)
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(diameters, velocities, label="Velocity vs Diameter", color="cyan", linewidth=2)
        ax.axhline(vmax, color="red", linestyle="--", linewidth=2, label=f"Max Velocity = {vmax:.2f} m/s")
        ax.axvline(self.dcrit_velocity, color="green", linestyle="--", linewidth=2,
                   label=f"Critical Diameter = {self.dcrit_velocity:.4f} m")
        ax.set_xlabel("Diameter (m)")
        ax.set_ylabel("Velocity (m/s)")
        ax.set_title("Velocity vs Diameter")
        ax.legend()
        ax.grid(True, alpha=0.3)

        
        # Save immediately before showing
        img_folder = Path(__file__).parent / "report_imgs"
        img_folder.mkdir(exist_ok=True)
        plt.savefig(img_folder / "velocity_vs_diameter.png", dpi=150, bbox_inches="tight")
        plt.show()

        # Store basic inputs
        self.flowrate = Q
        self.pipe_length = L
        self.vmax = vmax
        self.temperature = temp
        self.operating_pressure = pres

    # ----------------------------------------------------------
    # Pressure-drop page
    # ----------------------------------------------------------
    def create_pressure_drop_page(self):
        self.clear_root()
        bg_file = resource_path("bg_eppm.png")
        print("Looking for background at:", bg_file)

        if os.path.isfile(bg_file):
            self.set_bg(bg_file)
        else:
            print("Background image not found!")
        main_frame = self.create_scrollable_frame(self.root)

        tb.Label(main_frame, text="Pressure Drop Calculation", font=("Helvetica", 24, "bold"),
                 bootstyle="primary").grid(row=0, column=0, columnspan=4, pady=(0, 20))

        info = (f"Fluid: {self.selected_fluid} | ρ={self.rho:.1f} kg/m³ | μ={self.mu:.5f} Pa·s\n"
                f"Flowrate: {self.flowrate} m³/h | Length: {self.pipe_length} m | Max Velocity: {self.vmax} m/s")
        tb.Label(main_frame, text=info, font=("Helvetica", 14), bootstyle="info").grid(
            row=1, column=0, columnspan=4, pady=(0, 20))

        tb.Label(main_frame, text="Max Pressure Drop (bar):", font=("Helvetica", 14, "bold")).grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=10)
        self.dp_max_entry = tb.Entry(main_frame, font=("Helvetica", 14), width=20)
        self.dp_max_entry.grid(row=2, column=1, sticky="w", pady=10)

        tb.Label(main_frame, text="Fittings Configuration:", font=("Helvetica", 16, "bold"),
                 bootstyle="warning").grid(row=3, column=0, columnspan=4, pady=(20, 10), sticky="w")

        headers = ["Fitting Type:", "Number:"]
        for i, h in enumerate(headers):
            tb.Label(main_frame, text=h, font=("Helvetica", 12, "bold")).grid(
                row=4, column=i, sticky="w", padx=(20 if i == 0 else 0, 10))

        self.fitting_widgets = []
        self.add_fitting_row(main_frame, 5)

        tb.Button(main_frame, text="Add Fitting +", style="info-outline",
                  command=lambda: self.add_new_fitting_row(main_frame), width=15).grid(
            row=100, column=0, columnspan=2, pady=20, sticky="w", padx=(20, 0))

        tb.Button(main_frame, text="Calculate Pressure Drop", bootstyle="success-outline",
                  width=25, command=self.calculate_pressure_drop).grid(row=101, column=0, columnspan=4, pady=20)

        self.result_label = tb.Label(main_frame, text="", font=("Helvetica", 12), bootstyle="info")
        self.result_label.grid(row=102, column=0, columnspan=4, pady=20)

        self.material_buttons_frame = tb.Frame(main_frame, bootstyle="secondary")
        self.material_buttons_frame.grid(row=103, column=0, columnspan=4, pady=10)

        tb.Button(main_frame, text="Thickness & Schedule Selection", bootstyle="primary-outline",
                  command=self.create_thickness_schedule_page, width=30).grid(row=104, column=0, columnspan=4, pady=10)

    # ----------------------------------------------------------
    # Fitting utilities
    # ----------------------------------------------------------
    def add_fitting_row(self, parent, row_num):
        cb = tb.Combobox(parent, font=("Helvetica", 12), width=30,
                         values=fittings_df["Fitting Type"].tolist(), state="readonly")
        cb.grid(row=row_num, column=0, padx=(20, 10), pady=5, sticky="w")
        ent = tb.Entry(parent, font=("Helvetica", 12), width=15)
        ent.grid(row=row_num, column=1, padx=(0, 10), pady=5, sticky="w")
        self.fitting_widgets.append((cb, ent))

    def add_new_fitting_row(self, parent):
        self.add_fitting_row(parent, 5 + len(self.fitting_widgets))

    # ----------------------------------------------------------
    # Pressure-drop calculation
    # ----------------------------------------------------------
    # ----------------------------------------------------------
    # 1. Entry point – shows progress window and launches thread
    # ----------------------------------------------------------
    def calculate_pressure_drop(self):
        """Shows a tiny modal progress window and starts the worker."""
        self.progress = tb.Toplevel(self.root)
        self.progress.title("Calculating…")
        self.progress.geometry("300x120")
        self.progress.transient(self.root)
        self.progress.grab_set()
        self.progress.resizable(False, False)

        tb.Label(self.progress,
                 text="Computing pressure drop,\nplease wait…",
                 font=("Segoe UI", 12)).pack(pady=25)

        pbar = tb.Progressbar(self.progress, mode='indeterminate')
        pbar.pack(fill='x', padx=20, pady=10)
        pbar.start(10)

        self.progress.update_idletasks()          # paint it now

        # gather the input values **before** the thread
        try:
            dp_max = float(self.dp_max_entry.get()) * 1e5
        except ValueError:
            self.result_label.config(text="⚠️ Enter valid max pressure drop.")
            self.progress.destroy()
            return

        # pack what the thread needs into a dict
        job = dict(
            dp_max=dp_max,
            Q=self.flowrate,
            L=self.pipe_length,
            rho=self.rho,
            mu=self.mu,
            vmax=self.vmax,
            diameters=np.linspace(0.008, 2.05, 1000),
            compatible_materials=self.compatible_materials,
            fitting_widgets=self.fitting_widgets,
        )

        # launch the worker
        threading.Thread(
            target=self._pressure_drop_worker,
            args=(job,),
            daemon=True
        ).start()


    # ----------------------------------------------------------
    # 2. Background thread – does only the math
    # ----------------------------------------------------------
    def _pressure_drop_worker(self, job):
        """Heavy calculation (no GUI calls)."""
        try:
            diameters = np.linspace(0.008, 2.05, 1000)   # <- ADD THIS LINE
            results = []
            calc_results = {}

            for mat in job["compatible_materials"]:
                row = material_df[material_df["Material"] == mat].iloc[0]
                k = float(str(row["Roughness (mm)"]).replace("Â", "")) * 1e-3
                total_deltas = []

                for D in diameters:
                    V = (job["Q"] / 3600) / ((np.pi * D ** 2) / 4)
                    if V > job["vmax"]:
                        total_deltas.append(np.nan)
                        continue
                    Re = job["rho"] * V * D / job["mu"]
                    lam = 64 / Re if Re < 2300 else self.colebrook(Re, D, k)
                    H = V ** 2 / (2 * 9.81)
                    dP_lin = lam * (job["L"] / D) * job["rho"] * 9.81 * H

                    dP_sing = 0
                    for cb, e in job["fitting_widgets"]:
                        typ = cb.get().strip()
                        if not typ:
                            continue
                        try:
                            n = int(e.get() or 0)
                        except ValueError:
                            n = 0
                        if n > 0:
                            frow = fittings_df[
                                fittings_df["Fitting Type"].str.strip().str.lower() == typ.strip().lower()]
                            if frow.empty:
                                continue
                            frow = frow.iloc[0]
                            K1, Kinf, Kd = float(frow["K1"]), float(frow["K∞"]), float(frow["Kd"])
                            K = Kinf + (K1 - Kinf) * (Re ** (-1 / Kd))
                            dP_sing += n * K * job["rho"] * V ** 2 / 2
                    total_deltas.append(dP_lin + dP_sing)

                td = np.array(total_deltas)
                valid = (td <= job["dp_max"]) & (~np.isnan(td))
                idx = np.where(valid)[0]
                dcrit = diameters[idx[0]] if len(idx) else None
                results.append((mat, dcrit))

                if dcrit:
                    self.store_detailed_calculation(mat, dcrit,
                                                    job["Q"], job["L"],
                                                    job["rho"], job["mu"])
                    calc_results[mat] = self.calculation_results[mat]

            payload = dict(results=results, calc_results=calc_results)

        except Exception as exc:
            payload = dict(error=str(exc))

        # ---- generate plot in the worker ----
        matplotlib.use('Agg')          # headless backend
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(figsize=(10, 7))
        for mat, dcrit in results:
            row = material_df[material_df["Material"] == mat].iloc[0]
            k = float(str(row["Roughness (mm)"]).replace("Â", "")) * 1e-3
            total_deltas = []

            for D in np.linspace(0.008, 2.05, 1000):
                V = (self.flowrate / 3600) / ((np.pi * D ** 2) / 4)
                if V > self.vmax:
                    total_deltas.append(np.nan)
                    continue
                Re = self.rho * V * D / self.mu
                lam = 64 / Re if Re < 2300 else self.colebrook(Re, D, k)
                H = V ** 2 / (2 * 9.81)
                dP_lin = lam * (self.pipe_length / D) * self.rho * 9.81 * H

                dP_sing = 0
                for cb, e in self.fitting_widgets:
                    typ = cb.get().strip()
                    if not typ:
                        continue
                    try:
                        n = int(e.get() or 0)
                    except ValueError:
                        n = 0
                    if n > 0:
                        frow = fittings_df[
                            fittings_df["Fitting Type"].str.strip().str.lower() == typ.strip().lower()]
                        if frow.empty:
                            continue
                        frow = frow.iloc[0]
                        K1, Kinf, Kd = float(frow["K1"]), float(frow["K∞"]), float(frow["Kd"])
                        K = Kinf + (K1 - Kinf) * (Re ** (-1 / Kd))
                        dP_sing += n * K * self.rho * V ** 2 / 2
                total_deltas.append(dP_lin + dP_sing)

            ax.plot(np.linspace(0.008, 2.05, 1000),
                    np.array(total_deltas) / 1e5,
                    label=mat, linewidth=2)

        ax.axhline(float(self.dp_max_entry.get()), color="red",
                   linestyle="--", linewidth=2, label="ΔP max")
        ax.set_xlabel("Diameter (m)")
        ax.set_ylabel("Total ΔP (bar)")
        ax.set_title("Total Pressure Drop vs Diameter")
        ax.legend()
        ax.grid(True, alpha=0.3)
        img_folder = Path(__file__).parent / "report_imgs"
        img_folder.mkdir(exist_ok=True)
        plt.savefig(img_folder / "pressure_drop_vs_diameter.png",
                    dpi=150, bbox_inches="tight")
        plt.close(fig)

        self.root.after(0, self._calculation_done, payload)

    # ----------------------------------------------------------
    # 3. Back in main thread – close progress, plot, update GUI
    # ----------------------------------------------------------
    def _calculation_done(self, payload):
        """Executed in main thread – safe for Tk & Matplotlib."""
        self.progress.destroy()

        if payload.get("error"):
            self.result_label.config(text=f"Error:\n{payload['error']}")
            return

        results = payload["results"]
        self.calculation_results.update(payload["calc_results"])



        # ---- update GUI ----
        self.pressure_drop_results = {}
        for mat, dcrit in results:
            self.pressure_drop_results[mat] = {
                "Critical Diameter (m)": dcrit,
                "Critical Diameter (mm)": dcrit * 1000 if dcrit else None
            }

        text = "✅ Critical Diameter Summary:\n\n"
        for mat, dc in results:
            text += f"🔹 {mat}: {dc:.4f} m\n" if dc else f"🔹 {mat}: No valid diameter\n"
        self.result_label.config(text=text)
        self.show_material_buttons()




    
    def store_detailed_calculation(self, mat, dcrit, Q, L, rho, mu):
        V = (Q / 3600) / ((np.pi * dcrit ** 2) / 4)
        Re = rho * V * dcrit / mu
        row = material_df[material_df["Material"] == mat].iloc[0]
        k = float(str(row["Roughness (mm)"]).replace("Â", "")) * 1e-3
        lam = 64 / Re if Re < 2300 else self.colebrook(Re, dcrit, k)
        H = V ** 2 / (2 * 9.81)
        dP_lin = lam * (L / dcrit) * rho * 9.81 * H
        dP_sing = 0
        details = []
        for cb, e in self.fitting_widgets:
            typ = cb.get().strip()
            if not typ:
                continue
            try:
                n = int(e.get() or 0)
            except ValueError:
                n = 0
            if n > 0:
                frow = fittings_df[fittings_df["Fitting Type"].str.strip().str.lower() == typ.strip().lower()]
                if frow.empty:
                    continue   # skip unknown / mistyped fittings
                frow = frow.iloc[0]
                K1, Kinf, Kd = float(frow["K1"]), float(frow["K∞"]), float(frow["Kd"])
                K = Kinf + (K1 - Kinf) * (Re ** (-1 / Kd))
                dp_part = n * K * rho * V ** 2 / 2
                dP_sing += dp_part
                details.append((typ, n, K))
        self.calculation_results[mat] = {
            'diameter': dcrit,
            'velocity': V,
            'reynolds': Re,
            'lambda': lam,
            'H': H,
            'dp_linear': dP_lin,
            'dp_singular': dP_sing,
            'details': details
        }

    def show_material_buttons(self):
        for w in self.material_buttons_frame.winfo_children():
            w.destroy()
        tb.Label(self.material_buttons_frame, text="Detailed Calculations:",
                 font=("Helvetica", 14, "bold"), bootstyle="warning").pack(pady=(0, 10))
        btn_frame = tb.Frame(self.material_buttons_frame)
        btn_frame.pack()
        for i, mat in enumerate(self.compatible_materials):
            if mat in self.calculation_results:
                tb.Button(btn_frame, text=mat, bootstyle="info-outline", width=20,
                          command=lambda m=mat: self.show_detailed_calculation(m)).grid(
                    row=i // 3, column=i % 3, padx=5, pady=5)

    def show_detailed_calculation(self, material):
        """Modern card-based details window."""
        calc = self.calculation_results[material]

        win = tb.Toplevel(self.root)
        win.title(f"{material} — Calculation Details")
        win.geometry("680x650")
        win.minsize(600, 500)
        win.transient(self.root)
        win.grab_set()
        win.configure(bg="#1e1e1e")

        # ---------- scrollable frame ----------
        container = tb.Frame(win, bootstyle="dark")
        container.pack(fill="both", expand=True, padx=24, pady=24)

        canvas = tk.Canvas(container, bg="#1e1e1e", highlightthickness=0)
        scrollbar = tb.Scrollbar(container, orient="vertical", command=canvas.yview)
        body = tb.Frame(canvas, bootstyle="dark")

        canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")

        # allow mouse wheel
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # keep body width same as canvas
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # make inner frame follow canvas width
        def _on_canvas_resize(ev):
            canvas.itemconfig(canvas_window, width=ev.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        # ---------- reusable card ----------
        def card(title, items):
            """Return a frame that looks like a modern card."""
            card_frame = tb.Frame(body, bootstyle="secondary")
            card_frame.pack(fill="x", pady=(0, 20))

            # header
            hdr = tb.Label(
                card_frame,
                text=title,
                font=("Segoe UI", 12, "bold"),
                bootstyle="inverse-secondary",
                anchor="w",
            )
            hdr.pack(fill="x", padx=16, pady=12)

            # divider
            sep = tk.Frame(card_frame, height=1, bg="#2e2e2e")
            sep.pack(fill="x", padx=16)

            # content
            content = tb.Frame(card_frame, bootstyle="secondary")
            content.pack(fill="x", padx=16, pady=12)

            for label, value in items:
                row = tb.Frame(content, bootstyle="secondary")
                row.pack(fill="x", pady=4)

                tb.Label(
                    row,
                    text=f"{label}:",
                    font=("Segoe UI", 10),
                    bootstyle="secondary",
                ).pack(side="left")

                tb.Label(
                    row,
                    text=value,
                    font=("Segoe UI", 10, "bold"),
                    bootstyle="info",
                ).pack(side="right")

            return card_frame

        # ---------- data ----------
        card("Input Parameters", [
            ("Fluid Type", self.selected_fluid),
            ("Density ρ", f"{self.rho:.1f} kg/m³"),
            ("Viscosity μ", f"{self.mu:.6f} Pa·s"),
            ("Flow Rate Q", f"{self.flowrate} m³/h"),
            ("Pipe Length L", f"{self.pipe_length} m"),
        ])

        card("Calculated Values", [
            ("Critical Diameter", f"{calc['diameter']:.4f} m"),
            ("Flow Velocity", f"{calc['velocity']:.3f} m/s"),
            ("Reynolds Number", f"{calc['reynolds']:.0f}"),
            ("Friction Factor λ", f"{calc['lambda']:.6f}"),
            ("Linear ΔP", f"{calc['dp_linear'] / 1e5:.4f} bar"),
            ("Singular ΔP", f"{calc['dp_singular'] / 1e5:.4f} bar"),
            ("Total ΔP", f"{(calc['dp_linear'] + calc['dp_singular']) / 1e5:.4f} bar"),
        ])

        if calc["details"]:
            card("Fittings", [(f"{n} × {ft}", f"K = {K:.3f}") for ft, n, K in calc["details"]])

        # ---------- footer ----------
        tb.Button(
            body,
            text="Close",
            bootstyle="secondary-outline",
            command=win.destroy,
        ).pack(pady=(10, 24))

        # center window
        win.update_idletasks()
        win.geometry(
            "+{}+{}".format(
                (win.winfo_screenwidth() - win.winfo_width()) // 2,
                (win.winfo_screenheight() - win.winfo_height()) // 2,
            )
        )
# ----------------------------------------------------------
# Thickness & schedule page  —  shown in a child Toplevel
# ----------------------------------------------------------
    def create_thickness_schedule_page(self):
        """Open a scrollable thickness & schedule window, not full-screen."""
        # ---- new child window (modal or not, your choice) ----
        top = tk.Toplevel(self.root)
        top.title("Thickness & Schedule Selection")
        top.geometry("800x600")
        top.transient(self.root)            # stay above parent
        top.grab_set()                      # make it modal (remove if you like)
        top.resizable(True, True)

        # ---- icon (same trick as main window) ----
        try:
            ico_path = Path(__file__).with_name("Pipe Design Optimizer.ico")
            top.iconbitmap(str(ico_path))
        except Exception:
            pass

        # ---- style frame inside the Toplevel ----
        root_frame = tb.Frame(top, bootstyle="dark")
        root_frame.pack(fill="both", expand=True)

        # header
        hdr = tb.Frame(root_frame)
        hdr.pack(fill="x", pady=10)
        tb.Label(hdr, text="Thickness & Schedule Selection",
                 font=("Segoe UI", 18, "bold"), bootstyle="primary").pack()
        tb.Label(hdr, text="Automated pipe sizing based on critical diameter and pressure requirements",
                 font=("Segoe UI", 9), bootstyle="secondary").pack()

        # scrollable canvas
        canvas = tk.Canvas(root_frame, bg="#1e1e1e", highlightthickness=0)
        scrollbar = tb.Scrollbar(root_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tb.Frame(canvas, bootstyle="dark")

        scroll_frame.bind("<Configure>",
                          lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(canvas_frame, width=e.width))
        canvas_frame = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---- load schedule (unchanged) ----
        try:
            sched_df = pd.read_excel("schedule.xlsx")
            sched_df["Outside diameter (mm)"] = pd.to_numeric(sched_df["Outside diameter (mm)"], errors="coerce")
            sched_df["Wall thickness (mm)"] = pd.to_numeric(sched_df["Wall thickness (mm)"], errors="coerce")
            sched_df.dropna(subset=["Outside diameter (mm)", "Wall thickness (mm)"], inplace=True)
            status_msg = "✓ Schedule data loaded"
        except Exception as e:
            status_msg = f"⚠ Failed to load schedule.xlsx: {e}"

        status_bar = tb.Frame(scroll_frame, bootstyle="success" if "✓" in status_msg else "danger")
        status_bar.pack(fill="x", pady=(0, 10), padx=5)
        tb.Label(status_bar, text=status_msg, font=("Segoe UI", 9)).pack(pady=5)
        if "⚠" in status_msg:
            return

        # ---- parameters (unchanged) ----
        F_map = {"<10 buildings": 0.72, "<46 buildings": 0.60,
                 ">46 buildings": 0.50, "High-density/traffic": 0.40}
        F = F_map.get(self.location_var.get(), 0.72)

        params_lf = tb.LabelFrame(scroll_frame, text="Parameters", bootstyle="info")
        params_lf.pack(fill="x", pady=(0, 10), padx=5)
        params_inner = tb.Frame(params_lf)
        params_inner.pack(fill="x", padx=10, pady=5)
        tb.Label(params_inner, text=f"Factor (F): {F}").grid(row=0, column=0, sticky="w")
        tb.Label(params_inner, text=f"Pressure: {self.design_pressure/1e6:.2f} MPa").grid(row=0, column=1, sticky="w", padx=(20, 0))

        # ---- compute & display results (unchanged) ----
        results = []
        for mat in self.compatible_materials:
            dc_vel = self.dcrit_velocity * 1000
            dc_pres = self.pressure_drop_results.get(mat, {}).get("Critical Diameter (mm)")
            dcrit = min(dc_vel, dc_pres) if dc_pres else dc_vel
            if dcrit is None:
                continue
            row = material_df[material_df["Material"] == mat].iloc[0]
            try:
                E = float(row["Weld Joint Factor (E)"])
                Sy = float(row["SMYS (MPa)"]) * 1e6
            except:
                continue
            CA = self.corrosion_allowance / 1000
            S = F * E * Sy
            t_req = (self.design_pressure * 10 * (dcrit / 1000)) / (20 * S - 2 * self.design_pressure) + CA
            OD = dcrit + 2 * t_req * 1000

            cand = sched_df[sched_df["Outside diameter (mm)"] >= OD]
            if cand.empty:
                continue
            min_od = cand["Outside diameter (mm)"].min()
            cand = cand[cand["Outside diameter (mm)"] == min_od]
            cand = cand[cand["Wall thickness (mm)"] >= t_req * 1000]
            if cand.empty:
                continue
            best = cand.iloc[0]
            results.append({
                "Material": mat,
                "dcrit (m)": dcrit / 1000,
                "t_required_mm": t_req * 1000,
                "OD_computed_mm": OD,
                "OD_norm_mm": best["Outside diameter (mm)"],
                "t_norm_mm": best["Wall thickness (mm)"],
                "NPS": best.get("Nominal size (inches)", "N/A"),
                "API": best.get("Specif. API", "N/A")
            })

        if not results:
            tb.Label(scroll_frame, text="⚠ No matching schedule found",
                     font=("Segoe UI", 14, "bold"), bootstyle="warning").pack(pady=20)
        else:
            tb.Label(scroll_frame, text="Recommended Pipe Specifications",
                     font=("Segoe UI", 16, "bold"), bootstyle="primary").pack(pady=(10, 20))

            for i, r in enumerate(results):
                card = tb.LabelFrame(scroll_frame, text=f"Option {i+1}: {r['Material']}", bootstyle="secondary")
                card.pack(fill="x", pady=(0, 15), padx=10)
                inner = tb.Frame(card)
                inner.pack(fill="x", padx=15, pady=10)

                left = tb.Frame(inner)
                left.grid(row=0, column=0, sticky="nw", padx=(0, 20))
                tb.Label(left, text="Critical Parameters", font=("Segoe UI", 11, "bold")).pack(anchor="w")
                tb.Label(left, text=f"Critical Diameter: {r['dcrit (m)']:.4f} m").pack(anchor="w")
                tb.Label(left, text=f"Required Thickness: {r['t_required_mm']:.2f} mm").pack(anchor="w")
                tb.Label(left, text=f"Computed OD: {r['OD_computed_mm']:.2f} mm").pack(anchor="w")

                right = tb.Frame(inner)
                right.grid(row=0, column=1, sticky="nw")
                tb.Label(right, text="Standard Specifications", font=("Segoe UI", 11, "bold")).pack(anchor="w")
                tb.Label(right, text=f"Standard OD: {r['OD_norm_mm']:.2f} mm").pack(anchor="w")
                tb.Label(right, text=f"Wall Thickness: {r['t_norm_mm']:.2f} mm").pack(anchor="w")
                tb.Label(right, text=f"NPS: {r['NPS']} | API: {r['API']}").pack(anchor="w")

        # ---- buttons ----
        btn_bar = tb.Frame(scroll_frame)
        btn_bar.pack(fill="x", pady=30)
        btn_box = tb.Frame(btn_bar)
        btn_box.pack()
        tb.Button(btn_box, text="📄 Generate PDF Report", bootstyle="success", width=20,
                  command=lambda: [top.destroy(), self.generate_report(results)]).pack(side="left", padx=5)
        tb.Button(btn_box, text="← Back / Close", bootstyle="secondary-outline", width=20,
                  command=top.destroy).pack(side="left", padx=5)
    # ----------------------------------------------------------
    # Report generator
    # ----------------------------------------------------------
    def generate_report(self, thickness_results):
        # Collect inputs
        inputs = {
            "project_name": self.project_name,
            "flowrate_m3h": self.flowrate,
            "pipe_length_m": self.pipe_length,
            "phase": self.selected_phase,
            "fluid": self.selected_fluid,
            "max_velocity_mps": self.vmax,
            "temperature_c": self.temperature,
            "operating_pressure_bar": self.operating_pressure,
            "design_pressure_bar": self.design_pressure / 1e5,
            "location_type": self.location_var.get()
        }

        # Fittings
        inputs["fittings"] = []
        for cb, e in self.fitting_widgets:
            try:
                typ = cb.get().strip()
                if typ:
                    try:
                        n = int(e.get() or 0)
                    except ValueError:
                        n = 0
                    if n > 0:
                        inputs["fittings"].append((typ, n))
            except tk.TclError:
                continue

        # Compatible materials with ranges
        compatible = []
        for mat in self.compatible_materials:
            row = material_df[material_df["Material"] == mat].iloc[0]
            compatible.append((
                mat,
                float(row["Temperature Min"]),
                float(row["Temperature Max"]),
                float(row["Pressure Min"]),
                float(row["Pressure Max"])
            ))

        # Plots
        plots = {
            "dcrit_velocity": self.dcrit_velocity,
            "dcrit_pressure": min(
                [self.pressure_drop_results[m]["Critical Diameter (m)"]
                 for m in self.pressure_drop_results
                 if self.pressure_drop_results[m]["Critical Diameter (m)"] is not None],
                default=0
            ),
            "chosen_d": min(
                self.dcrit_velocity,
                *[self.pressure_drop_results[m]["Critical Diameter (m)"]
                  for m in self.pressure_drop_results
                  if self.pressure_drop_results[m]["Critical Diameter (m)"] is not None]
            )
        }

        # Results – pick first material for legacy hydraulic results
        if self.compatible_materials:
            res_mat = self.compatible_materials[0]
            if res_mat in self.calculation_results:
                calc = self.calculation_results[res_mat]
                results = {
                    "V": calc["velocity"],
                    "Re": calc["reynolds"],
                    "lambda": calc["lambda"],
                    "H": calc["H"],
                    "dp_linear": calc["dp_linear"],
                    "dp_singular": calc["dp_singular"]
                }
                fittings = [(ft, n, K) for ft, n, K in calc["details"]]
            else:
                results = {"V": 0, "Re": 0, "lambda": 0, "H": 0,
                           "dp_linear": 0, "dp_singular": 0}
                fittings = []
        else:
            results = {"V": 0, "Re": 0, "lambda": 0, "H": 0,
                       "dp_linear": 0, "dp_singular": 0}
            fittings = []

        # Calculate prices
        prices = self.calculate_pipe_prices(thickness_results)

        # Ask user for file
        default_name = f"{self.project_name}_report.pdf"
        file_path = Path.home() / "Documents" / default_name
        file_path = filedialog.asksaveasfilename(
            initialdir=file_path.parent,
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return

        # Build report
        build_pipeline_report(
            inputs=inputs,
            compatible=compatible,
            plots=plots,
            thickness_results=thickness_results,   # <── fixed keyword
            results=results,
            fittings=fittings,
            prices=prices,
            file_path=Path(file_path)
        )
        messagebox.showinfo("Report", f"Report saved to:\n{file_path}")


    # ----------------------------------------------------------
    # Price calculations
    # ----------------------------------------------------------
    def calculate_pipe_prices(self, thickness_results):
        prices = {}
        for r in thickness_results:
            mat = r["Material"]
            try:
                # Get material price
                mat_row = material_df[material_df["Material"] == mat].iloc[0]
                price_per_kg = float(mat_row["Price"])
                density = 7850  # kg/m³ (steel)
                OD_m = r["OD_norm_mm"] / 1000
                t_m = r["t_norm_mm"] / 1000
                volume = np.pi * (OD_m**2 - (OD_m - 2*t_m)**2) / 4 * self.pipe_length
                mass = volume * density
                material_cost = mass * price_per_kg

                # Fittings cost – skip widgets that may be destroyed
                fittings_cost = 0
                try:
                    for cb, e in self.fitting_widgets:
                        if not cb.winfo_exists() or not e.winfo_exists():
                            continue
                        typ = cb.get().strip()
                        if typ:
                            qty = int(e.get() or 0)
                            row = fittings_df[fittings_df["Fitting Type"] == typ]
                            if not row.empty:
                                unit_price = float(row.iloc[0]["Price"])
                                fittings_cost += qty * unit_price
                except tk.TclError:
                    pass  # ignore destroyed widgets

                total_cost = (material_cost + fittings_cost) * 1.10
                prices[mat] = {
                    'material_cost': material_cost,
                    'fittings_cost': fittings_cost,
                    'total_cost': total_cost,
                    'mass': mass
                }
            except Exception as e:
                print(f"Price calc error for {mat}: {e}")
        return prices


    # ----------------------------------------------------------
    # Utilities
    # ----------------------------------------------------------
    def colebrook(self, Re, D, k):
        """Swamee–Jain explicit approximation to Colebrook-White."""
        if Re <= 0 or D <= 0:
            return 0.02  # safe fallback
        ks = k / D
        A = (ks / 3.7) ** 1.11 + 5.74 / Re ** 0.9
        f = 0.25 / (np.log10(A) ** 2)
        return max(f, 1e-4)  # avoid zero / negatives

    def clear_root(self):
        for w in self.root.winfo_children():
            w.destroy()

    def set_bg(self, path):
        try:
            img = Image.open(path).resize((self.root.winfo_screenwidth(), self.root.winfo_screenheight()))
            photo = ImageTk.PhotoImage(img)
            lbl = tb.Label(self.root, image=photo)
            lbl.image = photo
            lbl.place(x=0, y=0, relwidth=1, relheight=1)
        except Exception:
            pass  # background image optional

    def create_scrollable_frame(self, parent, width=580, height=500):
        container = tb.Frame(parent)
        container.pack(pady=45)
        canvas = tk.Canvas(container, width=width, height=height, highlightthickness=0)
        scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tb.Frame(canvas, padding=20, bootstyle="secondary")
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        return scroll_frame

        
        # Configure scrolling
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            # Update scroll region when canvas changes
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Update frame width to match canvas width
            canvas.itemconfig(canvas_window, width=event.width)
        
        scroll_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Create window in canvas
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        # Configure scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack elements with proper expansion
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', bind_mousewheel)
        canvas.bind('<Leave>', unbind_mousewheel)
        
        return scroll_frame

    def update_fluid_list(self, event=None):
        phase = self.phase_var.get()
        if phase == "Liquid":
            fluids = liquid_df["Liquid"].dropna().tolist()
        else:
            fluids = gas_df["Gas"].dropna().tolist()
        self.fluid_cb['values'] = fluids
        if fluids:
            self.fluid_cb.current(0)





# ------------------------------------------------------------------
# 5.  Run app
# ------------------------------------------------------------------
if __name__ == "__main__":
    from tkinter import filedialog  # avoid import issues

    root = tb.Window(themename="cyborg")
    PipeDesignOptimizerApp(root)
    root.mainloop()
