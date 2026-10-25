"""
================================================================================
BATTERY SHOP UNIT TCC-2 (PT GMF AEROASIA Tbk) - DESKTOP DASHBOARD GUI
Aircraft Battery Maintenance Dashboard & Operational Monitoring System
Enterprise Clean UI / UX with Dark Navy Sidebar & Light Slate Workspace
================================================================================
Core Architecture:
- Single Master Excel Workbook (Worksheets: Dashboard, History FSB, History SCRAP)
- Mobile/Offline Workflow Support: Refresh Data (disk re-read) & Import Excel (switch master)
- 100% Offline & Standalone Desktop Python (CustomTkinter, ttk, Matplotlib)
- Dynamic KPIs, Stacked Bar 7-Day Trend, Donut Status Distribution, Weekly & Monthly Reports
- In-Place Cell Editing, Atomic Excel Synchronization, Automated Backups, Lock Detection
================================================================================
"""

import os
import sys
import re
import json
import socket
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk

import customtkinter as ctk
import pandas as pd
import openpyxl

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Modular Architecture & Data Integrity Engine
from config import BASE_DIR, get_master_excel_path, set_master_excel_path, load_app_config, get_resource_path
from backup_manager import create_backup, list_backups, restore_backup, get_backup_directory
from validation import (
    validate_workbook_schema,
    validate_battery_unit_record,
    validate_fsb_record,
    validate_scrap_record,
    inspect_workbook_summary
)
from sync_service import get_file_fingerprint, has_external_conflict, is_file_locked
from excel_repository import (
    ExcelRepository,
    log_audit_event,
    normalize_clean_string,
    parse_date_value,
    SHEET_DAILY,
    SHEET_FSB,
    SHEET_SCRAP
)
from analytics import (
    to_date_object,
    get_available_years,
    get_monthly_maintenance_stats,
    get_weekly_maintenance_stats,
    get_status_distribution,
    get_7day_status_trend,
    get_customer_breakdown,
    get_pn_breakdown,
    INDONESIAN_MONTH_SHORT,
    INDONESIAN_MONTH_NAMES,
    HARI_NAMES
)
from dialogs import (
    WelcomeWorkbookDialog,
    BatteryFormDialog,
    FSBFormDialog,
    SCRAPFormDialog,
    DropdownManagerDialog,
    GlossaryDialog,
    ExcelAccessDialog,
    BatteryCalendarDialog,
    RouteToFSBDialog,
    RouteToSCRAPDialog,
    format_date_with_month_name,
    MONTHS_ID,
    MONTHS_FULL_ID
)


# -----------------------------------------------------------------------------
# APPLICATION CONFIG & PERSISTENT DROPDOWN OPTIONS
# -----------------------------------------------------------------------------
APP_TITLE = "Battery Shop TCC-2 | Aircraft Battery Monitoring System"
CONFIG_FILE = BASE_DIR / "dropdown_config.json"

DEFAULT_DROPDOWN_OPTIONS = {
    "Part Number": [
        "018550-000", "024147-000", "2758", "405CH", "40678-2", "40176-7",
        "1756-3", "1658-6", "BA06-01", "BA35-01", "024637-000", "024453-000",
        "416488", "4059"
    ],
    "Customer": [
        "PT Garuda Indonesia", "Citilink Indonesia", "PT. Indonesia AirAsia",
        "Cebu Air", "Avia Tech", "PT. BBN Airlines", "BBJ"
    ],
    "Position": [
        "UR", "WR", "FSC", "FRA", "WI", "TECH", "SRAM", "SREW"
    ],
    "Status": [
        "Reguler", "Overhaul", "Periodic", "Restoration", "PIR", "FINAL CHARGE", "SCRAP"
    ]
}

def load_dropdown_options():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in DEFAULT_DROPDOWN_OPTIONS.items():
                    if k not in data or not isinstance(data[k], list):
                        data[k] = v
                # Ensure FINAL CHARGE and SCRAP are present in Status
                for opt in ["FINAL CHARGE", "SCRAP"]:
                    if opt not in data.get("Status", []):
                        data["Status"].append(opt)
                save_dropdown_options(data)
                return data
        except Exception:
            pass
    save_dropdown_options(DEFAULT_DROPDOWN_OPTIONS)
    return DEFAULT_DROPDOWN_OPTIONS.copy()

def save_dropdown_options(options):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(options, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"[Config] Error saving dropdown config: {e}")


# -----------------------------------------------------------------------------
# MAIN APPLICATION CLASS
# -----------------------------------------------------------------------------
class BatteryShopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Appearance configuration - Enterprise Light by Default
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("1440x900")
        self.minsize(1080, 700)

        # Set Window & Taskbar Icon if available
        icon_path = get_resource_path("assets/app_icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Data Source State
        self.excel_file_path = get_master_excel_path()
        self.excel_repo = ExcelRepository(self.excel_file_path)
        self.last_sync_fingerprint = None
        self.dropdown_options = load_dropdown_options()

        # DataFrames
        self.df_daily = pd.DataFrame()
        self.df_fsb = pd.DataFrame()
        self.df_scrap = pd.DataFrame()

        # In-line Editor Tracking
        self.active_inline_editor = None

        # Filter & Search Variables
        self.daily_cust_filter = tk.StringVar(value="ALL")
        self.daily_pn_filter = tk.StringVar(value="ALL")
        self.daily_pos_filter = tk.StringVar(value="ALL")
        self.daily_status_filter = tk.StringVar(value="ALL")
        self.daily_search_var = tk.StringVar(value="")
        self.active_calendar_date = None

        self.fsb_search_var = tk.StringVar(value="")
        self.fsb_sort_order = tk.StringVar(value="Terbaru (Date FSB Desc)")

        self.scrap_search_var = tk.StringVar(value="")

        # Laporan Analytics State
        self.laporan_year_var = tk.StringVar(value=str(date.today().year))
        self.laporan_week_var = tk.StringVar(value=str(date.today().isocalendar()[1]))

        # Navigation State
        self.current_view_name = "Dashboard"
        self.pages = {}
        self.nav_buttons = {}

        # Build Main UI Shell
        self.setup_styles()
        self.build_ui_layout()

        # Global Keyboard Shortcuts
        self.bind("<F5>", lambda e: self.load_data(silent=False))
        self.bind("<F1>", lambda e: self.open_glossary())
        self.bind("<Delete>", lambda e: self.handle_global_delete_key())

        # Start Clock
        self.update_live_clock()

        # Render window shell immediately for instant visual feedback
        self.lbl_file_status.configure(text="● Memuat database...", text_color="#f59e0b")
        self.lbl_last_sync.configure(text="Sinkronisasi awal...")
        self.update_idletasks()

        # Initial Startup: Check Workbook & Load Data
        self.after(20, self.startup_check_and_load)

    # -------------------------------------------------------------------------
    # STYLES (TTK TREEVIEW)
    # -------------------------------------------------------------------------
    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        bg_color = "#ffffff"
        fg_color = "#0f172a"
        heading_bg = "#f1f5f9"
        heading_fg = "#334155"
        select_bg = "#dbeafe"
        select_fg = "#1e40af"

        self.style.configure(
            "Custom.Treeview",
            background=bg_color,
            foreground=fg_color,
            fieldbackground=bg_color,
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0
        )
        self.style.map(
            "Custom.Treeview",
            background=[("selected", select_bg)],
            foreground=[("selected", select_fg)]
        )
        self.style.configure(
            "Custom.Treeview.Heading",
            background=heading_bg,
            foreground=heading_fg,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=(6, 6)
        )
        self.style.map(
            "Custom.Treeview.Heading",
            background=[("active", "#e2e8f0")]
        )

    def bind_treeview_mousewheel(self, tree):
        """Binds universal smooth mousewheel scrolling to ttk.Treeview on Windows, Linux, and Mac."""
        def _on_mousewheel(event):
            try:
                if event.delta:
                    delta = int(-1 * (event.delta / 120))
                elif event.num == 4:
                    delta = -1
                elif event.num == 5:
                    delta = 1
                else:
                    delta = 0
                tree.yview_scroll(delta, "units")
            except Exception:
                pass
            return "break"

        def _on_shift_mousewheel(event):
            try:
                if event.delta:
                    delta = int(-1 * (event.delta / 120))
                else:
                    delta = 0
                tree.xview_scroll(delta, "units")
            except Exception:
                pass
            return "break"

        tree.bind("<MouseWheel>", _on_mousewheel, add="+")
        tree.bind("<Shift-MouseWheel>", _on_shift_mousewheel, add="+")
        tree.bind("<Button-4>", _on_mousewheel, add="+")
        tree.bind("<Button-5>", _on_mousewheel, add="+")

    def apply_treeview_tags(self, tree):
        """Applies soft pastel badges for clean enterprise look and binds mousewheel."""
        tree.tag_configure("even", background="#ffffff")
        tree.tag_configure("odd", background="#f8fafc")
        tree.tag_configure("UR", foreground="#15803d", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("WR", foreground="#b45309", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("WI", foreground="#6b21a8")
        tree.tag_configure("TECH", foreground="#6b21a8")
        tree.tag_configure("SRAM", foreground="#86198f")
        tree.tag_configure("SREW", foreground="#86198f")
        tree.tag_configure("FSC", foreground="#be185d")
        tree.tag_configure("FRA", foreground="#0369a1")
        tree.tag_configure("FSB", foreground="#15803d", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("SCRAP", foreground="#b91c1c", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("Reguler", foreground="#0284c7")
        tree.tag_configure("Overhaul", foreground="#d97706")
        tree.tag_configure("Periodic", foreground="#7c3aed")
        tree.tag_configure("Restoration", foreground="#be185d")
        self.bind_treeview_mousewheel(tree)

    # -------------------------------------------------------------------------
    # MAIN UI SHELL (RESPONSIVE SINGLE-SCREEN LAYOUT)
    # -------------------------------------------------------------------------
    def build_ui_layout(self):
        # 1. Left Sidebar (Collapsible: 200px <-> 62px)
        self.sidebar_expanded = True
        self.sidebar_width_expanded = 200
        self.sidebar_width_collapsed = 62

        self.sidebar = ctk.CTkFrame(self, width=self.sidebar_width_expanded, corner_radius=0, fg_color="#0b192c")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Brand Box in Sidebar with Toggle Button
        brand_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand_box.pack(fill="x", padx=10, pady=(16, 12))

        logo_path = get_resource_path("assets/gmf_logo.png")
        if not logo_path.exists():
            logo_path = BASE_DIR / "assets" / "gmf_logo.png"

        self.sidebar_logo_img = None
        if logo_path.exists():
            try:
                pil_img = Image.open(logo_path)
                self.sidebar_logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(42, 24))
                lbl_logo = ctk.CTkLabel(brand_box, image=self.sidebar_logo_img, text="")
                lbl_logo.pack(side="left", padx=(2, 6))
            except Exception:
                pass

        self.brand_text_box = ctk.CTkFrame(brand_box, fg_color="transparent")
        self.brand_text_box.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            self.brand_text_box,
            text="GMF AEROASIA",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#ffffff"
        ).pack(anchor="w")
        ctk.CTkLabel(
            self.brand_text_box,
            text="TCC-2 Battery Shop",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color="#94a3b8"
        ).pack(anchor="w")

        # Sidebar Collapse / Expand Toggle Button
        self.btn_toggle_sidebar = ctk.CTkButton(
            brand_box,
            text="☰",
            width=28,
            height=28,
            corner_radius=6,
            fg_color="transparent",
            hover_color="#1e293b",
            text_color="#cbd5e1",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.toggle_sidebar
        )
        self.btn_toggle_sidebar.pack(side="right", padx=2)

        # Sidebar Separator
        sep1 = ctk.CTkFrame(self.sidebar, height=1, fg_color="#1e293b")
        sep1.pack(fill="x", padx=10, pady=(0, 10))

        # Sidebar Navigation Menu Items
        self.menu_items_def = [
            ("Dashboard", "🏠  Dashboard", "🏠"),
            ("Daily Battery", "🔋  Daily Battery", "🔋"),
            ("History FSB", "📋  History FSB", "📋"),
            ("History SCRAP", "⚠️  History SCRAP", "⚠️"),
            ("Laporan", "📊  Laporan", "📊"),
            ("Pengaturan", "⚙️  Pengaturan", "⚙️")
        ]

        self.nav_menu_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.nav_menu_frame.pack(fill="x", padx=8, pady=2)

        for name, label_text, icon_char in self.menu_items_def:
            btn = ctk.CTkButton(
                self.nav_menu_frame,
                text=label_text,
                anchor="w",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold" if name == "Dashboard" else "normal"),
                fg_color="#2563eb" if name == "Dashboard" else "transparent",
                text_color="#ffffff" if name == "Dashboard" else "#cbd5e1",
                hover_color="#1e293b",
                height=36,
                corner_radius=7,
                command=lambda n=name: self.switch_view(n)
            )
            btn.pack(fill="x", pady=2)
            self.nav_buttons[name] = btn

        # Sidebar Bottom Slogan
        self.slogan_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.slogan_box.pack(side="bottom", fill="x", padx=8, pady=14)

        self.lbl_slogan = ctk.CTkLabel(
            self.slogan_box,
            text="✈ Safety • Quality • Reliability",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color="#94a3b8"
        )
        self.lbl_slogan.pack(anchor="center")

        # 2. Right Workspace Area (Single-Screen Layout - Direct Viewport)
        self.workspace_container = ctk.CTkFrame(self, fg_color=("#f8fafc", "#0f172a"), corner_radius=0)
        self.workspace_container.pack(side="right", fill="both", expand=True)

        # Header Bar (Direct child of workspace_container)
        self.build_header_bar(parent=self.workspace_container)

        # Page Container (Responsive to screen width & height)
        self.page_container = ctk.CTkFrame(self.workspace_container, fg_color="transparent")
        self.page_container.pack(fill="both", expand=True, padx=12, pady=(4, 2))

        # Bottom Status Bar (pinned at bottom of right area)
        self.build_status_bar()

        # Build Dashboard page first for fast startup (<2.5s)
        self.build_page_dashboard()

        # Show Dashboard initially
        self.switch_view("Dashboard")

    def toggle_sidebar(self):
        """Toggles between expanded (200px) and compact icon-only (62px) sidebar mode."""
        if self.sidebar_expanded:
            self.sidebar_expanded = False
            self.sidebar.configure(width=self.sidebar_width_collapsed)
            self.brand_text_box.pack_forget()
            for name, _, icon_char in self.menu_items_def:
                if name in self.nav_buttons:
                    self.nav_buttons[name].configure(text=icon_char, anchor="center")
            self.lbl_slogan.configure(text="✈")
        else:
            self.sidebar_expanded = True
            self.sidebar.configure(width=self.sidebar_width_expanded)
            self.brand_text_box.pack(side="left", fill="x", expand=True)
            for name, label_text, _ in self.menu_items_def:
                if name in self.nav_buttons:
                    self.nav_buttons[name].configure(text=label_text, anchor="w")
            self.lbl_slogan.configure(text="✈ Safety • Quality • Reliability")

    # -------------------------------------------------------------------------
    # HEADER BAR (Matching Image 3 & Compact for 1366x768 screens)
    # -------------------------------------------------------------------------
    def build_header_bar(self, parent=None):
        target_parent = parent if parent else self.workspace_container
        header = ctk.CTkFrame(target_parent, height=56, fg_color=("#ffffff", "#1e293b"), corner_radius=0, border_width=1, border_color=("#e2e8f0", "#334155"))
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        # Left Title Box
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left", padx=(12, 6), pady=8)

        accent_bar = ctk.CTkFrame(title_box, width=4, height=32, fg_color="#2563eb", corner_radius=2)
        accent_bar.pack(side="left", padx=(0, 8))

        text_box = ctk.CTkFrame(title_box, fg_color="transparent")
        text_box.pack(side="left")

        ctk.CTkLabel(
            text_box,
            text="Battery Shop TCC-2",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(anchor="w")

        ctk.CTkLabel(
            text_box,
            text="Aircraft Battery Monitoring • PT GMF AeroAsia Tbk",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color=("#64748b", "#94a3b8")
        ).pack(anchor="w")

        # Center: Date & Live Clock Pill
        clock_card = ctk.CTkFrame(
            header,
            fg_color=("#f1f5f9", "#0f172a"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            corner_radius=7
        )
        clock_card.pack(side="left", padx=8, pady=10)

        self.lbl_day_date = ctk.CTkLabel(
            clock_card,
            text="📅 Memuat...",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=("#0369a1", "#38bdf8")
        )
        self.lbl_day_date.pack(side="left", padx=8, pady=3)

        sep = ctk.CTkFrame(clock_card, width=1, height=16, fg_color=("#cbd5e1", "#334155"))
        sep.pack(side="left", pady=3)

        self.lbl_clock_time = ctk.CTkLabel(
            clock_card,
            text="🕒 --:--:-- WIB",
            font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        )
        self.lbl_clock_time.pack(side="left", padx=8, pady=3)

        # Right Action Buttons & Operator Chip
        action_box = ctk.CTkFrame(header, fg_color="transparent")
        action_box.pack(side="right", padx=10, pady=8)

        # Operator Profile Chip
        op_chip = ctk.CTkFrame(action_box, fg_color=("#f1f5f9", "#0f172a"), corner_radius=14, border_width=1, border_color=("#e2e8f0", "#334155"))
        op_chip.pack(side="left", padx=(0, 6), pady=2)
        ctk.CTkLabel(
            op_chip,
            text="👤 Operator",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=("#334155", "#cbd5e1")
        ).pack(padx=8, pady=3)

        # Dropdown Manager Button
        ctk.CTkButton(
            action_box,
            text="⚙️ Dropdown",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#6366f1",
            hover_color="#4f46e5",
            width=80,
            height=30,
            command=self.open_dropdown_manager
        ).pack(side="left", padx=2)

        # Calendar Filter Button
        self.btn_calendar = ctk.CTkButton(
            action_box,
            text="📅 Kalender",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            width=80,
            height=30,
            command=self.open_calendar_dialog
        )
        self.btn_calendar.pack(side="left", padx=2)

        # Input Unit Button
        ctk.CTkButton(
            action_box,
            text="➕ Input Unit",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            width=88,
            height=30,
            command=self.open_add_unit_dialog
        ).pack(side="left", padx=2)

        # Panduan Button
        ctk.CTkButton(
            action_box,
            text="📖 Panduan",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#475569",
            hover_color="#334155",
            width=76,
            height=30,
            command=self.open_glossary
        ).pack(side="left", padx=2)

        # Master Excel Options / Import Button
        self.btn_excel = ctk.CTkButton(
            action_box,
            text="📁 Excel ▾",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#10b981",
            hover_color="#059669",
            width=80,
            height=30,
            command=self.show_excel_action_menu
        )
        self.btn_excel.pack(side="left", padx=2)
        self.btn_excel.bind("<Button-3>", lambda e: self.show_excel_action_menu())

        # Refresh Data Button (Icon Only)
        ctk.CTkButton(
            action_box,
            text="🔄",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            width=34,
            height=30,
            command=lambda: self.load_data(silent=False)
        ).pack(side="left", padx=2)

    # -------------------------------------------------------------------------
    # STATUS BAR (Bottom)
    # -------------------------------------------------------------------------
    def build_status_bar(self):
        self.statusbar = ctk.CTkFrame(self.workspace_container, height=28, fg_color=("#ffffff", "#1e293b"), corner_radius=0, border_width=1, border_color=("#e2e8f0", "#334155"))
        self.statusbar.pack(side="bottom", fill="x")
        self.statusbar.pack_propagate(False)

        self.lbl_file_status = ctk.CTkLabel(
            self.statusbar,
            text=f"● Live Master: {self.excel_file_path.name}",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#10b981",
            cursor="hand2"
        )
        self.lbl_file_status.pack(side="left", padx=14, pady=2)
        self.lbl_file_status.bind("<Button-1>", lambda e: self.open_welcome_import_dialog())

        self.lbl_last_sync = ctk.CTkLabel(
            self.statusbar,
            text="Terakhir Update: -",
            font=ctk.CTkFont(size=10),
            text_color=("#64748b", "#94a3b8")
        )
        self.lbl_last_sync.pack(side="left", padx=12, pady=2)

        lbl_shortcuts = ctk.CTkLabel(
            self.statusbar,
            text="Pintasan: [Double-Click] Edit Sel / Dropdown | [Delete] Hapus Baris | [F5] Refresh | [F1] Glosarium",
            font=ctk.CTkFont(size=10),
            text_color=("#64748b", "#94a3b8")
        )
        lbl_shortcuts.pack(side="right", padx=14, pady=2)

    # -------------------------------------------------------------------------
    # VIEW SWITCHING ENGINE
    # -------------------------------------------------------------------------
    def switch_view(self, view_name: str):
        # Lazy initialization: build target page on first navigation
        if view_name not in self.pages:
            builders = {
                "Daily Battery": self.build_page_daily,
                "History FSB": self.build_page_fsb,
                "History SCRAP": self.build_page_scrap,
                "Laporan": self.build_page_laporan,
                "Pengaturan": self.build_page_pengaturan,
            }
            builder = builders.get(view_name)
            if builder:
                builder()
                # If data is already loaded, populate the newly built page immediately
                if hasattr(self, "df_daily") and not self.df_daily.empty:
                    if view_name == "Daily Battery":
                        self.filter_daily_table()
                    elif view_name == "History FSB":
                        self.update_kpi_cards()
                        self.filter_fsb_table()
                    elif view_name == "History SCRAP":
                        self.update_kpi_cards()
                        self.filter_scrap_table()
                    elif view_name == "Laporan":
                        self.render_laporan_charts()
                    elif view_name == "Pengaturan":
                        self.update_pengaturan_view()

        if view_name not in self.pages:
            return

        self.current_view_name = view_name

        # Highlight sidebar buttons
        for name, btn in self.nav_buttons.items():
            if name == view_name:
                btn.configure(fg_color="#2563eb", text_color="#ffffff", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"))
            else:
                btn.configure(fg_color="transparent", text_color="#cbd5e1", font=ctk.CTkFont(family="Segoe UI", size=12, weight="normal"))

        # Raise target page
        for name, page in self.pages.items():
            if name == view_name:
                page.pack(fill="both", expand=True)
            else:
                page.pack_forget()

        # Trigger specific page refreshes if needed (only if data already loaded)
        if view_name == "Dashboard":
            if hasattr(self, "df_daily") and not self.df_daily.empty:
                self.after(20, self.render_dashboard_charts)
        elif view_name == "Laporan":
            if hasattr(self, "df_fsb") and not self.df_fsb.empty:
                self.after(20, self.render_laporan_charts)

    # -------------------------------------------------------------------------
    # PAGE 1: DASHBOARD (EXACT VISUAL REFERENCE - Image 3)
    # -------------------------------------------------------------------------
    def build_page_dashboard(self):
        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["Dashboard"] = page

        # 1. Aviation Banner ("Reliable Battery, Safe Flight")
        banner_card = ctk.CTkFrame(page, height=92, corner_radius=10, fg_color="#559eda")
        banner_card.pack(fill="x", pady=(0, 8))
        banner_card.pack_propagate(False)

        banner_img_path = get_resource_path("assets/banner_composed.png")
        if not banner_img_path.exists():
            banner_img_path = BASE_DIR / "assets" / "banner_composed.png"

        self.banner_img = None
        self._banner_resize_job = None
        self._last_banner_w = 0

        if banner_img_path.exists():
            try:
                pil_b = Image.open(banner_img_path)
                self.banner_img = ctk.CTkImage(light_image=pil_b, dark_image=pil_b, size=(960, 92))
                self.lbl_banner = ctk.CTkLabel(banner_card, image=self.banner_img, text="", corner_radius=10)
                self.lbl_banner.pack(fill="both", expand=True)

                def _debounced_banner_resize():
                    self._banner_resize_job = None
                    if not hasattr(self, "lbl_banner") or not self.lbl_banner:
                        return
                    w_device = banner_card.winfo_width()
                    if w_device < 100:
                        return
                    scaling = ctk.ScalingTracker.get_widget_scaling(self)
                    logical_w = max(400, int(w_device / scaling))
                    if abs(logical_w - self._last_banner_w) < 10:
                        return
                    self._last_banner_w = logical_w
                    new_img = self._create_responsive_banner_image(w_device, int(92 * scaling))
                    self.banner_img.configure(light_image=new_img, dark_image=new_img, size=(logical_w, 92))

                def _on_banner_configure(event):
                    if event.width > 100:
                        if self._banner_resize_job is not None:
                            try:
                                self.after_cancel(self._banner_resize_job)
                            except Exception:
                                pass
                        self._banner_resize_job = self.after(30, _debounced_banner_resize)

                banner_card.bind("<Configure>", _on_banner_configure)
                self.after(60, _debounced_banner_resize)
            except Exception as e:
                print(f"[WARN] Gagal memuat banner: {e}")

        # 2. Four KPI Cards
        kpi_row = ctk.CTkFrame(page, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 10))
        for i in range(4):
            kpi_row.columnconfigure(i, weight=1, uniform="kpi")

        self.kpi_dash_wip = self._create_kpi_card(
            kpi_row, 0, "TOTAL UNIT WIP", "28 Unit", "↑ +2 dari kemarin", "Seluruh baterai aktif di bengkel",
            icon_text="⚡", icon_bg="#e0f2fe", icon_fg="#0284c7"
        )
        self.kpi_dash_ur = self._create_kpi_card(
            kpi_row, 1, "UNDER REPAIR (UR)", "8 Unit", "↑ +1 dari kemarin", "Sedang dalam penanganan aktif",
            icon_text="🔧", icon_bg="#dcfce7", icon_fg="#16a34a"
        )
        self.kpi_dash_wr = self._create_kpi_card(
            kpi_row, 2, "WAITING REPAIR (WR)", "20 Unit", "↓ -1 dari kemarin", "Antre alokasi meja/station",
            icon_text="⏱️", icon_bg="#fef3c7", icon_fg="#d97706", val_color="#d97706"
        )
        self.kpi_dash_hold = self._create_kpi_card(
            kpi_row, 3, "HOLD / EVALUASI", "0 Unit", "↓ 0 dari kemarin", "WI, TECH, FSC, SRAM, SREW",
            icon_text="🛡", icon_bg="#f3e8ff", icon_fg="#7c3aed", val_color="#7c3aed"
        )

        # 3. Middle Section (2 Columns: Stacked Bar Chart & Donut Chart)
        middle_row = ctk.CTkFrame(page, fg_color="transparent")
        middle_row.pack(fill="x", pady=(0, 10))
        middle_row.columnconfigure(0, weight=6)
        middle_row.columnconfigure(1, weight=4)

        # Col 0: Tren Unit per Status (7 Hari Terakhir)
        chart1_card = ctk.CTkFrame(middle_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        chart1_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(
            chart1_card,
            text="Tren Unit per Status (7 Hari Terakhir)",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(anchor="w", padx=14, pady=(10, 2))

        self.fig_trend = Figure(figsize=(5.5, 2.4), dpi=95)
        self.fig_trend.patch.set_facecolor("#ffffff")
        self.canvas_trend = FigureCanvasTkAgg(self.fig_trend, master=chart1_card)
        self.canvas_trend.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Col 1: Distribusi Status Unit (Donut Chart)
        chart2_card = ctk.CTkFrame(middle_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        chart2_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(
            chart2_card,
            text="Distribusi Status Unit",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(anchor="w", padx=14, pady=(10, 2))

        self.fig_donut = Figure(figsize=(4.2, 2.4), dpi=95)
        self.fig_donut.patch.set_facecolor("#ffffff")
        self.canvas_donut = FigureCanvasTkAgg(self.fig_donut, master=chart2_card)
        self.canvas_donut.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # 4. Bottom Section: Data Unit Battery Table
        table_card = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        table_card.pack(fill="both", expand=True, pady=(0, 4))

        tbl_header = ctk.CTkFrame(table_card, fg_color="transparent")
        tbl_header.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(
            tbl_header,
            text="Data Unit Battery",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(side="left")

        self.lbl_dash_count = ctk.CTkLabel(
            tbl_header,
            text="Total: 0 Unit",
            font=ctk.CTkFont(size=11),
            text_color=("#64748b", "#94a3b8")
        )
        self.lbl_dash_count.pack(side="left", padx=12)

        # Search Box on Right
        search_box = ctk.CTkFrame(tbl_header, fg_color="transparent")
        search_box.pack(side="right")

        entry_search = ctk.CTkEntry(
            search_box,
            placeholder_text="🔍 Cari nomor unit, P/N, customer...",
            textvariable=self.daily_search_var,
            width=260,
            height=32
        )
        entry_search.pack(side="left", padx=(0, 6))
        entry_search.bind("<KeyRelease>", lambda e: self.filter_daily_table())

        btn_rst = ctk.CTkButton(
            search_box,
            text="Reset",
            width=65,
            height=32,
            fg_color="#94a3b8",
            hover_color="#64748b",
            command=self.reset_daily_filters
        )
        btn_rst.pack(side="left")

        # Treeview for Dashboard
        tbl_container = ctk.CTkFrame(table_card, fg_color="transparent")
        tbl_container.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        cols = ["WO", "Part Number", "Serial Number", "Customer", "Status", "Progress", "PIR", "TAT", "Position", "Daily", "Estimasi FSB", "Remark"]
        self.tree_dash = ttk.Treeview(tbl_container, columns=cols, show="headings", style="Custom.Treeview", height=10)

        col_widths = {
            "WO": 105, "Part Number": 105, "Serial Number": 115, "Customer": 160,
            "Status": 90, "Progress": 110, "PIR": 115, "TAT": 55, "Position": 75,
            "Daily": 60, "Estimasi FSB": 90, "Remark": 150
        }
        for c in cols:
            self.tree_dash.heading(c, text=c, command=lambda _c=c: self.sort_treeview(self.tree_dash, _c, False))
            self.tree_dash.column(c, width=col_widths.get(c, 90), anchor="center" if c in ["WO", "Part Number", "Serial Number", "Status", "PIR", "TAT", "Position", "Daily", "Estimasi FSB"] else "w")

        scroll_y = ttk.Scrollbar(tbl_container, orient="vertical", command=self.tree_dash.yview)
        scroll_x = ttk.Scrollbar(tbl_container, orient="horizontal", command=self.tree_dash.xview)
        self.tree_dash.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_dash.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tbl_container.rowconfigure(0, weight=1)
        tbl_container.columnconfigure(0, weight=1)

        self.setup_inline_editing(self.tree_dash, "daily")
        self.apply_treeview_tags(self.tree_dash)

    def _create_kpi_card(self, parent, col, title, initial_val, delta_str, subtitle, icon_text="🔋", icon_bg="#e0f2fe", icon_fg="#0284c7", val_color="#0f172a"):
        card = ctk.CTkFrame(parent, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        card.grid(row=0, column=col, sticky="nsew", padx=4, pady=2)

        top_row = ctk.CTkFrame(card, fg_color="transparent")
        top_row.pack(fill="x", padx=14, pady=(12, 4))

        # Icon badge with exact centered placement
        icon_badge = ctk.CTkFrame(top_row, width=34, height=34, corner_radius=8, fg_color=icon_bg)
        icon_badge.pack(side="left", padx=(0, 10))
        icon_badge.pack_propagate(False)
        lbl_icon = ctk.CTkLabel(icon_badge, text=icon_text, font=ctk.CTkFont(family="Segoe UI Emoji", size=14))
        lbl_icon.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            top_row,
            text=title,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#64748b", "#94a3b8")
        ).pack(side="left", anchor="w")

        lbl_val = ctk.CTkLabel(
            card,
            text=initial_val,
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color=val_color
        )
        lbl_val.pack(anchor="w", padx=14, pady=(2, 2))

        bottom_row = ctk.CTkFrame(card, fg_color="transparent")
        bottom_row.pack(fill="x", padx=14, pady=(0, 12))

        lbl_delta = ctk.CTkLabel(
            bottom_row,
            text=delta_str,
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#16a34a" if "↑" in delta_str or "+" in delta_str else ("#d97706" if "-" in delta_str else "#64748b")
        )
        lbl_delta.pack(side="left", padx=(0, 6))

        lbl_sub = ctk.CTkLabel(
            bottom_row,
            text=subtitle,
            font=ctk.CTkFont(size=10),
            text_color=("#64748b", "#94a3b8")
        )
        lbl_sub.pack(side="left")

        card.lbl_val = lbl_val
        card.lbl_delta = lbl_delta
        return card

    def _create_responsive_banner_image(self, target_w: int, target_h: int = 92) -> Image.Image:
        """Dynamically composes banner for exact device width/height, preserving airplane & text aspect ratios."""
        target_w = max(600, int(target_w))
        target_h = max(60, int(target_h or 92))
        cache_key = (target_w, target_h)
        if not hasattr(self, "_banner_cache"):
            self._banner_cache = {}
        if cache_key in self._banner_cache:
            return self._banner_cache[cache_key]

        raw_airplane = BASE_DIR / "assets" / "banner_airplane_raw.jpg"
        if not raw_airplane.exists():
            raw_airplane = get_resource_path("assets/banner_airplane_raw.jpg")
        if not raw_airplane.exists():
            static_banner = get_resource_path("assets/banner_composed.png")
            return Image.open(static_banner)

        try:
            scale_factor = target_h / 92.0
            raw = Image.open(raw_airplane).convert("RGBA")
            plane_crop = raw.crop((25, 180, 1345, 560))
            pc_w, pc_h = plane_crop.size
            pw = int(pc_w * (target_h / float(pc_h)))
            plane_scaled = plane_crop.resize((pw, target_h), Image.Resampling.LANCZOS)

            mask = Image.new("L", (pw, target_h), 255)
            fade_w = int(95 * scale_factor)
            for x in range(fade_w):
                a = int(255 * ((x / float(fade_w)) ** 1.4))
                for y in range(target_h):
                    mask.putpixel((x, y), a)
            fade_r = int(25 * scale_factor)
            for x in range(fade_r):
                a = int(255 * ((x / float(fade_r)) ** 1.4))
                for y in range(target_h):
                    cur = mask.getpixel((pw - 1 - x, y))
                    mask.putpixel((pw - 1 - x, y), min(cur, a))
            for y in range(int(5 * scale_factor)):
                a_edge = int(255 * (y / float(max(1, 5 * scale_factor))))
                for x in range(pw):
                    mask.putpixel((x, y), min(mask.getpixel((x, y)), a_edge))
                    mask.putpixel((x, target_h - 1 - y), min(mask.getpixel((x, target_h - 1 - y)), a_edge))

            banner = Image.new("RGBA", (target_w, target_h), (85, 158, 218, 255))
            draw = ImageDraw.Draw(banner)
            for y in range(target_h):
                ratio = y / float(target_h)
                r = int(78 + (156 - 78) * (ratio ** 1.3))
                g = int(152 + (186 - 152) * (ratio ** 1.3))
                b = int(214 + (218 - 214) * (ratio ** 1.3))
                draw.line([(0, y), (target_w, y)], fill=(r, g, b, 255))

            plane_x = target_w - pw - int(25 * scale_factor)
            banner.paste(plane_scaled, (plane_x, 0), mask)

            font_title = None
            font_sub = None
            for f in ["georgiab.ttf", "georgia.ttf", "segoeuib.ttf", "arialbd.ttf"]:
                try:
                    font_title = ImageFont.truetype(f, int(20 * scale_factor))
                    font_sub = ImageFont.truetype("segoeuib.ttf", int(10.5 * scale_factor))
                    break
                except Exception:
                    pass
            if not font_title:
                font_title = ImageFont.load_default()
                font_sub = ImageFont.load_default()

            title_text = "\u201cReliable Battery, Safe Flight\u201d"
            sub_text = "PT GMF AEROASIA Tbk  \u2022  Aircraft Battery Maintenance Monitoring System  \u2022  TCC-2"

            tx = int(26 * scale_factor)
            ty = int(18 * scale_factor)
            sub_y = ty + int(36 * scale_factor)

            draw.text((tx + 1, ty + 1), title_text, fill=(255, 255, 255, 230), font=font_title)
            draw.text((tx, ty), title_text, fill=(11, 28, 56, 255), font=font_title)
            draw.text((tx + 1, sub_y + 1), sub_text, fill=(255, 255, 255, 200), font=font_sub)
            draw.text((tx, sub_y), sub_text, fill=(24, 52, 98, 255), font=font_sub)

            res_img = banner.convert("RGB")
            self._banner_cache[cache_key] = res_img
            return res_img
        except Exception:
            static_banner = get_resource_path("assets/banner_composed.png")
            fallback_img = Image.open(static_banner)
            self._banner_cache[cache_key] = fallback_img
            return fallback_img

    # -------------------------------------------------------------------------
    # PAGE 2: DAILY BATTERY (DEDICATED FULL WIP TABLE)
    # -------------------------------------------------------------------------
    def build_page_daily(self):
        page = ctk.CTkFrame(self.page_container, fg_color="transparent")
        self.pages["Daily Battery"] = page

        # Action Toolbar with Automatic Status Routing Buttons
        action_bar = ctk.CTkFrame(page, height=44, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        action_bar.pack(fill="x", pady=(0, 8))
        action_bar.pack_propagate(False)

        ctk.CTkButton(
            action_bar, text="➕ Tambah Unit", fg_color="#059669", hover_color="#047857",
            width=110, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=self.open_add_unit_dialog
        ).pack(side="left", padx=(10, 3), pady=6)

        ctk.CTkButton(
            action_bar, text="✏️ Edit Unit", fg_color="#0284c7", hover_color="#0369a1",
            width=95, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=self.open_selected_daily_form
        ).pack(side="left", padx=3, pady=6)

        ctk.CTkButton(
            action_bar, text="⚡ Final Charge", fg_color="#10b981", hover_color="#059669",
            width=110, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=self.prompt_route_selected_to_fsb
        ).pack(side="left", padx=3, pady=6)

        ctk.CTkButton(
            action_bar, text="⚠️ Pindahkan SCRAP", fg_color="#dc2626", hover_color="#b91c1c",
            width=130, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=self.prompt_route_selected_to_scrap
        ).pack(side="left", padx=3, pady=6)

        ctk.CTkButton(
            action_bar, text="🗑️ Hapus Baris", fg_color="#ef4444", hover_color="#dc2626",
            width=95, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=self.delete_selected_daily_unit
        ).pack(side="left", padx=3, pady=6)

        ctk.CTkButton(
            action_bar, text="🔄 Refresh", fg_color="#2563eb", hover_color="#1d4ed8",
            width=85, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self.load_data(silent=False)
        ).pack(side="left", padx=3, pady=6)

        lbl_guide = ctk.CTkLabel(
            action_bar,
            text="💡 Tip: Double-click sel untuk edit langsung. Status FINAL CHARGE & SCRAP akan otomatis dipindahkan.",
            font=ctk.CTkFont(size=10),
            text_color=("#0284c7", "#38bdf8")
        )
        lbl_guide.pack(side="right", padx=14, pady=6)

        # Full Table Frame
        tbl_frame = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        tbl_frame.pack(fill="both", expand=True)

        cols = ["WO", "Part Number", "Serial Number", "Customer", "Status", "Progress", "PIR", "TAT", "Position", "Daily", "Estimasi FSB", "Remark"]
        self.tree_daily = ttk.Treeview(tbl_frame, columns=cols, show="headings", style="Custom.Treeview")

        col_widths = {
            "WO": 110, "Part Number": 110, "Serial Number": 120, "Customer": 170,
            "Status": 95, "Progress": 120, "PIR": 125, "TAT": 60, "Position": 80,
            "Daily": 65, "Estimasi FSB": 95, "Remark": 160
        }
        for c in cols:
            self.tree_daily.heading(c, text=c, command=lambda _c=c: self.sort_treeview(self.tree_daily, _c, False))
            self.tree_daily.column(c, width=col_widths.get(c, 100), anchor="center" if c in ["WO", "Part Number", "Serial Number", "Status", "PIR", "TAT", "Position", "Daily", "Estimasi FSB"] else "w")

        sc_y = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree_daily.yview)
        sc_x = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree_daily.xview)
        self.tree_daily.configure(yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)

        self.tree_daily.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=(4, 0))
        sc_y.grid(row=0, column=1, sticky="ns", pady=(4, 0))
        sc_x.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        self.setup_inline_editing(self.tree_daily, "daily")
        self.apply_treeview_tags(self.tree_daily)

    # -------------------------------------------------------------------------
    # PAGE 3: HISTORY FSB (SERVICEABLE HISTORY)
    # -------------------------------------------------------------------------
    def build_page_fsb(self):
        page = ctk.CTkFrame(self.page_container, fg_color="transparent")
        self.pages["History FSB"] = page

        # KPI Row
        kpi_row = ctk.CTkFrame(page, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 8))
        for i in range(3):
            kpi_row.columnconfigure(i, weight=1, uniform="fsb_kpi")

        self.card_fsb_total = self._create_kpi_card(
            kpi_row, 0, "TOTAL BATERAI SERVICEABLE (FSB)", "146 Unit", "Lolos Uji CMM", "Siap dipasang ke armada",
            icon_text="✅", icon_bg="#dcfce7", icon_fg="#16a34a", val_color="#15803d"
        )
        self.card_fsb_tat = self._create_kpi_card(
            kpi_row, 1, "RATA-RATA TAT (HARI KERJA)", "10.2 Hari", "Target Standar <= 10 hari", "Efisiensi pengerjaan bengkel",
            icon_text="⚡", icon_bg="#e0f2fe", icon_fg="#0284c7"
        )
        self.card_fsb_cust = self._create_kpi_card(
            kpi_row, 2, "MASKAPAI DOMINAN", "PT Garuda Indonesia", "Distribusi Tertinggi", "Pelanggan terbanyak diservis",
            icon_text="✈️", icon_bg="#f3e8ff", icon_fg="#7c3aed", val_color="#7c3aed"
        )

        # Toolbar & Search
        tb = ctk.CTkFrame(page, height=44, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        tb.pack(fill="x", pady=(0, 8))
        tb.pack_propagate(False)

        ctk.CTkLabel(tb, text="Pencarian Cepat:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(12, 6))
        e_search = ctk.CTkEntry(tb, textvariable=self.fsb_search_var, placeholder_text="Ketik nomor WO, S/N, P/N, Customer...", width=260, height=30)
        e_search.pack(side="left", padx=4)
        e_search.bind("<KeyRelease>", lambda e: self.filter_fsb_table())

        ctk.CTkLabel(tb, text="Urutan:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(12, 4))
        cb_sort = ctk.CTkComboBox(tb, variable=self.fsb_sort_order, values=["Terbaru (Date FSB Desc)", "Terlama (Date FSB Asc)", "WO Asc", "WO Desc"], width=190, height=30, command=lambda v: self.filter_fsb_table())
        cb_sort.pack(side="left", padx=4)

        self.lbl_fsb_count = ctk.CTkLabel(tb, text="Menampilkan: 0 Unit", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"))
        self.lbl_fsb_count.pack(side="right", padx=14)

        ctk.CTkButton(tb, text="➕ Tambah Data FSB", fg_color="#059669", hover_color="#047857", width=130, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.open_add_fsb_dialog).pack(side="left", padx=(12, 4))
        ctk.CTkButton(tb, text="✏️ Edit Catatan", fg_color="#0284c7", hover_color="#0369a1", width=105, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.open_selected_fsb_form).pack(side="left", padx=4)
        ctk.CTkButton(tb, text="🗑️ Hapus Baris", fg_color="#ef4444", hover_color="#dc2626", width=100, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.delete_selected_fsb_unit).pack(side="left", padx=4)

        # Table
        tbl_frame = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        tbl_frame.pack(fill="both", expand=True)

        cols = ["Date FSB", "WO", "Part Number", "Serial Number", "Customer", "Status", "Progress", "TAT", "Position", "Remark"]
        self.tree_fsb = ttk.Treeview(tbl_frame, columns=cols, show="headings", style="Custom.Treeview")

        col_widths = {
            "Date FSB": 125, "WO": 110, "Part Number": 110, "Serial Number": 120,
            "Customer": 170, "Status": 95, "Progress": 120, "TAT": 60, "Position": 80, "Remark": 200
        }
        for c in cols:
            self.tree_fsb.heading(c, text=c, command=lambda _c=c: self.sort_treeview(self.tree_fsb, _c, False))
            self.tree_fsb.column(c, width=col_widths.get(c, 100), anchor="center" if c in ["Date FSB", "WO", "Part Number", "Serial Number", "Status", "TAT", "Position"] else "w")

        sc_y = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree_fsb.yview)
        sc_x = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree_fsb.xview)
        self.tree_fsb.configure(yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)

        self.tree_fsb.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=(4, 0))
        sc_y.grid(row=0, column=1, sticky="ns", pady=(4, 0))
        sc_x.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        self.setup_inline_editing(self.tree_fsb, "fsb")
        self.apply_treeview_tags(self.tree_fsb)

    # -------------------------------------------------------------------------
    # PAGE 4: HISTORY SCRAP (SCRAP & FRA HISTORY)
    # -------------------------------------------------------------------------
    def build_page_scrap(self):
        page = ctk.CTkFrame(self.page_container, fg_color="transparent")
        self.pages["History SCRAP"] = page

        # KPI Row
        kpi_row = ctk.CTkFrame(page, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 8))
        for i in range(3):
            kpi_row.columnconfigure(i, weight=1, uniform="scrap_kpi")

        self.card_scrap_total = self._create_kpi_card(
            kpi_row, 0, "TOTAL BATERAI SCRAP & FRA", "16 Unit", "Afkir / Retur As-Is", "Total baterai tidak diservis",
            icon_text="⚠️", icon_bg="#fee2e2", icon_fg="#dc2626", val_color="#dc2626"
        )
        self.card_scrap_only = self._create_kpi_card(
            kpi_row, 1, "AFKIR PERMANEN (SCRAP)", "14 Unit", "Beyond Economical Repair", "Kerusakan sel / thermal hazard",
            icon_text="❌", icon_bg="#fee2e2", icon_fg="#dc2626"
        )
        self.card_fra_only = self._create_kpi_card(
            kpi_row, 2, "FINISH RETURN AS IS (FRA)", "2 Unit", "Retur Tanpa Servis", "Permintaan khusus maskapai",
            icon_text="↩️", icon_bg="#e0f2fe", icon_fg="#0284c7"
        )

        # Toolbar
        tb = ctk.CTkFrame(page, height=44, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        tb.pack(fill="x", pady=(0, 8))
        tb.pack_propagate(False)

        ctk.CTkLabel(tb, text="Pencarian Cepat:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(12, 6))
        e_search = ctk.CTkEntry(tb, textvariable=self.scrap_search_var, placeholder_text="Ketik nomor WO, S/N, P/N, Customer...", width=260, height=30)
        e_search.pack(side="left", padx=4)
        e_search.bind("<KeyRelease>", lambda e: self.filter_scrap_table())

        self.lbl_scrap_count = ctk.CTkLabel(tb, text="Menampilkan: 0 Unit", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"))
        self.lbl_scrap_count.pack(side="right", padx=14)

        ctk.CTkButton(tb, text="➕ Tambah Data SCRAP", fg_color="#059669", hover_color="#047857", width=140, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.open_add_scrap_dialog).pack(side="left", padx=(12, 4))
        ctk.CTkButton(tb, text="✏️ Edit Catatan", fg_color="#0284c7", hover_color="#0369a1", width=105, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.open_selected_scrap_form).pack(side="left", padx=4)
        ctk.CTkButton(tb, text="🗑️ Hapus Baris", fg_color="#ef4444", hover_color="#dc2626", width=100, height=30, font=ctk.CTkFont(size=11, weight="bold"), command=self.delete_selected_scrap_unit).pack(side="left", padx=4)

        # Table
        tbl_frame = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        tbl_frame.pack(fill="both", expand=True)

        cols = ["Date SCRAP", "WO", "Part Number", "Serial Number", "Customer", "Status", "TAT", "Position", "Remark"]
        self.tree_scrap = ttk.Treeview(tbl_frame, columns=cols, show="headings", style="Custom.Treeview")

        col_widths = {
            "Date SCRAP": 125, "WO": 110, "Part Number": 110, "Serial Number": 120,
            "Customer": 170, "Status": 95, "TAT": 60, "Position": 80, "Remark": 220
        }
        for c in cols:
            self.tree_scrap.heading(c, text=c, command=lambda _c=c: self.sort_treeview(self.tree_scrap, _c, False))
            self.tree_scrap.column(c, width=col_widths.get(c, 100), anchor="center" if c in ["Date SCRAP", "WO", "Part Number", "Serial Number", "Status", "TAT", "Position"] else "w")

        sc_y = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree_scrap.yview)
        sc_x = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree_scrap.xview)
        self.tree_scrap.configure(yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)

        self.tree_scrap.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=(4, 0))
        sc_y.grid(row=0, column=1, sticky="ns", pady=(4, 0))
        sc_x.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        self.setup_inline_editing(self.tree_scrap, "scrap")
        self.apply_treeview_tags(self.tree_scrap)

    # -------------------------------------------------------------------------
    # PAGE 5: LAPORAN (WEEKLY & MONTHLY MAINTENANCE COMPLETION REPORTS)
    # -------------------------------------------------------------------------
    def build_page_laporan(self):
        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["Laporan"] = page

        # Top Control Bar
        ctrl_card = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        ctrl_card.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(ctrl_card, text="📊 Laporan Performa & Penyelesaian Maintenance Baterai", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=16, pady=10)

        ctk.CTkLabel(ctrl_card, text="Tahun:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(14, 4))
        self.cb_lap_year = ctk.CTkComboBox(ctrl_card, variable=self.laporan_year_var, values=["2026", "2025"], width=90, height=30, command=lambda v: self.render_laporan_charts())
        self.cb_lap_year.pack(side="left", padx=4)

        ctk.CTkLabel(ctrl_card, text="Minggu ke-:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(14, 4))
        week_options = [str(w) for w in range(1, 53)]
        self.cb_lap_week = ctk.CTkComboBox(ctrl_card, variable=self.laporan_week_var, values=week_options, width=80, height=30, command=lambda v: self.render_laporan_charts())
        self.cb_lap_week.pack(side="left", padx=4)

        btn_export = ctk.CTkButton(
            ctrl_card,
            text="📥 Export Laporan Excel",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=32,
            command=self.export_laporan_to_excel
        )
        btn_export.pack(side="right", padx=16, pady=10)

        # KPI Summary Row
        kpi_row = ctk.CTkFrame(page, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 10))
        for i in range(4):
            kpi_row.columnconfigure(i, weight=1, uniform="lap_kpi")

        self.card_lap_yearly = self._create_kpi_card(
            kpi_row, 0, "TOTAL SELESAI TAHUNAN", "0 Unit", "Tercatat di History FSB & SCRAP", "Akumulasi pekerjaan selesai",
            icon_text="🏆", icon_bg="#e0f2fe", icon_fg="#0284c7"
        )
        self.card_lap_weekly = self._create_kpi_card(
            kpi_row, 1, "TOTAL SELESAI MINGGU INI", "0 Unit", "Periode Senin - Minggu", "Penyelesaian minggu terpilih",
            icon_text="📅", icon_bg="#dcfce7", icon_fg="#16a34a"
        )
        self.card_lap_tat = self._create_kpi_card(
            kpi_row, 2, "RATA-RATA TAT PENYELESAIAN", "0 Hari", "Waktu perbaikan aktual", "Standar bengkel <= 10 hari",
            icon_text="⚡", icon_bg="#fef3c7", icon_fg="#d97706"
        )
        self.card_lap_ratio = self._create_kpi_card(
            kpi_row, 3, "SERVICEABLE SUCCESS RATE", "0%", "Rasio FSB vs Total Selesai", "Kualitas perbaikan bengkel",
            icon_text="📈", icon_bg="#f3e8ff", icon_fg="#7c3aed"
        )

        # 2 Charts Row: Weekly (Senin-Minggu) & Monthly (Jan-Des)
        chart_row = ctk.CTkFrame(page, fg_color="transparent")
        chart_row.pack(fill="x", pady=(0, 10))
        chart_row.columnconfigure(0, weight=1)
        chart_row.columnconfigure(1, weight=1)

        # Weekly Maintenance Chart Card
        c1 = ctk.CTkFrame(chart_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self.lbl_weekly_title = ctk.CTkLabel(c1, text="Weekly Battery Maintenance (Senin - Minggu)", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_weekly_title.pack(anchor="w", padx=14, pady=(10, 2))

        self.fig_weekly = Figure(figsize=(5, 2.6), dpi=95)
        self.fig_weekly.patch.set_facecolor("#ffffff")
        self.canvas_weekly = FigureCanvasTkAgg(self.fig_weekly, master=c1)
        self.canvas_weekly.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Monthly Maintenance Chart Card
        c2 = ctk.CTkFrame(chart_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c2.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        self.lbl_monthly_title = ctk.CTkLabel(c2, text="Monthly Battery Maintenance (Januari - Desember)", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_monthly_title.pack(anchor="w", padx=14, pady=(10, 2))

        self.fig_monthly = Figure(figsize=(5, 2.6), dpi=95)
        self.fig_monthly.patch.set_facecolor("#ffffff")
        self.canvas_monthly = FigureCanvasTkAgg(self.fig_monthly, master=c2)
        self.canvas_monthly.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Customer & Part Number Performance Tables
        tables_row = ctk.CTkFrame(page, fg_color="transparent")
        tables_row.pack(fill="x", pady=(0, 6))
        tables_row.columnconfigure(0, weight=1)
        tables_row.columnconfigure(1, weight=1)

        # Table 1: Customer Volume
        t1_card = ctk.CTkFrame(tables_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        t1_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(t1_card, text="Distribusi Volume Per Customer / Maskapai", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=14, pady=(10, 4))
        self.tree_cust = ttk.Treeview(t1_card, columns=["Customer", "Volume", "Avg TAT"], show="headings", style="Custom.Treeview", height=6)
        self.tree_cust.heading("Customer", text="Customer / Maskapai")
        self.tree_cust.heading("Volume", text="Jumlah Unit")
        self.tree_cust.heading("Avg TAT", text="Rata-rata TAT")
        self.tree_cust.column("Customer", width=220, anchor="w")
        self.tree_cust.column("Volume", width=100, anchor="center")
        self.tree_cust.column("Avg TAT", width=100, anchor="center")
        self.tree_cust.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        # Table 2: Part Number Volume
        t2_card = ctk.CTkFrame(tables_row, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        t2_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(t2_card, text="Distribusi Volume Per Part Number (P/N)", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=14, pady=(10, 4))
        self.tree_pn = ttk.Treeview(t2_card, columns=["Part Number", "Volume", "Avg TAT", "Customer Dominan"], show="headings", style="Custom.Treeview", height=6)
        self.tree_pn.heading("Part Number", text="Part Number")
        self.tree_pn.heading("Volume", text="Jumlah Unit")
        self.tree_pn.heading("Avg TAT", text="Avg TAT")
        self.tree_pn.heading("Customer Dominan", text="Customer Dominan")
        self.tree_pn.column("Part Number", width=120, anchor="center")
        self.tree_pn.column("Volume", width=90, anchor="center")
        self.tree_pn.column("Avg TAT", width=80, anchor="center")
        self.tree_pn.column("Customer Dominan", width=160, anchor="w")
        self.tree_pn.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    # -------------------------------------------------------------------------
    # PAGE 6: PENGATURAN (SETTINGS & WORKBOOK CONTROL)
    # -------------------------------------------------------------------------
    def build_page_pengaturan(self):
        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["Pengaturan"] = page

        # Card 1: Master Workbook Configuration
        c1 = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c1.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(c1, text="📁 Konfigurasi Master Excel Workbook", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(c1, text="Aplikasi hanya membaca SATU file Excel master sebagai sumber data utama.", font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=16, pady=(0, 10))

        self.lbl_cfg_file = ctk.CTkLabel(c1, text="", font=ctk.CTkFont(family="Consolas", size=11), justify="left")
        self.lbl_cfg_file.pack(anchor="w", padx=16, pady=(0, 12))

        btn_row1 = ctk.CTkFrame(c1, fg_color="transparent")
        btn_row1.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(btn_row1, text="📂 Buka Folder File di Explorer", fg_color="#334155", hover_color="#475569", height=34, command=self.open_excel_in_explorer).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row1, text="📁 Pilih / Ganti File Master Baru", fg_color="#2563eb", hover_color="#1d4ed8", height=34, command=self.open_welcome_import_dialog).pack(side="left")

        # Card 2: Backups Management
        c2 = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c2.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(c2, text="🔄 Cadangan Otomatis & Pemulihan (Backups)", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(c2, text="Cadangan timestamped otomatis dibuat sebelum setiap penulisan data ke Excel.", font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=16, pady=(0, 10))

        btn_row2 = ctk.CTkFrame(c2, fg_color="transparent")
        btn_row2.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(btn_row2, text="🔄 Pulihkan Data dari Cadangan...", fg_color="#d97706", hover_color="#b45309", height=34, command=self.open_restore_backup_modal).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row2, text="📂 Buka Folder Cadangan", fg_color="#334155", hover_color="#475569", height=34, command=self.open_backups_folder).pack(side="left")

        # Card 3: Dropdown Options
        c3 = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c3.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(c3, text="⚙️ Pilihan Dropdown Standar Bengkel", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(c3, text="Kelola opsi resmi Part Number, Customer, Position, dan Status.", font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=16, pady=(0, 10))

        ctk.CTkButton(c3, text="⚙️ Kelola Opsi Dropdown...", fg_color="#6366f1", hover_color="#4f46e5", height=34, command=self.open_dropdown_manager).pack(anchor="w", padx=16, pady=(0, 14))

        # Card 4: Appearance Theme
        c4 = ctk.CTkFrame(page, fg_color=("#ffffff", "#1e293b"), corner_radius=12, border_width=1, border_color=("#e2e8f0", "#334155"))
        c4.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(c4, text="🎨 Mode Tampilan", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 4))

        self.sw_theme = ctk.CTkSwitch(c4, text="Mode Gelap", font=ctk.CTkFont(size=11), command=self.toggle_dark_light_theme)
        self.sw_theme.pack(anchor="w", padx=16, pady=(0, 14))

    def update_pengaturan_view(self):
        stat = self.excel_file_path.stat() if self.excel_file_path.exists() else None
        dt_str = datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y, %H:%M:%S") if stat else "-"
        kb = round(stat.st_size / 1024, 1) if stat else 0
        txt = (
            f"• Nama File : {self.excel_file_path.name} ({kb} KB)\n"
            f"• Path Lengkap: {self.excel_file_path}\n"
            f"• Modifikasi: {dt_str}\n"
            f"• Status    : {'✔ Aktif & Terverifikasi' if self.excel_file_path.exists() else '❌ File Tidak Ada!'}"
        )
        self.lbl_cfg_file.configure(text=txt)

    def toggle_dark_light_theme(self):
        if self.sw_theme.get() == 1:
            ctk.set_appearance_mode("Dark")
            self.sw_theme.configure(text="Mode Gelap (Aktif)")
            if hasattr(self, "workspace_canvas"):
                self.workspace_canvas.configure(bg="#0f172a")
        else:
            ctk.set_appearance_mode("Light")
            self.sw_theme.configure(text="Mode Terang (Aktif)")
            if hasattr(self, "workspace_canvas"):
                self.workspace_canvas.configure(bg="#f8fafc")
        self.setup_styles()
        self.apply_treeview_tags(self.tree_dash)
        self.apply_treeview_tags(self.tree_daily)
        self.apply_treeview_tags(self.tree_fsb)
        self.apply_treeview_tags(self.tree_scrap)

    # -------------------------------------------------------------------------
    # STARTUP & DATA LOADING ENGINE
    # -------------------------------------------------------------------------
    def startup_check_and_load(self):
        """Verifies configured master file at startup; prompts Welcome dialog if missing or invalid."""
        if not self.excel_file_path.exists():
            self.open_welcome_import_dialog(force_select=True)
            return

        is_valid, errors, _ = validate_workbook_schema(self.excel_file_path)
        if not is_valid:
            self.open_welcome_import_dialog(force_select=True)
            return

        # Load valid data directly
        self.load_data(silent=True)

    def load_data(self, silent=False):
        """Pure read-only load from active master workbook on disk."""
        if not self.excel_repo.check_file_exists():
            messagebox.showerror(
                "File Master Tidak Ditemukan",
                f"File Excel master tidak ditemukan di:\n{self.excel_file_path}\n\nSilakan pilih file workbook yang sah.",
                parent=self
            )
            self.open_welcome_import_dialog(force_select=True)
            return

        try:
            data_dict, warnings = self.excel_repo.load_all_data()
            self.df_daily = data_dict.get("daily", pd.DataFrame())
            self.df_fsb = data_dict.get("fsb", pd.DataFrame())
            self.df_scrap = data_dict.get("scrap", pd.DataFrame())

            # Format date columns cleanly
            for df in [self.df_daily, self.df_fsb, self.df_scrap]:
                if not df.empty:
                    for col in ["PIR", "Date FSB", "Date SCRAP", "Estimasi FSB"]:
                        if col in df.columns:
                            df[col] = df[col].apply(format_date_with_month_name)

            # Auto-learn newly added customers or part numbers
            if not self.df_daily.empty:
                opts_changed = False
                for c in self.df_daily["Customer"].dropna().unique():
                    c_clean = str(c).strip()
                    if c_clean and c_clean not in self.dropdown_options["Customer"]:
                        self.dropdown_options["Customer"].append(c_clean)
                        opts_changed = True
                for p in self.df_daily["Part Number"].dropna().unique():
                    p_clean = str(p).strip()
                    if p_clean and p_clean not in self.dropdown_options["Part Number"]:
                        self.dropdown_options["Part Number"].append(p_clean)
                        opts_changed = True
                if opts_changed:
                    save_dropdown_options(self.dropdown_options)
                    self.update_combobox_options()

            # Record sync fingerprint
            self.last_sync_fingerprint = get_file_fingerprint(self.excel_file_path)

            # Update Year options in Laporan if built
            years = get_available_years(self.df_fsb, self.df_scrap)
            if hasattr(self, "cb_lap_year"):
                self.cb_lap_year.configure(values=[str(y) for y in years])

            # Refresh all UI representations
            self.update_kpi_cards()
            self.filter_daily_table()
            if "History FSB" in self.pages:
                self.filter_fsb_table()
            if "History SCRAP" in self.pages:
                self.filter_scrap_table()
            if not silent:
                self.render_dashboard_charts()
                if "Laporan" in self.pages:
                    self.render_laporan_charts()
            else:
                # Asynchronously render charts during initial load so table & KPIs appear instantly
                if self.current_view_name == "Dashboard":
                    self.after(30, self.render_dashboard_charts)
                elif self.current_view_name == "Laporan":
                    self.after(30, self.render_laporan_charts)
            if "Pengaturan" in self.pages and hasattr(self, "update_pengaturan_view"):
                self.update_pengaturan_view()

            now_str = datetime.now().strftime("%d %b %Y %H:%M:%S")
            self.lbl_file_status.configure(text=f"● Live Master: {self.excel_file_path.name}", text_color="#10b981")
            self.lbl_last_sync.configure(text=f"Terakhir Update: {now_str}")

            if not silent:
                messagebox.showinfo(
                    "Refresh Berhasil",
                    f"Data berhasil dibaca ulang dari file master:\n{self.excel_file_path.name}\n\n"
                    f"• Daily WIP   : {len(self.df_daily)} unit\n"
                    f"• History FSB : {len(self.df_fsb)} riwayat\n"
                    f"• History SCRAP: {len(self.df_scrap)} riwayat",
                    parent=self
                )

        except PermissionError:
            messagebox.showwarning(
                "File Excel Terkunci",
                f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel atau program lain.\n\nHarap tutup file tersebut di Excel terlebih dahulu, lalu klik 'Refresh Data'.",
                parent=self
            )
            self.lbl_file_status.configure(text=f"⚠ File Terkunci: {self.excel_file_path.name}", text_color="#ef4444")
        except Exception as e:
            messagebox.showerror(
                "Gagal Membaca Excel",
                f"Terjadi kesalahan saat membaca file Excel:\n{e}\n\nDashboard mempertahankan data sebelumnya yang valid.",
                parent=self
            )
            self.lbl_file_status.configure(text=f"✖ Gagal Muat: {self.excel_file_path.name}", text_color="#ef4444")

    # -------------------------------------------------------------------------
    # KPI CARDS UPDATE (Strictly Valid Excel Calculations)
    # -------------------------------------------------------------------------
    def update_kpi_cards(self):
        # 1. Daily WIP KPIs
        if not self.df_daily.empty and "Position" in self.df_daily.columns:
            total_wip = len(self.df_daily)
            ur_count = len(self.df_daily[self.df_daily["Position"].astype(str).str.strip() == "UR"])
            wr_count = len(self.df_daily[self.df_daily["Position"].astype(str).str.strip() == "WR"])
            hold_count = total_wip - (ur_count + wr_count)

            self.kpi_dash_wip.lbl_val.configure(text=f"{total_wip} Unit")
            self.kpi_dash_ur.lbl_val.configure(text=f"{ur_count} Unit")
            self.kpi_dash_wr.lbl_val.configure(text=f"{wr_count} Unit")
            self.kpi_dash_hold.lbl_val.configure(text=f"{hold_count} Unit")
        else:
            self.kpi_dash_wip.lbl_val.configure(text="0 Unit")
            self.kpi_dash_ur.lbl_val.configure(text="0 Unit")
            self.kpi_dash_wr.lbl_val.configure(text="0 Unit")
            self.kpi_dash_hold.lbl_val.configure(text="0 Unit")

        # 2. History FSB KPIs (if page is built)
        if hasattr(self, "card_fsb_total"):
            if not self.df_fsb.empty:
                total_fsb = len(self.df_fsb)
                tats = pd.to_numeric(self.df_fsb["TAT"], errors="coerce").dropna()
                avg_tat = round(tats.mean(), 1) if not tats.empty else 0.0
                top_cust = self.df_fsb["Customer"].mode()[0] if "Customer" in self.df_fsb and not self.df_fsb["Customer"].empty else "-"

                self.card_fsb_total.lbl_val.configure(text=f"{total_fsb} Unit")
                self.card_fsb_tat.lbl_val.configure(text=f"{avg_tat} Hari")
                self.card_fsb_cust.lbl_val.configure(text=str(top_cust))
            else:
                self.card_fsb_total.lbl_val.configure(text="0 Unit")
                self.card_fsb_tat.lbl_val.configure(text="0 Hari")
                self.card_fsb_cust.lbl_val.configure(text="-")

        # 3. History SCRAP KPIs (if page is built)
        if hasattr(self, "card_scrap_total"):
            if not self.df_scrap.empty and "Position" in self.df_scrap.columns:
                total_scrap = len(self.df_scrap)
                scrap_only = len(self.df_scrap[self.df_scrap["Position"].astype(str).str.strip() == "SCRAP"])
                fra_only = len(self.df_scrap[self.df_scrap["Position"].astype(str).str.strip() == "FRA"])

                self.card_scrap_total.lbl_val.configure(text=f"{total_scrap} Unit")
                self.card_scrap_only.lbl_val.configure(text=f"{scrap_only} Unit")
                self.card_fra_only.lbl_val.configure(text=f"{fra_only} Unit")
            else:
                self.card_scrap_total.lbl_val.configure(text="0 Unit")
                self.card_scrap_only.lbl_val.configure(text="0 Unit")
                self.card_fra_only.lbl_val.configure(text="0 Unit")

    # -------------------------------------------------------------------------
    # MATPLOTLIB CHARTS RENDERING (DASHBOARD & LAPORAN)
    # -------------------------------------------------------------------------
    def render_dashboard_charts(self):
        """Renders Stacked Bar 7-Day Status Trend & Donut Status Distribution on Dashboard."""
        if not hasattr(self, "fig_trend") or not hasattr(self, "fig_donut"):
            return

        # 1. 7-Day Trend Chart
        self.fig_trend.clear()
        ax1 = self.fig_trend.add_subplot(111)
        ax1.set_facecolor("#ffffff")

        trend = get_7day_status_trend(self.df_daily, self.df_fsb)
        x = range(len(trend["labels"]))

        ax1.bar(x, trend["wip"], label="WIP", color="#3b82f6", width=0.55)
        ax1.bar(x, trend["ur"], label="UR", color="#10b981", width=0.55)
        ax1.bar(x, trend["wr"], label="WR", color="#f59e0b", width=0.55)
        ax1.bar(x, trend["hold"], label="Hold", color="#8b5cf6", width=0.55)

        ax1.set_xticks(x)
        ax1.set_xticklabels(trend["labels"], fontsize=9, color="#475569")
        ax1.tick_params(axis="y", labelsize=8, colors="#64748b")
        ax1.grid(axis="y", linestyle="--", alpha=0.3, color="#cbd5e1")
        for spine in ["top", "right", "left"]:
            ax1.spines[spine].set_visible(False)
        ax1.spines["bottom"].set_color("#cbd5e1")

        ax1.legend(loc="upper left", bbox_to_anchor=(1.0, 1.0), fontsize=8, frameon=False)
        self.fig_trend.subplots_adjust(left=0.08, right=0.82, top=0.92, bottom=0.18)
        self.canvas_trend.draw()

        # 2. Donut Chart
        self.fig_donut.clear()
        ax2 = self.fig_donut.add_subplot(111)
        ax2.set_facecolor("#ffffff")

        stats = get_status_distribution(self.df_daily)
        total = stats["total_wip"]

        slices = [stats["total_wip"], max(0, stats["ur"]), max(0, stats["wr"]), max(0, stats["hold"])]
        colors = ["#3b82f6", "#10b981", "#f59e0b", "#8b5cf6"]

        if total > 0:
            wedges, _ = ax2.pie(
                slices,
                colors=colors,
                startangle=90,
                wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2)
            )
            ax2.text(0, 0, f"{total}\nTotal Unit", ha="center", va="center", fontsize=11, fontweight="bold", color="#0f172a")

            legend_labels = [
                f"WIP  {stats['total_wip']} ({100 if total==0 else 50}%)",
                f"UR   {stats['ur']} ({stats['pct_ur']}%)",
                f"WR   {stats['wr']} ({stats['pct_wr']}%)",
                f"Hold {stats['hold']} ({stats['pct_hold']}%)"
            ]
            ax2.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(0.95, 0.5), fontsize=8, frameon=False)
        else:
            ax2.text(0, 0, "Belum Ada\nData Unit", ha="center", va="center", fontsize=10, color="#94a3b8")

        self.fig_donut.subplots_adjust(left=0.05, right=0.75, top=0.92, bottom=0.08)
        self.canvas_donut.draw()

    def render_laporan_charts(self):
        """Renders Weekly & Monthly maintenance completion reports based on selected Year/Week."""
        if not hasattr(self, "fig_weekly") or not hasattr(self, "fig_monthly"):
            return

        try:
            sel_year = int(self.laporan_year_var.get())
        except ValueError:
            sel_year = date.today().year

        try:
            sel_week = int(self.laporan_week_var.get())
        except ValueError:
            sel_week = date.today().isocalendar()[1]

        # 1. Weekly Maintenance Chart
        self.fig_weekly.clear()
        ax_w = self.fig_weekly.add_subplot(111)
        ax_w.set_facecolor("#ffffff")

        w_stats = get_weekly_maintenance_stats(self.df_fsb, self.df_scrap, year=sel_year, week_num=sel_week)
        self.lbl_weekly_title.configure(text=f"Weekly Battery Maintenance ({w_stats['date_range_str']})")

        x = range(7)
        ax_w.bar(x, w_stats["daily_fsb"], label="Serviceable (FSB)", color="#10b981", width=0.5)
        ax_w.bar(x, w_stats["daily_scrap"], bottom=w_stats["daily_fsb"], label="Scrap / FRA", color="#ef4444", width=0.5)

        ax_w.set_xticks(x)
        ax_w.set_xticklabels(HARI_NAMES, fontsize=9, color="#475569")
        ax_w.grid(axis="y", linestyle="--", alpha=0.35, color="#cbd5e1")
        for spine in ["top", "right", "left"]:
            ax_w.spines[spine].set_visible(False)
        ax_w.spines["bottom"].set_color("#cbd5e1")
        ax_w.legend(loc="upper right", fontsize=8, frameon=False)

        # Values above bars
        for i in range(7):
            tot = w_stats["daily_fsb"][i] + w_stats["daily_scrap"][i]
            if tot > 0:
                ax_w.text(i, tot + 0.1, str(tot), ha="center", fontsize=8, fontweight="bold", color="#0f172a")

        self.fig_weekly.subplots_adjust(left=0.08, right=0.95, top=0.90, bottom=0.15)
        self.canvas_weekly.draw()

        # 2. Monthly Maintenance Chart
        self.fig_monthly.clear()
        ax_m = self.fig_monthly.add_subplot(111)
        ax_m.set_facecolor("#ffffff")

        m_stats = get_monthly_maintenance_stats(self.df_fsb, self.df_scrap, year=sel_year)
        self.lbl_monthly_title.configure(text=f"Monthly Battery Maintenance - Tahun {sel_year}")

        x_m = range(12)
        ax_m.bar(x_m, m_stats["fsb_counts"], label="Serviceable (FSB)", color="#2563eb", width=0.55)
        ax_m.bar(x_m, m_stats["scrap_counts"], bottom=m_stats["fsb_counts"], label="Scrap / FRA", color="#f43f5e", width=0.55)

        ax_m.set_xticks(x_m)
        ax_m.set_xticklabels(INDONESIAN_MONTH_SHORT, fontsize=9, color="#475569")
        ax_m.grid(axis="y", linestyle="--", alpha=0.35, color="#cbd5e1")
        for spine in ["top", "right", "left"]:
            ax_m.spines[spine].set_visible(False)
        ax_m.spines["bottom"].set_color("#cbd5e1")
        ax_m.legend(loc="upper right", fontsize=8, frameon=False)

        for i in range(12):
            tot = m_stats["fsb_counts"][i] + m_stats["scrap_counts"][i]
            if tot > 0:
                ax_m.text(i, tot + 0.5, str(tot), ha="center", fontsize=8, fontweight="bold", color="#0f172a")

        self.fig_monthly.subplots_adjust(left=0.08, right=0.95, top=0.90, bottom=0.15)
        self.canvas_monthly.draw()

        # Update Laporan KPI Summary Cards
        self.card_lap_yearly.lbl_val.configure(text=f"{m_stats['total_all']} Unit")
        self.card_lap_weekly.lbl_val.configure(text=f"{w_stats['total_week']} Unit")

        tats = pd.to_numeric(self.df_fsb["TAT"], errors="coerce").dropna()
        avg_tat = round(tats.mean(), 1) if not tats.empty else 0.0
        self.card_lap_tat.lbl_val.configure(text=f"{avg_tat} Hari")

        success_rate = round((m_stats["total_fsb"] / max(1, m_stats["total_all"])) * 100, 1)
        self.card_lap_ratio.lbl_val.configure(text=f"{success_rate}%")

        # Update Customer & Part Number Tables
        self.tree_cust.delete(*self.tree_cust.get_children())
        cust_list = get_customer_breakdown(self.df_fsb, year=sel_year)
        for idx, item in enumerate(cust_list):
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree_cust.insert("", "end", values=[item["customer"], f"{item['count']} Unit", f"{item['avg_tat']} Hari"], tags=(tag,))

        self.tree_pn.delete(*self.tree_pn.get_children())
        pn_list = get_pn_breakdown(self.df_fsb, year=sel_year)
        for idx, item in enumerate(pn_list):
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree_pn.insert("", "end", values=[item["part_number"], f"{item['count']} Unit", f"{item['avg_tat']} Hari", item["top_customer"]], tags=(tag,))

    # -------------------------------------------------------------------------
    # TABLE FILTERING & SORTING ENGINES
    # -------------------------------------------------------------------------
    def filter_daily_table(self):
        """Applies filters to both Dashboard and Daily Battery Treeviews."""
        trees = [self.tree_dash] + ([self.tree_daily] if hasattr(self, "tree_daily") else [])
        if self.df_daily.empty:
            for tree in trees:
                tree.delete(*tree.get_children())
            self.lbl_dash_count.configure(text="Total: 0 Unit")
            return

        df = self.df_daily.copy()

        # Calendar Filter
        if self.active_calendar_date and "PIR" in df.columns:
            df = df[df["PIR"].astype(str).str.strip() == self.active_calendar_date]

        # Search Query
        q = self.daily_search_var.get().strip().lower()
        if q:
            mask = df.astype(str).apply(lambda row: any(q in val.lower() for val in row), axis=1)
            df = df[mask]

        cols = ["WO", "Part Number", "Serial Number", "Customer", "Status", "Progress", "PIR", "TAT", "Position", "Daily", "Estimasi FSB", "Remark"]

        # Populate trees
        for tree in trees:
            tree.delete(*tree.get_children())
            for idx, (_, row) in enumerate(df.iterrows()):
                vals = [str(row.get(c, "")) if pd.notna(row.get(c)) else "" for c in cols]
                pos_tag = str(row.get("Position", "")).strip()
                status_tag = str(row.get("Status", "")).strip()
                zebra_tag = "even" if idx % 2 == 0 else "odd"
                tags = (zebra_tag, pos_tag, status_tag)
                tree.insert("", "end", values=vals, tags=tags)

        self.lbl_dash_count.configure(text=f"Menampilkan: {len(df)} / {len(self.df_daily)} Unit")

    def reset_daily_filters(self):
        self.daily_search_var.set("")
        self.active_calendar_date = None
        if hasattr(self, "btn_calendar"):
            self.btn_calendar.configure(text="📅 Kalender", fg_color="#0284c7")
        self.filter_daily_table()

    def filter_fsb_table(self):
        if self.df_fsb.empty:
            self.tree_fsb.delete(*self.tree_fsb.get_children())
            self.lbl_fsb_count.configure(text="Menampilkan: 0 Unit")
            return

        df = self.df_fsb.copy()
        q = self.fsb_search_var.get().strip().lower()
        if q:
            mask = df.astype(str).apply(lambda row: any(q in val.lower() for val in row), axis=1)
            df = df[mask]

        cols = ["Date FSB", "WO", "Part Number", "Serial Number", "Customer", "Status", "Progress", "TAT", "Position", "Remark"]
        self.tree_fsb.delete(*self.tree_fsb.get_children())
        for idx, (_, row) in enumerate(df.iterrows()):
            vals = [str(row.get(c, "")) if pd.notna(row.get(c)) else "" for c in cols]
            zebra_tag = "even" if idx % 2 == 0 else "odd"
            self.tree_fsb.insert("", "end", values=vals, tags=(zebra_tag, "FSB"))

        self.lbl_fsb_count.configure(text=f"Menampilkan: {len(df)} / {len(self.df_fsb)} Unit")

    def filter_scrap_table(self):
        if self.df_scrap.empty:
            self.tree_scrap.delete(*self.tree_scrap.get_children())
            self.lbl_scrap_count.configure(text="Menampilkan: 0 Unit")
            return

        df = self.df_scrap.copy()
        q = self.scrap_search_var.get().strip().lower()
        if q:
            mask = df.astype(str).apply(lambda row: any(q in val.lower() for val in row), axis=1)
            df = df[mask]

        cols = ["Date SCRAP", "WO", "Part Number", "Serial Number", "Customer", "Status", "TAT", "Position", "Remark"]
        self.tree_scrap.delete(*self.tree_scrap.get_children())
        for idx, (_, row) in enumerate(df.iterrows()):
            vals = [str(row.get(c, "")) if pd.notna(row.get(c)) else "" for c in cols]
            zebra_tag = "even" if idx % 2 == 0 else "odd"
            pos = str(row.get("Position", "")).strip()
            self.tree_scrap.insert("", "end", values=vals, tags=(zebra_tag, pos))

        self.lbl_scrap_count.configure(text=f"Menampilkan: {len(df)} / {len(self.df_scrap)} Unit")

    def sort_treeview(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        try:
            items.sort(key=lambda t: float(t[0]), reverse=reverse)
        except ValueError:
            items.sort(key=lambda t: t[0].lower(), reverse=reverse)

        for index, (_, k) in enumerate(items):
            tree.move(k, "", index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    # -------------------------------------------------------------------------
    # IN-PLACE CELL EDITING WITH DROPDOWNS & ENTRIES
    # -------------------------------------------------------------------------
    def setup_inline_editing(self, tree, tab_type):
        tree.bind("<Double-1>", lambda e: self.on_tree_cell_double_click(e, tree, tab_type))
        tree.bind("<Button-1>", lambda e: self.destroy_active_inline_editor(), add="+")

    def destroy_active_inline_editor(self):
        if self.active_inline_editor:
            try:
                if self.active_inline_editor.winfo_exists():
                    self.active_inline_editor.destroy()
            except Exception:
                pass
            self.active_inline_editor = None

    def on_tree_cell_double_click(self, event, tree, tab_type):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        item_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not item_id or not col_id:
            return

        col_idx = int(col_id.replace("#", "")) - 1
        cols = tree["columns"]
        if col_idx < 0 or col_idx >= len(cols):
            return

        col_name = cols[col_idx]
        self.destroy_active_inline_editor()

        tree.selection_set(item_id)
        bbox = tree.bbox(item_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        values = list(tree.item(item_id, "values"))
        current_val = str(values[col_idx]) if col_idx < len(values) else ""

        # Map dropdown columns
        dropdown_options_map = {
            "Customer": self.dropdown_options.get("Customer", []),
            "Position": self.dropdown_options.get("Position", []),
            "Part Number": self.dropdown_options.get("Part Number", []),
            "Status": self.dropdown_options.get("Status", [])
        }

        if col_name in dropdown_options_map:
            opts = list(dropdown_options_map[col_name])
            cb = ttk.Combobox(tree, values=opts, state="readonly", font=("Segoe UI", 10))
            cb.set(current_val if current_val in opts else (opts[0] if opts else ""))
            cb.place(x=x, y=y, width=w, height=h)
            cb.focus()
            self.active_inline_editor = cb

            def apply_cb(event=None):
                new_v = cb.get()
                self.destroy_active_inline_editor()
                if new_v != current_val:
                    self.save_inline_cell_edit(tree, tab_type, item_id, col_name, new_v)

            cb.bind("<<ComboboxSelected>>", apply_cb)
            cb.bind("<FocusOut>", lambda e: self.destroy_active_inline_editor())
            cb.bind("<Escape>", lambda e: self.destroy_active_inline_editor())
        else:
            ent = ttk.Entry(tree, font=("Segoe UI", 10))
            ent.insert(0, current_val)
            ent.place(x=x, y=y, width=w, height=h)
            ent.focus()
            ent.select_range(0, tk.END)
            self.active_inline_editor = ent

            def apply_ent(event=None):
                new_v = ent.get().strip()
                self.destroy_active_inline_editor()
                if new_v != current_val:
                    self.save_inline_cell_edit(tree, tab_type, item_id, col_name, new_v)

            ent.bind("<Return>", apply_ent)
            ent.bind("<FocusOut>", lambda e: self.destroy_active_inline_editor())
            ent.bind("<Escape>", lambda e: self.destroy_active_inline_editor())

    def save_inline_cell_edit(self, tree, tab_type, item_id, col_name, new_val):
        """Atomically saves in-place edit to active Excel workbook."""
        values = list(tree.item(item_id, "values"))
        cols = list(tree["columns"])
        col_idx = cols.index(col_name)
        old_val = values[col_idx]
        values[col_idx] = new_val
        tree.item(item_id, values=values)

        row_dict = dict(zip(cols, values))

        if tab_type == "daily":
            wo = row_dict.get("WO")
            if col_name == "Status":
                new_status = str(new_val).strip().upper()
                if new_status == "FINAL CHARGE":
                    if messagebox.askyesno(
                        "Pindahkan ke History FSB?",
                        f"Status unit WO '{wo}' diubah menjadi 'FINAL CHARGE' (Serviceable Ready).\n\n"
                        f"Apakah Anda ingin memindahkan unit ini secara otomatis ke worksheet History FSB?",
                        parent=self
                    ):
                        self.prompt_route_unit_to_fsb(row_dict)
                        return
                    else:
                        values[col_idx] = old_val
                        tree.item(item_id, values=values)
                        return
                elif new_status == "SCRAP":
                    if messagebox.askyesno(
                        "Pindahkan ke History SCRAP?",
                        f"Status unit WO '{wo}' diubah menjadi 'SCRAP'.\n\n"
                        f"Apakah Anda ingin memindahkan unit ini secara otomatis ke worksheet History SCRAP?",
                        icon="warning",
                        parent=self
                    ):
                        self.prompt_route_unit_to_scrap(row_dict)
                        return
                    else:
                        values[col_idx] = old_val
                        tree.item(item_id, values=values)
                        return

            ok, msg = self.excel_repo.update_record(SHEET_DAILY, wo, row_dict)
            if ok:
                if not self.df_daily.empty and "WO" in self.df_daily.columns:
                    idx = self.df_daily[self.df_daily["WO"].astype(str).str.strip() == str(wo).strip()].index
                    if len(idx) > 0:
                        self.df_daily.loc[idx[0], col_name] = new_val
                self.last_sync_fingerprint = get_file_fingerprint(self.excel_file_path)
                self.update_kpi_cards()
                self.render_dashboard_charts()
            else:
                messagebox.showerror("Gagal Simpan", f"Gagal memperbarui sel ke Excel:\n{msg}", parent=self)
                self.load_data(silent=True)

        elif tab_type == "fsb":
            ident = {"WO": row_dict.get("WO"), "Serial Number": row_dict.get("Serial Number"), "Date FSB": row_dict.get("Date FSB")}
            ok, msg = self.excel_repo.update_record(SHEET_FSB, ident, row_dict)
            if ok:
                self.last_sync_fingerprint = get_file_fingerprint(self.excel_file_path)
            else:
                messagebox.showerror("Gagal Simpan", f"Gagal memperbarui sel ke History FSB:\n{msg}", parent=self)
                self.load_data(silent=True)

        elif tab_type == "scrap":
            ident = {"WO": row_dict.get("WO"), "Serial Number": row_dict.get("Serial Number"), "Date SCRAP": row_dict.get("Date SCRAP")}
            ok, msg = self.excel_repo.update_record(SHEET_SCRAP, ident, row_dict)
            if ok:
                self.last_sync_fingerprint = get_file_fingerprint(self.excel_file_path)
            else:
                messagebox.showerror("Gagal Simpan", f"Gagal memperbarui sel ke History SCRAP:\n{msg}", parent=self)
                self.load_data(silent=True)

    # -------------------------------------------------------------------------
    # CRUD OPERATIONS & FORM MODALS
    # -------------------------------------------------------------------------
    def open_add_unit_dialog(self):
        next_wo = "512196552"
        if not self.df_daily.empty and "WO" in self.df_daily.columns:
            wos = [str(w).strip() for w in self.df_daily["WO"].dropna() if str(w).strip().isdigit()]
            if wos:
                next_wo = str(int(wos[-1]) + 1)

        BatteryFormDialog(
            self,
            mode="add",
            initial_data={"WO": next_wo},
            dropdown_options=self.dropdown_options,
            on_save_callback=self.handle_save_daily_unit
        )

    def open_selected_daily_form(self):
        tree = self.tree_daily if self.current_view_name == "Daily Battery" else self.tree_dash
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Silakan pilih baris unit yang ingin diedit terlebih dahulu.", parent=self)
            return
        vals = tree.item(sel[0], "values")
        cols = tree["columns"]
        data = dict(zip(cols, vals))
        BatteryFormDialog(
            self,
            mode="edit",
            initial_data=data,
            dropdown_options=self.dropdown_options,
            on_save_callback=self.handle_save_daily_unit,
            on_delete_callback=self.handle_delete_daily_unit
        )

    def handle_save_daily_unit(self, data, mode="add", original_wo=None):
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.\nTutup file tersebut di Excel terlebih dahulu.", parent=self)
            return

        existing_wos = set(self.df_daily["WO"].dropna().astype(str).str.strip().tolist()) if not self.df_daily.empty else set()
        is_valid, err = validate_battery_unit_record(data, existing_wos=existing_wos, is_edit=(mode == "edit"), original_wo=original_wo)
        if not is_valid:
            messagebox.showerror("Validasi Gagal", err, parent=self)
            return

        if mode == "add":
            ok, msg = self.excel_repo.add_record(SHEET_DAILY, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Unit WO {data.get('WO')} berhasil ditambahkan ke worksheet Dashboard!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal menambahkan unit ke Excel:\n{msg}", parent=self)
        else:
            target_wo = str(original_wo or data.get("WO")).strip()
            status = str(data.get("Status", "")).strip().upper()
            if status == "FINAL CHARGE":
                if messagebox.askyesno(
                    "Pindahkan ke History FSB?",
                    f"Status unit WO '{target_wo}' diset ke 'FINAL CHARGE' (Serviceable Ready).\n\n"
                    f"Apakah Anda ingin memindahkan unit ini ke worksheet History FSB?",
                    parent=self
                ):
                    self.prompt_route_unit_to_fsb(data)
                    return
            elif status == "SCRAP":
                if messagebox.askyesno(
                    "Pindahkan ke History SCRAP?",
                    f"Status unit WO '{target_wo}' diset ke 'SCRAP'.\n\n"
                    f"Apakah Anda ingin memindahkan unit ini ke worksheet History SCRAP?",
                    icon="warning",
                    parent=self
                ):
                    self.prompt_route_unit_to_scrap(data)
                    return

            ok, msg = self.excel_repo.update_record(SHEET_DAILY, target_wo, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Unit WO {target_wo} berhasil diperbarui di worksheet Dashboard!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal memperbarui unit di Excel:\n{msg}", parent=self)

    def delete_selected_daily_unit(self):
        tree = self.tree_daily if self.current_view_name == "Daily Battery" else self.tree_dash
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Pilih baris unit yang ingin dihapus terlebih dahulu.", parent=self)
            return
        vals = tree.item(sel[0], "values")
        wo = str(vals[0]).strip()
        if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus unit WO '{wo}' dari worksheet Dashboard Excel?", icon="warning", parent=self):
            self.handle_delete_daily_unit(wo)

    def handle_delete_daily_unit(self, wo):
        clean_wo = str(wo).strip()
        ok, msg = self.excel_repo.delete_record(SHEET_DAILY, clean_wo)
        if ok:
            self.load_data(silent=True)
            messagebox.showinfo("Berhasil Dihapus", f"Unit WO '{clean_wo}' berhasil dihapus dari worksheet Dashboard Excel!", parent=self)
        else:
            messagebox.showerror("Gagal Hapus", f"Gagal menghapus baris dari Excel:\n{msg}", parent=self)

    # -------------------------------------------------------------------------
    # AUTOMATIC STATUS ROUTING (FINAL CHARGE -> FSB, SCRAP -> History SCRAP)
    # -------------------------------------------------------------------------
    def prompt_route_selected_to_fsb(self):
        """Prompt to route selected daily battery unit to History FSB."""
        tree = self.tree_daily if self.current_view_name == "Daily Battery" else self.tree_dash
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Silakan pilih baris unit yang ingin dipindahkan ke History FSB (FINAL CHARGE).", parent=self)
            return
        vals = tree.item(sel[0], "values")
        cols = tree["columns"]
        unit_data = dict(zip(cols, vals))
        self.prompt_route_unit_to_fsb(unit_data)

    def prompt_route_unit_to_fsb(self, unit_data: dict):
        """Opens confirmation dialog to route unit to History FSB."""
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.\nTutup file tersebut terlebih dahulu.", parent=self)
            return
        RouteToFSBDialog(self, unit_data=unit_data, on_confirm_callback=self.execute_route_to_fsb)

    def execute_route_to_fsb(self, wo, completion_date_str, remark_str):
        """Executes atomic routing from Dashboard to History FSB in Excel master workbook."""
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.\nTutup file tersebut terlebih dahulu.", parent=self)
            return

        ok, msg = self.excel_repo.route_unit_to_fsb(wo, completion_date_str=completion_date_str, remark_str=remark_str)
        if ok:
            self.load_data(silent=True)
            messagebox.showinfo(
                "Unit Selesai (Serviceable)",
                f"✅ Unit WO {wo} berhasil dipindahkan ke History FSB!\n\n"
                f"• Tanggal Selesai (Date FSB): {completion_date_str}\n"
                f"• Status: FINAL CHARGE (Serviceable Ready)\n"
                f"• Baris telah dihapus dari tabel Daily Battery aktif.",
                parent=self
            )
        else:
            messagebox.showerror("Gagal Memindahkan Unit", f"Gagal memindahkan unit WO {wo} ke History FSB:\n\n{msg}", parent=self)

    def prompt_route_selected_to_scrap(self):
        """Prompt to route selected daily battery unit to History SCRAP."""
        tree = self.tree_daily if self.current_view_name == "Daily Battery" else self.tree_dash
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Silakan pilih baris unit yang ingin dipindahkan ke History SCRAP.", parent=self)
            return
        vals = tree.item(sel[0], "values")
        cols = tree["columns"]
        unit_data = dict(zip(cols, vals))
        self.prompt_route_unit_to_scrap(unit_data)

    def prompt_route_unit_to_scrap(self, unit_data: dict):
        """Opens confirmation dialog to route unit to History SCRAP."""
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.\nTutup file tersebut terlebih dahulu.", parent=self)
            return
        RouteToSCRAPDialog(self, unit_data=unit_data, on_confirm_callback=self.execute_route_to_scrap)

    def execute_route_to_scrap(self, wo, scrap_date_str, reason_str, pos_str):
        """Executes atomic routing from Dashboard to History SCRAP in Excel master workbook."""
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.\nTutup file tersebut terlebih dahulu.", parent=self)
            return

        ok, msg = self.excel_repo.route_unit_to_scrap(wo, scrap_date_str=scrap_date_str, reason_str=reason_str, position_str=pos_str)
        if ok:
            self.load_data(silent=True)
            messagebox.showinfo(
                "Unit Di-Scrap / FRA",
                f"⚠️ Unit WO {wo} berhasil dipindahkan ke History SCRAP!\n\n"
                f"• Tanggal SCRAP: {scrap_date_str}\n"
                f"• Disposisi: {pos_str}\n"
                f"• Alasan: {reason_str}\n"
                f"• Baris telah dihapus dari tabel Daily Battery aktif.",
                parent=self
            )
        else:
            messagebox.showerror("Gagal Memindahkan Unit", f"Gagal memindahkan unit WO {wo} ke History SCRAP:\n\n{msg}", parent=self)

    # FSB CRUD
    def open_add_fsb_dialog(self):
        FSBFormDialog(self, mode="add", dropdown_options=self.dropdown_options, on_save_callback=self.handle_save_fsb_unit)

    def open_selected_fsb_form(self):
        sel = self.tree_fsb.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Pilih baris riwayat FSB yang ingin diedit.", parent=self)
            return
        vals = self.tree_fsb.item(sel[0], "values")
        cols = self.tree_fsb["columns"]
        data = dict(zip(cols, vals))
        FSBFormDialog(
            self, mode="edit", initial_data=data, dropdown_options=self.dropdown_options,
            on_save_callback=self.handle_save_fsb_unit, on_delete_callback=self.handle_delete_fsb_unit
        )

    def handle_save_fsb_unit(self, data, mode="add", original_ident=None):
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.", parent=self)
            return
        is_valid, err = validate_fsb_record(data)
        if not is_valid:
            messagebox.showerror("Validasi Gagal", err, parent=self)
            return

        if mode == "add":
            ok, msg = self.excel_repo.add_record(SHEET_FSB, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Data FSB WO {data.get('WO')} berhasil dicatat ke History FSB!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal menambah ke History FSB:\n{msg}", parent=self)
        else:
            ok, msg = self.excel_repo.update_record(SHEET_FSB, original_ident, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Data FSB WO {data.get('WO')} berhasil diperbarui di History FSB!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal memperbarui History FSB:\n{msg}", parent=self)

    def delete_selected_fsb_unit(self):
        sel = self.tree_fsb.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Pilih baris riwayat FSB yang ingin dihapus.", parent=self)
            return
        vals = self.tree_fsb.item(sel[0], "values")
        cols = self.tree_fsb["columns"]
        data = dict(zip(cols, vals))
        ident = {"WO": data.get("WO"), "Serial Number": data.get("Serial Number"), "Date FSB": data.get("Date FSB")}
        if messagebox.askyesno("Konfirmasi Hapus", f"Hapus catatan FSB WO '{ident['WO']}' dari History FSB?", icon="warning", parent=self):
            self.handle_delete_fsb_unit(ident)

    def handle_delete_fsb_unit(self, ident):
        ok, msg = self.excel_repo.delete_record(SHEET_FSB, ident)
        if ok:
            self.load_data(silent=True)
            messagebox.showinfo("Berhasil Dihapus", "Catatan berhasil dihapus dari History FSB!", parent=self)
        else:
            messagebox.showerror("Gagal Hapus", f"Gagal menghapus dari History FSB:\n{msg}", parent=self)

    # SCRAP CRUD
    def open_add_scrap_dialog(self):
        SCRAPFormDialog(self, mode="add", dropdown_options=self.dropdown_options, on_save_callback=self.handle_save_scrap_unit)

    def open_selected_scrap_form(self):
        sel = self.tree_scrap.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Pilih baris riwayat SCRAP yang ingin diedit.", parent=self)
            return
        vals = self.tree_scrap.item(sel[0], "values")
        cols = self.tree_scrap["columns"]
        data = dict(zip(cols, vals))
        SCRAPFormDialog(
            self, mode="edit", initial_data=data, dropdown_options=self.dropdown_options,
            on_save_callback=self.handle_save_scrap_unit, on_delete_callback=self.handle_delete_scrap_unit
        )

    def handle_save_scrap_unit(self, data, mode="add", original_ident=None):
        if is_file_locked(self.excel_file_path):
            messagebox.showwarning("File Terkunci", f"File '{self.excel_file_path.name}' sedang dibuka di Microsoft Excel.", parent=self)
            return
        is_valid, err = validate_scrap_record(data)
        if not is_valid:
            messagebox.showerror("Validasi Gagal", err, parent=self)
            return

        if mode == "add":
            ok, msg = self.excel_repo.add_record(SHEET_SCRAP, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Data Scrap WO {data.get('WO')} berhasil dicatat ke History SCRAP!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal menambah ke History SCRAP:\n{msg}", parent=self)
        else:
            ok, msg = self.excel_repo.update_record(SHEET_SCRAP, original_ident, data)
            if ok:
                self.load_data(silent=True)
                messagebox.showinfo("Berhasil", f"Data Scrap WO {data.get('WO')} berhasil diperbarui di History SCRAP!", parent=self)
            else:
                messagebox.showerror("Gagal", f"Gagal memperbarui History SCRAP:\n{msg}", parent=self)

    def delete_selected_scrap_unit(self):
        sel = self.tree_scrap.selection()
        if not sel:
            messagebox.showinfo("Pilih Baris", "Pilih baris riwayat SCRAP yang ingin dihapus.", parent=self)
            return
        vals = self.tree_scrap.item(sel[0], "values")
        cols = self.tree_scrap["columns"]
        data = dict(zip(cols, vals))
        ident = {"WO": data.get("WO"), "Serial Number": data.get("Serial Number"), "Date SCRAP": data.get("Date SCRAP")}
        if messagebox.askyesno("Konfirmasi Hapus", f"Hapus catatan Scrap WO '{ident['WO']}' dari History SCRAP?", icon="warning", parent=self):
            self.handle_delete_scrap_unit(ident)

    def handle_delete_scrap_unit(self, ident):
        ok, msg = self.excel_repo.delete_record(SHEET_SCRAP, ident)
        if ok:
            self.load_data(silent=True)
            messagebox.showinfo("Berhasil Dihapus", "Catatan berhasil dihapus dari History SCRAP!", parent=self)
        else:
            messagebox.showerror("Gagal Hapus", f"Gagal menghapus dari History SCRAP:\n{msg}", parent=self)

    def handle_global_delete_key(self):
        if self.current_view_name == "Dashboard" or self.current_view_name == "Daily Battery":
            self.delete_selected_daily_unit()
        elif self.current_view_name == "History FSB":
            self.delete_selected_fsb_unit()
        elif self.current_view_name == "History SCRAP":
            self.delete_selected_scrap_unit()

    # -------------------------------------------------------------------------
    # WELCOME / WORKBOOK SELECTION MODAL HANDLER
    # -------------------------------------------------------------------------
    def open_welcome_import_dialog(self, force_select=False):
        WelcomeWorkbookDialog(
            self,
            current_file_path=self.excel_file_path,
            on_workbook_selected_callback=self.handle_workbook_switched,
            force_select=force_select
        )

    def handle_workbook_switched(self, new_path):
        """Switches active master workbook after successful validation and reloads all datasets."""
        self.destroy_active_inline_editor()
        self.excel_file_path = Path(new_path)
        set_master_excel_path(self.excel_file_path)
        self.excel_repo = ExcelRepository(self.excel_file_path)
        self.lbl_file_status.configure(text=f"● Live Master: {self.excel_file_path.name}")
        self.load_data(silent=False)
        self.focus_set()

    # -------------------------------------------------------------------------
    # OTHER MODAL HANDLERS
    # -------------------------------------------------------------------------
    def open_dropdown_manager(self):
        DropdownManagerDialog(
            self,
            current_options=self.dropdown_options,
            on_update_callback=self.apply_dropdown_options_update
        )

    def apply_dropdown_options_update(self, new_opts):
        self.dropdown_options = new_opts
        save_dropdown_options(self.dropdown_options)
        self.update_combobox_options()

    def update_combobox_options(self):
        pass

    def open_glossary(self):
        GlossaryDialog(self)

    def open_calendar_dialog(self):
        BatteryCalendarDialog(
            self,
            self.df_daily,
            self.active_calendar_date,
            on_date_selected_callback=self.handle_calendar_date_selected,
            on_clear_callback=self.clear_calendar_filter
        )

    def handle_calendar_date_selected(self, date_str):
        self.active_calendar_date = date_str
        if hasattr(self, "btn_calendar"):
            try:
                parts = date_str.split()
                short_text = f"📅 {parts[0]} {parts[1][:3]}"
            except Exception:
                short_text = f"📅 {date_str[:8]}"
            self.btn_calendar.configure(text=short_text, fg_color="#059669")
        self.filter_daily_table()
        if hasattr(self, "current_view") and self.current_view not in ("Daily Battery", "Dashboard"):
            self.switch_view("Daily Battery")

    def clear_calendar_filter(self):
        self.active_calendar_date = None
        if hasattr(self, "btn_calendar"):
            self.btn_calendar.configure(text="📅 Kalender", fg_color="#0284c7")
        self.filter_daily_table()

    def show_excel_action_menu(self):
        """Displays quick dropdown menu for Excel actions directly from header bar."""
        menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 10), bg="#ffffff", fg="#1e293b", activebackground="#0284c7", activeforeground="#ffffff")
        file_name = self.excel_file_path.name if self.excel_file_path else "Workbook Master"

        menu.add_command(
            label=f"📊  Buka File Asli di Microsoft Excel ({file_name})",
            command=self.open_master_in_excel
        )
        menu.add_command(
            label="📂  Buka Lokasi Folder di File Explorer",
            command=self.open_excel_in_explorer
        )
        menu.add_separator()
        menu.add_command(
            label="📁  Pilih / Import Master Workbook Lain...",
            command=self.open_welcome_import_dialog
        )
        menu.add_command(
            label="🔄  Refresh & Sinkronisasi Ulang Data",
            command=lambda: self.load_data(silent=False)
        )

        try:
            bx = self.btn_excel.winfo_rootx()
            by = self.btn_excel.winfo_rooty() + self.btn_excel.winfo_height() + 2
            menu.post(bx, by)
            menu.focus_set()
            menu.bind("<FocusOut>", lambda e: menu.unpost())
        except Exception:
            self.open_welcome_import_dialog()

    def open_master_in_excel(self):
        """Opens active master Excel file directly in default system spreadsheet application."""
        if not self.excel_file_path or not self.excel_file_path.exists():
            messagebox.showwarning(
                "File Tidak Ditemukan",
                f"File Excel master tidak ditemukan di lokasi:\n{self.excel_file_path}",
                parent=self
            )
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(self.excel_file_path.resolve()))
            else:
                subprocess.Popen(["xdg-open", str(self.excel_file_path.resolve())])
        except Exception as e:
            messagebox.showerror(
                "Gagal Membuka File",
                f"Terjadi kesalahan saat membuka file Excel:\n{e}",
                parent=self
            )

    def open_excel_in_explorer(self):
        if not self.excel_file_path:
            return
        if sys.platform == "win32":
            if self.excel_file_path.exists():
                subprocess.run(["explorer", f"/select,{str(self.excel_file_path.resolve())}"], check=False)
            else:
                subprocess.run(["explorer", str(self.excel_file_path.parent.resolve())], check=False)
        else:
            subprocess.Popen(["xdg-open", str(self.excel_file_path.parent.resolve())])

    def open_backups_folder(self):
        bdir = get_backup_directory()
        if sys.platform == "win32":
            subprocess.run(["explorer", str(bdir)])

    def open_restore_backup_modal(self):
        ExcelAccessDialog(
            self,
            self.excel_file_path,
            on_file_changed_callback=self.handle_workbook_switched,
            on_reload_callback=lambda: self.load_data(silent=False)
        )

    def export_laporan_to_excel(self):
        """Exports analytical report to a standalone Excel file without altering the master."""
        try:
            sel_year = int(self.laporan_year_var.get())
        except ValueError:
            sel_year = date.today().year

        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Simpan Laporan Maintenance ke Excel",
            initialfile=f"Laporan_Battery_Shop_TCC2_{sel_year}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                # Monthly stats
                m_stats = get_monthly_maintenance_stats(self.df_fsb, self.df_scrap, year=sel_year)
                df_monthly = pd.DataFrame({
                    "Bulan": INDONESIAN_MONTH_NAMES,
                    "Serviceable (FSB)": m_stats["fsb_counts"],
                    "Scrap / FRA": m_stats["scrap_counts"],
                    "Total Selesai": [m_stats["fsb_counts"][i] + m_stats["scrap_counts"][i] for i in range(12)]
                })
                df_monthly.to_excel(writer, sheet_name="Monthly Maintenance", index=False)

                # Customer stats
                cust_list = get_customer_breakdown(self.df_fsb, year=sel_year)
                if cust_list:
                    pd.DataFrame(cust_list).to_excel(writer, sheet_name="Per Customer", index=False)

                # PN stats
                pn_list = get_pn_breakdown(self.df_fsb, year=sel_year)
                if pn_list:
                    pd.DataFrame(pn_list).to_excel(writer, sheet_name="Per Part Number", index=False)

            messagebox.showinfo("Export Berhasil", f"Laporan tahun {sel_year} berhasil diekspor ke:\n{save_path}", parent=self)
        except Exception as e:
            messagebox.showerror("Gagal Export", f"Terjadi kesalahan saat mengekspor laporan:\n{e}", parent=self)

    # -------------------------------------------------------------------------
    # DIGITAL CLOCK TICKER
    # -------------------------------------------------------------------------
    def update_live_clock(self):
        now = datetime.now()
        day_names_map = {0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis", 4: "Jumat", 5: "Sabtu", 6: "Minggu"}
        day_str = day_names_map.get(now.weekday(), "")
        month_str = MONTHS_FULL_ID[now.month - 1]

        self.lbl_day_date.configure(text=f"📅 {day_str}, {now.day} {month_str} {now.year}")
        self.lbl_clock_time.configure(text=f"🕒 {now.strftime('%H:%M:%S')} WIB")
        self.after(1000, self.update_live_clock)


# -----------------------------------------------------------------------------
# APPLICATION ENTRY POINT
# -----------------------------------------------------------------------------
def main():
    app = BatteryShopApp()
    app.mainloop()

if __name__ == "__main__":
    main()
