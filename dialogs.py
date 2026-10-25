"""
Dialogs and Modals for Aircraft Battery Maintenance Monitoring System (TCC-2 | GMF AeroAsia)
Contains:
1. WelcomeWorkbookDialog - Professional startup & import file selection modal with schema validation
2. BatteryFormDialog - Full Add/Edit modal for Daily Battery (WIP)
3. FSBFormDialog - Full Add/Edit modal for History FSB
4. SCRAPFormDialog - Full Add/Edit modal for History SCRAP
5. DropdownManagerDialog - Dropdown options manager
6. GlossaryDialog - Technical terminology & abbreviations reference
7. ExcelAccessDialog - Excel master manager & backup restoration
8. BatteryCalendarDialog - Calendar PIR date filter modal
"""

import os
import sys
import re
import json
import shutil
import subprocess
from datetime import datetime, date, timedelta
from collections import Counter
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image

import customtkinter as ctk
import pandas as pd
import openpyxl

from config import BASE_DIR, get_master_excel_path, set_master_excel_path, get_resource_path
from backup_manager import create_backup, list_backups, restore_backup, get_backup_directory
from sync_service import is_file_locked, check_file_locked, get_file_fingerprint
from validation import (
    validate_battery_unit_record,
    validate_fsb_record,
    validate_scrap_record,
    validate_workbook_schema,
    inspect_workbook_summary,
    MONTHS_ID,
    MONTHS_FULL_ID,
    DAYS_FULL_ID,
    parse_date_flexible,
    format_date_with_month_name
)

Calendar = None
HAS_TKCALENDAR = None

def _ensure_tkcalendar():
    global Calendar, HAS_TKCALENDAR
    if HAS_TKCALENDAR is None:
        try:
            from tkcalendar import Calendar as _Cal
            Calendar = _Cal
            HAS_TKCALENDAR = True
        except ImportError:
            Calendar = None
            HAS_TKCALENDAR = False
    return HAS_TKCALENDAR


def bring_window_to_front(win):
    """Forces window to the foreground on Windows on launch without preventing minimize."""
    try:
        win.lift()
        win.attributes("-topmost", True)
        win.after(100, lambda: win.attributes("-topmost", False))
        win.focus_force()
    except Exception:
        pass


# -----------------------------------------------------------------------------
# 1. WELCOME & WORKBOOK SELECTION MODAL
# -----------------------------------------------------------------------------
class WelcomeWorkbookDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_file_path=None, on_workbook_selected_callback=None, force_select=False):
        super().__init__(parent)
        self.on_workbook_selected_callback = on_workbook_selected_callback
        self.current_file_path = Path(current_file_path) if current_file_path else None
        self.selected_file_path = self.current_file_path
        self.force_select = force_select

        self.title("Pilih / Import Workbook Excel Master - Battery Shop TCC-2")
        self.geometry("740x520")
        self.minsize(660, 440)
        self.after(10, lambda: bring_window_to_front(self))

        # Main scroll container
        self.container = ctk.CTkScrollableFrame(self, fg_color=("white", "#0f172a"))
        self.container.pack(fill="both", expand=True, padx=16, pady=16)
        container = self.container

        # Header with Logo & Brand
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 14))

        logo_path = get_resource_path("assets/gmf_logo.png")
        if logo_path.exists():
            try:
                pil_img = Image.open(logo_path)
                self.img_logo = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(75, 42))
                ctk.CTkLabel(header, image=self.img_logo, text="").pack(side="left", padx=(0, 14))
            except Exception:
                pass

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")

        ctk.CTkLabel(
            title_box,
            text="BATTERY SHOP TCC-2 | PT GMF AEROASIA Tbk",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_box,
            text="Aircraft Battery Maintenance Dashboard & Operational Monitoring System",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#64748b", "#94a3b8")
        ).pack(anchor="w")

        # Operational Concept Info Box
        concept_card = ctk.CTkFrame(
            container,
            fg_color=("#f1f5f9", "#1e293b"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            corner_radius=10
        )
        concept_card.pack(fill="x", pady=(0, 16), padx=2)

        ctk.CTkLabel(
            concept_card,
            text="📌 Konsep Operasional Master Excel & Offline Dashboard",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#0284c7", "#38bdf8")
        ).pack(anchor="w", padx=14, pady=(10, 4))

        concept_text = (
            "• Hanya ada SATU workbook master Excel sebagai sumber data utama:\n"
            "   1. 'Dashboard' untuk Daily Battery (WIP)\n"
            "   2. 'History FSB' untuk Riwayat Baterai Serviceable Lolos Uji\n"
            "   3. 'History SCRAP' untuk Riwayat Baterai Scrap & Finish Return As Is (FRA)\n"
            "• File Excel dapat diedit secara manual di HP ketika di luar kantor. Saat tiba di kantor,\n"
            "   pindahkan file terbaru ke laptop dan lakukan Import / Refresh Data di dashboard ini.\n"
            "• Seluruh KPI, tabel, dan grafik dihitung 100% offline dari workbook aktif."
        )
        ctk.CTkLabel(
            concept_card,
            text=concept_text,
            font=ctk.CTkFont(size=11),
            justify="left",
            text_color=("#334155", "#cbd5e1")
        ).pack(anchor="w", padx=14, pady=(0, 12))

        # Last Used File Section
        self.last_card = ctk.CTkFrame(
            container,
            fg_color=("#ffffff", "#1e293b"),
            border_width=1,
            border_color=("#e2e8f0", "#334155"),
            corner_radius=10
        )
        self.last_card.pack(fill="x", pady=(0, 16), padx=2)

        ctk.CTkLabel(
            self.last_card,
            text="📄 Master Workbook Aktif Saat Ini",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=14, pady=(12, 4))

        self.last_summary = None
        if self.current_file_path and self.current_file_path.exists():
            self.last_summary = inspect_workbook_summary(self.current_file_path)
            stat = self.current_file_path.stat()
            dt_str = datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y, %H:%M:%S")

            status_color = "#16a34a" if self.last_summary["is_valid"] else "#dc2626"
            status_text = "✔ Siap & Terverifikasi" if self.last_summary["is_valid"] else "❌ Struktur Tidak Sesuai"

            try:
                rel_path = self.current_file_path.resolve().relative_to(BASE_DIR.resolve())
                loc_text = f"[Folder Aplikasi] \\ {rel_path}  (Portable / Otomatis Menyesuaikan Laptop)"
            except Exception:
                loc_text = f"{self.current_file_path.name}  (Di luar folder aplikasi)"

            file_info = (
                f"• Nama File : {self.current_file_path.name} ({self.last_summary['file_size_kb']} KB)\n"
                f"• Lokasi    : {loc_text}\n"
                f"  Path Lengkap: {self.current_file_path}\n"
                f"• Terakhir  : {dt_str}\n"
                f"• Status    : {status_text}\n"
                f"• Isi Data  : {self.last_summary['daily_count']} Daily WIP | {self.last_summary['fsb_count']} FSB | {self.last_summary['scrap_count']} SCRAP"
            )
            ctk.CTkLabel(
                self.last_card,
                text=file_info,
                font=ctk.CTkFont(family="Consolas", size=11),
                justify="left",
                text_color=("#1e293b", "#f8fafc")
            ).pack(anchor="w", padx=14, pady=(0, 12))

            if self.last_summary["is_valid"]:
                btn_actions_frame = ctk.CTkFrame(self.last_card, fg_color="transparent")
                btn_actions_frame.pack(fill="x", padx=14, pady=(0, 14))

                # Direct action to open file in Microsoft Excel
                btn_open_excel = ctk.CTkButton(
                    btn_actions_frame,
                    text=f"📊 Buka File Asli di Microsoft Excel ({self.current_file_path.name})",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    fg_color="#107c41",  # Excel green
                    hover_color="#0b5a2f",
                    height=38,
                    command=self.open_in_native_excel
                )
                btn_open_excel.pack(fill="x", pady=(0, 8))

                row_sub = ctk.CTkFrame(btn_actions_frame, fg_color="transparent")
                row_sub.pack(fill="x")

                btn_use_last = ctk.CTkButton(
                    row_sub,
                    text="⚡ Gunakan Sebagai Master Aktif",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    fg_color="#0284c7",
                    hover_color="#0369a1",
                    height=36,
                    command=self.confirm_use_last
                )
                btn_use_last.pack(side="left", fill="x", expand=True, padx=(0, 6))

                btn_reveal = ctk.CTkButton(
                    row_sub,
                    text="📂 Lokasi Folder",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    fg_color="#475569",
                    hover_color="#334155",
                    width=130,
                    height=36,
                    command=self.open_file_folder
                )
                btn_reveal.pack(side="right")
        else:
            ctk.CTkLabel(
                self.last_card,
                text="Belum ada file master yang tersimpan atau file sebelumnya tidak ditemukan di disk.",
                font=ctk.CTkFont(size=11),
                text_color="#ef4444"
            ).pack(anchor="w", padx=14, pady=(0, 12))

        # Select New File Section
        new_card = ctk.CTkFrame(
            container,
            fg_color=("#ffffff", "#1e293b"),
            border_width=1,
            border_color=("#e2e8f0", "#334155"),
            corner_radius=10
        )
        new_card.pack(fill="x", pady=(0, 16), padx=2)

        ctk.CTkLabel(
            new_card,
            text="📂 Pilih / Import File Workbook Lain (Versi Terbaru dari HP)",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=14, pady=(12, 4))

        ctk.CTkLabel(
            new_card,
            text="Pilih file Excel yang telah diperbarui atau dipindahkan dari HP:",
            font=ctk.CTkFont(size=11),
            text_color=("#64748b", "#94a3b8")
        ).pack(anchor="w", padx=14, pady=(0, 10))

        btn_browse = ctk.CTkButton(
            new_card,
            text="📁 Telusuri & Pilih File Excel (*.xlsx)...",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            height=38,
            command=self.browse_and_validate_file
        )
        btn_browse.pack(fill="x", padx=14, pady=(0, 14))

        # Validation Preview Frame
        self.preview_frame = ctk.CTkFrame(
            container,
            fg_color=("#f8fafc", "#0f172a"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            corner_radius=10
        )
        self.lbl_preview_title = ctk.CTkLabel(
            self.preview_frame,
            text="🔍 Hasil Validasi Struktur Workbook",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_preview_text = ctk.CTkLabel(
            self.preview_frame,
            text="",
            font=ctk.CTkFont(family="Consolas", size=11),
            justify="left"
        )
        self.btn_apply_new = ctk.CTkButton(
            self.preview_frame,
            text="✅ Terapkan File Ini & Muat Dashboard",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=38,
            command=self.confirm_use_selected
        )

    def confirm_use_last(self):
        if self.on_workbook_selected_callback and self.current_file_path:
            chosen = Path(self.current_file_path).resolve()
            set_master_excel_path(chosen)
            self.on_workbook_selected_callback(chosen)
        self.destroy()

    def open_in_native_excel(self):
        """Opens active master file directly in Microsoft Excel or system default spreadsheet editor."""
        if self.current_file_path and self.current_file_path.exists():
            try:
                if sys.platform == "win32":
                    os.startfile(str(self.current_file_path.resolve()))
                else:
                    subprocess.Popen(["xdg-open", str(self.current_file_path.resolve())])
            except Exception as e:
                messagebox.showerror("Gagal Membuka Excel", f"Tidak dapat membuka file Excel:\n{e}", parent=self)
        else:
            messagebox.showwarning("File Tidak Ditemukan", "File Excel master tidak ditemukan di disk.", parent=self)

    def open_file_folder(self):
        """Reveals file location in Windows File Explorer."""
        if self.current_file_path and self.current_file_path.exists():
            try:
                if sys.platform == "win32":
                    subprocess.run(["explorer", f"/select,{str(self.current_file_path.resolve())}"], check=False)
                else:
                    subprocess.Popen(["xdg-open", str(self.current_file_path.parent.resolve())])
            except Exception as e:
                messagebox.showerror("Gagal Membuka Folder", f"Tidak dapat membuka lokasi folder:\n{e}", parent=self)

    def browse_and_validate_file(self):
        init_dir = self.current_file_path.parent if self.current_file_path and self.current_file_path.exists() else BASE_DIR
        chosen = filedialog.askopenfilename(
            parent=self,
            title="Pilih File Master Excel Battery Shop TCC-2",
            initialdir=str(init_dir),
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not chosen:
            return

        chosen_path = Path(chosen).resolve()
        summary = inspect_workbook_summary(chosen_path)

        self.preview_frame.pack(fill="x", pady=(0, 16), padx=2)
        self.lbl_preview_title.pack(anchor="w", padx=14, pady=(10, 4))
        self.lbl_preview_text.pack(anchor="w", padx=14, pady=(0, 10))

        if summary["is_valid"]:
            try:
                rel = chosen_path.relative_to(BASE_DIR.resolve())
                loc_txt = f"[Folder Aplikasi] \\ {rel} (Portable)"
            except Exception:
                loc_txt = str(chosen_path)

            txt = (
                f"• Nama File   : {chosen_path.name} ({summary['file_size_kb']} KB)\n"
                f"• Lokasi      : {loc_txt}\n"
                f"• Sheet Daily : {summary['daily_count']} Unit Daily WIP (Akan Dimuat ke Dashboard)\n"
                f"• History FSB : {summary['fsb_count']} Unit Serviceable Ready\n"
                f"• Hist. SCRAP : {summary['scrap_count']} Unit Scrap & FRA\n"
                f"• Struktur    : ✔ Lengkap & Terverifikasi Sah"
            )
            self.lbl_preview_title.configure(text="🔍 Hasil Validasi: Workbook Sah & Siap Digunakan", text_color="#15803d")
            self.lbl_preview_text.configure(text=txt, text_color="#15803d")
            self.selected_file_path = chosen_path
            self.btn_apply_new.configure(text=f"✅ Terapkan '{chosen_path.name}' & Muat Dashboard")
            self.btn_apply_new.pack(fill="x", padx=14, pady=(0, 14))

            # Auto-scroll so preview and Apply button are immediately visible
            self.update_idletasks()
            try:
                self.container._parent_canvas.yview_moveto(1.0)
            except Exception:
                pass

            # Direct 1-click confirmation prompt so user can immediately apply without scrolling
            apply_now = messagebox.askyesno(
                "Terapkan Workbook Baru",
                f"File Excel '{chosen_path.name}' berhasil divalidasi!\n\n"
                f"• Daily WIP    : {summary['daily_count']} Unit\n"
                f"• History FSB  : {summary['fsb_count']} Unit\n"
                f"• History SCRAP: {summary['scrap_count']} Unit\n\n"
                f"Terapkan workbook ini sebagai sumber data aktif dashboard?",
                parent=self
            )
            if apply_now:
                self.confirm_use_selected()
        else:
            err_str = "\n".join([f"• {e}" for e in summary["errors"]])
            txt = f"❌ VALIDASI GAGAL:\n{err_str}\n\nHarap pilih workbook yang memiliki sheet Dashboard, History FSB, dan History SCRAP."
            self.lbl_preview_title.configure(text="❌ Hasil Validasi: Struktur Tidak Sesuai", text_color="#dc2626")
            self.lbl_preview_text.configure(text=txt, text_color="#dc2626")
            if self.btn_apply_new.winfo_ismapped():
                self.btn_apply_new.pack_forget()

            self.update_idletasks()
            try:
                self.container._parent_canvas.yview_moveto(1.0)
            except Exception:
                pass
            messagebox.showerror(
                "Validasi Gagal",
                f"File '{chosen_path.name}' tidak dapat digunakan sebagai master:\n\n{err_str}",
                parent=self
            )

    def confirm_use_selected(self):
        if self.selected_file_path and self.on_workbook_selected_callback:
            chosen = Path(self.selected_file_path).resolve()
            set_master_excel_path(chosen)
            self.on_workbook_selected_callback(chosen)
        self.destroy()


# -----------------------------------------------------------------------------
# 2. DROPDOWN MANAGER DIALOG
# -----------------------------------------------------------------------------
class DropdownManagerDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_options, on_update_callback):
        super().__init__(parent)
        self.on_update_callback = on_update_callback
        self.options = {k: list(v) for k, v in current_options.items()}
        self.categories = list(self.options.keys())
        self.current_category = self.categories[0] if self.categories else ""

        self.title("Pengaturan Opsi Dropdown Input Baterai")
        self.geometry("640x480")
        self.minsize(560, 400)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(16, 10))

        ctk.CTkLabel(
            header,
            text="⚙️ Kelola Opsi Dropdown Standar Bengkel",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        ).pack(anchor="w", padx=14, pady=(10, 2))

        ctk.CTkLabel(
            header,
            text="Tambahkan atau hapus pilihan Part Number, Customer, Position, dan Status.",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray70")
        ).pack(anchor="w", padx=14, pady=(0, 10))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=4)

        cat_frame = ctk.CTkFrame(body, width=170, corner_radius=10)
        cat_frame.pack(side="left", fill="y", padx=(0, 10))
        cat_frame.pack_propagate(False)

        ctk.CTkLabel(cat_frame, text="Kategori", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=(10, 6))

        self.cat_buttons = {}
        for cat in self.categories:
            btn = ctk.CTkButton(
                cat_frame,
                text=cat,
                command=lambda c=cat: self.select_category(c),
                fg_color="#2563eb" if cat == self.current_category else "transparent",
                text_color="white" if cat == self.current_category else ("gray20", "gray80"),
                anchor="w",
                height=32
            )
            btn.pack(fill="x", padx=8, pady=3)
            self.cat_buttons[cat] = btn

        edit_frame = ctk.CTkFrame(body, corner_radius=10)
        edit_frame.pack(side="right", fill="both", expand=True)

        input_box = ctk.CTkFrame(edit_frame, fg_color="transparent")
        input_box.pack(fill="x", padx=12, pady=(12, 6))

        self.entry_new_item = ctk.CTkEntry(input_box, placeholder_text="Ketik opsi baru...", height=34)
        self.entry_new_item.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_new_item.bind("<Return>", lambda e: self.add_item())

        btn_add = ctk.CTkButton(input_box, text="➕ Tambah", width=80, height=34, fg_color="#059669", hover_color="#047857", command=self.add_item)
        btn_add.pack(side="right")

        list_frame = ctk.CTkFrame(edit_frame, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=12, pady=6)

        self.listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 10),
            bg="#ffffff" if ctk.get_appearance_mode() == "Light" else "#1e293b",
            fg="#0f172a" if ctk.get_appearance_mode() == "Light" else "#f8fafc",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            relief="flat",
            borderwidth=1,
            highlightthickness=0
        )
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        btn_box = ctk.CTkFrame(edit_frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=12, pady=(6, 12))

        btn_del = ctk.CTkButton(btn_box, text="🗑️ Hapus yang Dipilih", width=140, height=30, fg_color="#ef4444", hover_color="#dc2626", command=self.delete_selected_item)
        btn_del.pack(side="left")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(6, 16))

        btn_save = ctk.CTkButton(footer, text="💾 Simpan & Terapkan", font=ctk.CTkFont(size=12, weight="bold"), fg_color="#2563eb", hover_color="#1d4ed8", height=36, command=self.save_and_apply)
        btn_save.pack(side="right", padx=(8, 0))

        btn_cancel = ctk.CTkButton(footer, text="Batal", width=90, height=36, fg_color="transparent", border_width=1, text_color=("gray20", "gray80"), command=self.destroy)
        btn_cancel.pack(side="right")

        self.refresh_listbox()

    def select_category(self, cat):
        self.current_category = cat
        for c, btn in self.cat_buttons.items():
            if c == cat:
                btn.configure(fg_color="#2563eb", text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color=("gray20", "gray80"))
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        items = self.options.get(self.current_category, [])
        for item in sorted(items):
            self.listbox.insert(tk.END, item)

    def add_item(self):
        val = self.entry_new_item.get().strip()
        if not val:
            return
        curr_list = self.options.setdefault(self.current_category, [])
        if val in curr_list:
            messagebox.showwarning("Duplikat", f"'{val}' sudah ada dalam kategori {self.current_category}!", parent=self)
            return
        curr_list.append(val)
        self.entry_new_item.delete(0, tk.END)
        self.refresh_listbox()

    def delete_selected_item(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        val = self.listbox.get(sel[0])
        curr_list = self.options.get(self.current_category, [])
        if val in curr_list:
            curr_list.remove(val)
        self.refresh_listbox()

    def save_and_apply(self):
        if self.on_update_callback:
            self.on_update_callback(self.options)
        self.destroy()


# -----------------------------------------------------------------------------
# 3. BATTERY FORM DIALOG (ADD / EDIT DAILY BATTERY)
# -----------------------------------------------------------------------------
class BatteryFormDialog(ctk.CTkToplevel):
    def __init__(self, parent, mode="add", initial_data=None, dropdown_options=None, on_save_callback=None, on_delete_callback=None):
        super().__init__(parent)
        self.mode = mode
        self.initial_data = initial_data or {}
        self.dropdown_options = dropdown_options or {}
        self.on_save_callback = on_save_callback
        self.on_delete_callback = on_delete_callback

        is_edit = (mode == "edit")
        title_text = "✏️ Form Lengkap: Edit Unit Baterai (WIP)" if is_edit else "➕ Registrasi Masuk Unit Baterai (WIP)"
        self.title(title_text)
        self.geometry("660x520")
        self.minsize(580, 420)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(header, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        sub_text = "Perbarui detail unit di worksheet Dashboard" if is_edit else "Tambah baterai baru yang masuk ke bengkel ke worksheet Dashboard"
        ctk.CTkLabel(header, text=sub_text, font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 10))

        form_frame = ctk.CTkScrollableFrame(self, corner_radius=10)
        form_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        self.entries = {}

        def create_row(label_text, var_name, widget_type="entry", options=None, default_val=""):
            box = ctk.CTkFrame(form_frame, fg_color="transparent")
            box.pack(fill="x", pady=4)

            ctk.CTkLabel(box, text=label_text, width=170, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(4, 8))

            if widget_type == "combobox":
                opts = list(options or [])
                val = str(default_val) if default_val in opts else (opts[0] if opts else "")
                w = ctk.CTkComboBox(box, values=opts, width=280, height=32)
                if val:
                    w.set(val)
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w
            else:
                w = ctk.CTkEntry(box, width=280, height=32)
                if default_val:
                    w.insert(0, str(default_val))
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w

        # Form fields
        create_row("Nomor Work Order (WO)*", "WO", "entry", default_val=self.initial_data.get("WO", ""))
        create_row("Part Number (P/N)*", "Part Number", "combobox", options=self.dropdown_options.get("Part Number", []), default_val=self.initial_data.get("Part Number", ""))
        create_row("Serial Number (S/N)*", "Serial Number", "entry", default_val=self.initial_data.get("Serial Number", ""))
        create_row("Customer / Maskapai*", "Customer", "combobox", options=self.dropdown_options.get("Customer", []), default_val=self.initial_data.get("Customer", ""))
        create_row("Status Perawatan*", "Status", "combobox", options=self.dropdown_options.get("Status", []), default_val=self.initial_data.get("Status", "Reguler"))
        create_row("Progress Pengerjaan", "Progress", "entry", default_val=self.initial_data.get("Progress", "-"))
        create_row("Tanggal PIR (Terima)*", "PIR", "entry", default_val=self.initial_data.get("PIR", datetime.now().strftime("%d %B %Y")))
        create_row("TAT (Turn Around Time)", "TAT", "entry", default_val=str(self.initial_data.get("TAT", "")))
        create_row("Position (Lokasi Bengkel)*", "Position", "combobox", options=self.dropdown_options.get("Position", []), default_val=self.initial_data.get("Position", "WR"))
        create_row("Daily Target", "Daily", "entry", default_val=str(self.initial_data.get("Daily", "")))
        create_row("Estimasi Selesai (FSB)", "Estimasi FSB", "entry", default_val=str(self.initial_data.get("Estimasi FSB", "")))
        create_row("Catatan Tambahan (Remark)", "Remark", "entry", default_val=self.initial_data.get("Remark", ""))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(6, 12))

        if is_edit and self.on_delete_callback:
            btn_del = ctk.CTkButton(footer, text="🗑️ Hapus Unit", fg_color="#ef4444", hover_color="#dc2626", width=110, height=36, command=self.confirm_delete)
            btn_del.pack(side="left")

        btn_save = ctk.CTkButton(footer, text="💾 Simpan ke Excel", font=ctk.CTkFont(size=12, weight="bold"), fg_color="#059669", hover_color="#047857", height=36, command=self.save_unit)
        btn_save.pack(side="right", padx=(8, 0))

        btn_cancel = ctk.CTkButton(footer, text="Batal", width=90, height=36, fg_color="transparent", border_width=1, text_color=("gray20", "gray80"), command=self.destroy)
        btn_cancel.pack(side="right")

    def save_unit(self):
        data = {k: w.get().strip() for k, w in self.entries.items()}
        if self.on_save_callback:
            orig_wo = self.initial_data.get("WO") if self.mode == "edit" else None
            self.on_save_callback(data, mode=self.mode, original_wo=orig_wo)
        self.destroy()

    def confirm_delete(self):
        wo = self.initial_data.get("WO")
        if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus unit WO '{wo}' dari worksheet Dashboard Excel?", icon="warning", parent=self):
            if self.on_delete_callback:
                self.on_delete_callback(wo)
            self.destroy()


# -----------------------------------------------------------------------------
# 4. FSB FORM DIALOG (HISTORY FSB ADD/EDIT)
# -----------------------------------------------------------------------------
class FSBFormDialog(ctk.CTkToplevel):
    def __init__(self, parent, mode="add", initial_data=None, dropdown_options=None, on_save_callback=None, on_delete_callback=None):
        super().__init__(parent)
        self.mode = mode
        self.initial_data = initial_data or {}
        self.dropdown_options = dropdown_options or {}
        self.on_save_callback = on_save_callback
        self.on_delete_callback = on_delete_callback

        is_edit = (mode == "edit")
        title_text = "✏️ Edit Riwayat Baterai Serviceable (FSB)" if is_edit else "➕ Catat Baterai Serviceable Selesai (FSB)"
        self.title(title_text)
        self.geometry("660x520")
        self.minsize(580, 420)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(header, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        sub_text = "Data disimpan ke worksheet 'History FSB' pada workbook aktif"
        ctk.CTkLabel(header, text=sub_text, font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 10))

        form_frame = ctk.CTkScrollableFrame(self, corner_radius=10)
        form_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        self.entries = {}

        def create_row(label_text, var_name, widget_type="entry", options=None, default_val=""):
            box = ctk.CTkFrame(form_frame, fg_color="transparent")
            box.pack(fill="x", pady=4)
            ctk.CTkLabel(box, text=label_text, width=170, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(4, 8))
            if widget_type == "combobox":
                opts = list(options or [])
                val = str(default_val) if default_val in opts else (opts[0] if opts else "")
                w = ctk.CTkComboBox(box, values=opts, width=280, height=32)
                if val:
                    w.set(val)
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w
            else:
                w = ctk.CTkEntry(box, width=280, height=32)
                if default_val:
                    w.insert(0, str(default_val))
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w

        create_row("Tanggal Selesai (Date FSB)*", "Date FSB", "entry", default_val=self.initial_data.get("Date FSB", datetime.now().strftime("%d %B %Y")))
        create_row("Nomor Work Order (WO)*", "WO", "entry", default_val=self.initial_data.get("WO", ""))
        create_row("Part Number (P/N)*", "Part Number", "combobox", options=self.dropdown_options.get("Part Number", []), default_val=self.initial_data.get("Part Number", ""))
        create_row("Serial Number (S/N)*", "Serial Number", "entry", default_val=self.initial_data.get("Serial Number", ""))
        create_row("Customer / Maskapai*", "Customer", "combobox", options=self.dropdown_options.get("Customer", []), default_val=self.initial_data.get("Customer", ""))
        create_row("Status Perawatan*", "Status", "combobox", options=self.dropdown_options.get("Status", []), default_val=self.initial_data.get("Status", "Reguler"))
        create_row("Progress Akhir", "Progress", "entry", default_val=self.initial_data.get("Progress", "Serviceable Ready"))
        create_row("TAT (Hari Kerja)", "TAT", "entry", default_val=str(self.initial_data.get("TAT", "6")))
        create_row("Position", "Position", "entry", default_val=self.initial_data.get("Position", "FSB"))
        create_row("Catatan / Form 1 Remark", "Remark", "entry", default_val=self.initial_data.get("Remark", ""))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(6, 16))

        if is_edit and self.on_delete_callback:
            btn_del = ctk.CTkButton(footer, text="🗑️ Hapus Catatan", fg_color="#ef4444", hover_color="#dc2626", width=120, height=36, command=self.confirm_delete)
            btn_del.pack(side="left")

        btn_save = ctk.CTkButton(footer, text="💾 Simpan ke History FSB", font=ctk.CTkFont(size=12, weight="bold"), fg_color="#059669", hover_color="#047857", height=36, command=self.save_fsb)
        btn_save.pack(side="right", padx=(8, 0))

        btn_cancel = ctk.CTkButton(footer, text="Batal", width=90, height=36, fg_color="transparent", border_width=1, text_color=("gray20", "gray80"), command=self.destroy)
        btn_cancel.pack(side="right")

    def save_fsb(self):
        data = {k: w.get().strip() for k, w in self.entries.items()}
        if self.on_save_callback:
            ident = {
                "WO": self.initial_data.get("WO"),
                "Serial Number": self.initial_data.get("Serial Number"),
                "Date FSB": self.initial_data.get("Date FSB")
            } if self.mode == "edit" else None
            self.on_save_callback(data, mode=self.mode, original_ident=ident)
        self.destroy()

    def confirm_delete(self):
        ident = {
            "WO": self.initial_data.get("WO"),
            "Serial Number": self.initial_data.get("Serial Number"),
            "Date FSB": self.initial_data.get("Date FSB")
        }
        if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus catatan FSB WO '{ident['WO']}' dari worksheet History FSB?", icon="warning", parent=self):
            if self.on_delete_callback:
                self.on_delete_callback(ident)
            self.destroy()


# -----------------------------------------------------------------------------
# 5. SCRAP FORM DIALOG (HISTORY SCRAP ADD/EDIT)
# -----------------------------------------------------------------------------
class SCRAPFormDialog(ctk.CTkToplevel):
    def __init__(self, parent, mode="add", initial_data=None, dropdown_options=None, on_save_callback=None, on_delete_callback=None):
        super().__init__(parent)
        self.mode = mode
        self.initial_data = initial_data or {}
        self.dropdown_options = dropdown_options or {}
        self.on_save_callback = on_save_callback
        self.on_delete_callback = on_delete_callback

        is_edit = (mode == "edit")
        title_text = "✏️ Edit Riwayat Scrap / FRA" if is_edit else "➕ Catat Baterai Scrap / FRA"
        self.title(title_text)
        self.geometry("660x520")
        self.minsize(580, 420)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(header, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        sub_text = "Data disimpan ke worksheet 'History SCRAP' pada workbook aktif"
        ctk.CTkLabel(header, text=sub_text, font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 10))

        form_frame = ctk.CTkScrollableFrame(self, corner_radius=10)
        form_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        self.entries = {}

        def create_row(label_text, var_name, widget_type="entry", options=None, default_val=""):
            box = ctk.CTkFrame(form_frame, fg_color="transparent")
            box.pack(fill="x", pady=4)
            ctk.CTkLabel(box, text=label_text, width=170, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(4, 8))
            if widget_type == "combobox":
                opts = list(options or [])
                val = str(default_val) if default_val in opts else (opts[0] if opts else "")
                w = ctk.CTkComboBox(box, values=opts, width=280, height=32)
                if val:
                    w.set(val)
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w
            else:
                w = ctk.CTkEntry(box, width=280, height=32)
                if default_val:
                    w.insert(0, str(default_val))
                w.pack(side="left", fill="x", expand=True)
                self.entries[var_name] = w

        create_row("Tanggal Scrap/FRA*", "Date SCRAP", "entry", default_val=self.initial_data.get("Date SCRAP", datetime.now().strftime("%d %B %Y")))
        create_row("Nomor Work Order (WO)*", "WO", "entry", default_val=self.initial_data.get("WO", ""))
        create_row("Part Number (P/N)*", "Part Number", "combobox", options=self.dropdown_options.get("Part Number", []), default_val=self.initial_data.get("Part Number", ""))
        create_row("Serial Number (S/N)*", "Serial Number", "entry", default_val=self.initial_data.get("Serial Number", ""))
        create_row("Customer / Maskapai*", "Customer", "combobox", options=self.dropdown_options.get("Customer", []), default_val=self.initial_data.get("Customer", ""))
        create_row("Status Perawatan*", "Status", "combobox", options=self.dropdown_options.get("Status", []), default_val=self.initial_data.get("Status", "Overhaul"))
        create_row("TAT (Hari Kerja)", "TAT", "entry", default_val=str(self.initial_data.get("TAT", "10")))
        create_row("Klasifikasi (SCRAP / FRA)*", "Position", "combobox", options=["SCRAP", "FRA"], default_val=self.initial_data.get("Position", "SCRAP"))
        create_row("Alasan Teknis Scrap / DMF Remark", "Remark", "entry", default_val=self.initial_data.get("Remark", ""))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(6, 16))

        if is_edit and self.on_delete_callback:
            btn_del = ctk.CTkButton(footer, text="🗑️ Hapus Catatan", fg_color="#ef4444", hover_color="#dc2626", width=120, height=36, command=self.confirm_delete)
            btn_del.pack(side="left")

        btn_save = ctk.CTkButton(footer, text="💾 Simpan ke History SCRAP", font=ctk.CTkFont(size=12, weight="bold"), fg_color="#059669", hover_color="#047857", height=36, command=self.save_scrap)
        btn_save.pack(side="right", padx=(8, 0))

        btn_cancel = ctk.CTkButton(footer, text="Batal", width=90, height=36, fg_color="transparent", border_width=1, text_color=("gray20", "gray80"), command=self.destroy)
        btn_cancel.pack(side="right")

    def save_scrap(self):
        data = {k: w.get().strip() for k, w in self.entries.items()}
        if self.on_save_callback:
            ident = {
                "WO": self.initial_data.get("WO"),
                "Serial Number": self.initial_data.get("Serial Number"),
                "Date SCRAP": self.initial_data.get("Date SCRAP")
            } if self.mode == "edit" else None
            self.on_save_callback(data, mode=self.mode, original_ident=ident)
        self.destroy()

    def confirm_delete(self):
        ident = {
            "WO": self.initial_data.get("WO"),
            "Serial Number": self.initial_data.get("Serial Number"),
            "Date SCRAP": self.initial_data.get("Date SCRAP")
        }
        if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus catatan Scrap WO '{ident['WO']}' dari worksheet History SCRAP?", icon="warning", parent=self):
            if self.on_delete_callback:
                self.on_delete_callback(ident)
            self.destroy()


# -----------------------------------------------------------------------------
# 6. GLOSSARY DIALOG
# -----------------------------------------------------------------------------
GLOSSARY_ITEMS = [
    ("WO", "Work Order", "Nomor surat perintah kerja resmi untuk perawatan komponen baterai pesawat."),
    ("PIR", "Part In Receive", "Tanggal saat baterai diterima di bengkel TCC-2 dari maskapai / line."),
    ("TAT", "Turn Around Time", "Durasi waktu pengerjaan baterai dalam hari kerja (Target: <= 10 hari)."),
    ("FSB", "Finish Serviceable Battery", "Baterai selesai dirawat, lulus uji (Discharge 1C, dll), siap pasang dengan ARC/Form 1."),
    ("SCRAP", "Scrap (Afkir Permanen)", "Baterai rusak permanen / tidak ekonomis diperbaiki (Beyond Economical Repair)."),
    ("FRA", "Finish Return As Is", "Baterai dikembalikan ke maskapai tanpa servis penuh atas instruksi maskapai."),
    ("UR", "Under Repair", "Unit sedang ditangani aktif teknisi (Disassembly, Boost Charge, Capacity Test)."),
    ("WR", "Waiting Repair", "Unit antre menunggu giliran meja/charging station (Waiting Queue)."),
    ("WI", "Waiting Inspection / Instruction", "Menunggu inspeksi QA atau instruksi teknis lanjutan."),
    ("TECH", "Technical Evaluation", "Evaluasi lanjutan tim Engineering / pabrikan."),
    ("FSC", "Finish Shop Check", "Selesai pemeriksaan teknisi awal, rekomendasi afkir/scrap."),
    ("SRAM/SREW", "Shop Investigation & Spare Wait", "Menunggu suku cadang baru (sel pengganti, sensor suhu, dll)."),
    ("Reguler", "Reguler Maintenance", "Pemeriksaan dan servis standar berkala sesuai CMM."),
    ("Overhaul", "Overhaul Maintenance", "Pembongkaran penuh seluruh sel, uji kapasitas, dan penggantian komponen."),
    ("Periodic", "Periodic Inspection", "Pemeriksaan berkala 3-6 bulanan."),
    ("Restoration", "Cell Restoration", "Siklus reconditioning berulang untuk memulihkan kapasitas sel.")
]


class GlossaryDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Glosarium Istilah & Singkatan Teknis Battery Shop TCC-2")
        self.geometry("680x500")
        self.minsize(600, 420)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(16, 10))

        ctk.CTkLabel(header, text="📖 Panduan Singkatan & Glosarium Teknis Bengkel", font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(header, text="Standar nomenklatur operasional Battery Shop Unit TCC-2 PT GMF AeroAsia Tbk", font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 10))

        search_box = ctk.CTkFrame(self, fg_color="transparent")
        search_box.pack(fill="x", padx=16, pady=4)
        self.search_entry = ctk.CTkEntry(search_box, placeholder_text="🔍 Cari istilah atau kode singkatan...", height=34)
        self.search_entry.pack(fill="x")
        self.search_entry.bind("<KeyRelease>", self.filter_terms)

        scroll_box = ctk.CTkScrollableFrame(self, corner_radius=10)
        scroll_box.pack(fill="both", expand=True, padx=16, pady=10)
        self.scroll_box = scroll_box

        self.populate(GLOSSARY_ITEMS)

    def populate(self, items):
        for w in self.scroll_box.winfo_children():
            w.destroy()
        for code, name, desc in items:
            card = ctk.CTkFrame(self.scroll_box, fg_color=("#ffffff", "#1e293b"), border_width=1, border_color=("#e2e8f0", "#334155"), corner_radius=8)
            card.pack(fill="x", pady=4, padx=2)
            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=12, pady=(8, 2))
            ctk.CTkLabel(top, text=code, font=ctk.CTkFont(size=13, weight="bold"), text_color=("#0284c7", "#38bdf8")).pack(side="left")
            ctk.CTkLabel(top, text=f" — {name}", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
            ctk.CTkLabel(card, text=desc, font=ctk.CTkFont(size=11), wraplength=640, justify="left", text_color=("#475569", "#cbd5e1")).pack(anchor="w", padx=12, pady=(0, 8))

    def filter_terms(self, event=None):
        q = self.search_entry.get().strip().lower()
        if not q:
            self.populate(GLOSSARY_ITEMS)
            return
        filtered = [item for item in GLOSSARY_ITEMS if q in item[0].lower() or q in item[1].lower() or q in item[2].lower()]
        self.populate(filtered)


# -----------------------------------------------------------------------------
# 7. EXCEL ACCESS & BACKUP DIALOG
# -----------------------------------------------------------------------------
class ExcelAccessDialog(ctk.CTkToplevel):
    def __init__(self, parent, excel_file_path, on_file_changed_callback=None, on_reload_callback=None):
        super().__init__(parent)
        self.excel_file_path = Path(excel_file_path)
        self.on_file_changed_callback = on_file_changed_callback
        self.on_reload_callback = on_reload_callback

        self.title("Pusat Pengelolaan Master Excel & Backup Data")
        self.geometry("680x500")
        self.minsize(600, 420)
        self.after(10, lambda: bring_window_to_front(self))

        header = ctk.CTkFrame(self, corner_radius=10, fg_color=("white", "#1e293b"))
        header.pack(fill="x", padx=16, pady=(16, 10))

        ctk.CTkLabel(header, text="📑 Manajemen Workbook Master & Cadangan", font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(header, text="Kelola file master Excel aktif, buka folder, dan pulihkan cadangan otomatis.", font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 10))

        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=4)

        # File Status Card
        card1 = ctk.CTkFrame(body, fg_color=("#ffffff", "#1e293b"), border_width=1, border_color=("#e2e8f0", "#334155"), corner_radius=10)
        card1.pack(fill="x", pady=6)

        try:
            rel_path = self.excel_file_path.resolve().relative_to(BASE_DIR.resolve())
            loc_str = f"[Folder Aplikasi] \\ {rel_path}  (Portable / Otomatis Menyesuaikan Laptop)"
        except Exception:
            loc_str = f"{self.excel_file_path.name}"

        info_txt = (
            f"• Nama File : {self.excel_file_path.name}\n"
            f"• Lokasi    : {loc_str}\n"
            f"  Path Lengkap: {self.excel_file_path}\n"
            f"• Status    : {'Ada di disk' if self.excel_file_path.exists() else 'Tidak ditemukan!'}"
        )
        ctk.CTkLabel(card1, text=info_txt, font=ctk.CTkFont(family="Consolas", size=11), justify="left").pack(anchor="w", padx=14, pady=(0, 10))

        btn_row1 = ctk.CTkFrame(card1, fg_color="transparent")
        btn_row1.pack(fill="x", padx=14, pady=(0, 12))

        ctk.CTkButton(btn_row1, text="📂 Buka di File Explorer", height=32, fg_color="#334155", hover_color="#475569", command=self.open_in_explorer).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row1, text="📁 Pilih File Master Lain", height=32, fg_color="#2563eb", hover_color="#1d4ed8", command=self.browse_another).pack(side="left")

        # Backups Card
        card2 = ctk.CTkFrame(body, fg_color=("#ffffff", "#1e293b"), border_width=1, border_color=("#e2e8f0", "#334155"), corner_radius=10)
        card2.pack(fill="x", pady=6)

        ctk.CTkLabel(card2, text="Cadangan Otomatis (Backups)", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(12, 4))

        backups = list_backups()
        backup_txt = f"Ditemukan {len(backups)} cadangan otomatis di folder 'backups/'."
        ctk.CTkLabel(card2, text=backup_txt, font=ctk.CTkFont(size=11), text_color=("gray40", "gray70")).pack(anchor="w", padx=14, pady=(0, 8))

        btn_row2 = ctk.CTkFrame(card2, fg_color="transparent")
        btn_row2.pack(fill="x", padx=14, pady=(0, 12))

        ctk.CTkButton(btn_row2, text="🔄 Pulihkan dari Backup", height=32, fg_color="#d97706", hover_color="#b45309", command=self.restore_dialog).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row2, text="📂 Buka Folder Backups", height=32, fg_color="#334155", hover_color="#475569", command=self.open_backups_folder).pack(side="left")

    def open_in_explorer(self):
        folder = self.excel_file_path.parent
        if sys.platform == "win32":
            subprocess.run(["explorer", str(folder)])

    def open_backups_folder(self):
        bdir = get_backup_directory()
        if sys.platform == "win32":
            subprocess.run(["explorer", str(bdir)])

    def browse_another(self):
        f = filedialog.askopenfilename(
            parent=self,
            title="Pilih File Master Excel",
            initialdir=str(self.excel_file_path.parent),
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if f and self.on_file_changed_callback:
            chosen = Path(f).resolve()
            is_valid, errors, _ = validate_workbook_schema(chosen)
            if not is_valid:
                messagebox.showerror(
                    "Validasi Gagal",
                    f"File '{chosen.name}' tidak dapat digunakan sebagai master:\n\n" + "\n".join([f"• {e}" for e in errors]),
                    parent=self
                )
                return
            set_master_excel_path(chosen)
            self.on_file_changed_callback(chosen)
            self.destroy()

    def restore_dialog(self):
        backups = list_backups()
        if not backups:
            messagebox.showinfo("Tidak Ada Backup", "Belum ada file backup yang tersimpan.", parent=self)
            return

        win = ctk.CTkToplevel(self)
        win.title("Pilih Backup untuk Dipulihkan")
        win.geometry("540x380")
        win.after(10, lambda: bring_window_to_front(win))

        ctk.CTkLabel(win, text="Pilih cadangan yang ingin dipulihkan:", font=ctk.CTkFont(size=12, weight="bold")).pack(padx=16, pady=(16, 8), anchor="w")

        lb = tk.Listbox(win, font=("Consolas", 10))
        lb.pack(fill="both", expand=True, padx=16, pady=4)
        for b in backups:
            lb.insert(tk.END, f"{b['filename']} ({b['size_kb']} KB)")

        def do_restore():
            sel = lb.curselection()
            if not sel:
                return
            target_backup = backups[sel[0]]
            if messagebox.askyesno("Konfirmasi Pulihkan", f"Apakah Anda yakin ingin memulihkan data dari:\n{target_backup['filename']}?\n\nFile master aktif saat ini akan dicadangkan terlebih dahulu.", icon="warning", parent=win):
                ok, msg = restore_backup(target_backup['path'], self.excel_file_path)
                if ok:
                    messagebox.showinfo("Berhasil", "Data berhasil dipulihkan dari backup!", parent=win)
                    win.destroy()
                    if self.on_reload_callback:
                        self.on_reload_callback()
                else:
                    messagebox.showerror("Gagal", f"Gagal memulihkan backup: {msg}", parent=win)

        btn_act = ctk.CTkButton(win, text="Pulihkan Sekarang", fg_color="#d97706", hover_color="#b45309", height=34, command=do_restore)
        btn_act.pack(fill="x", padx=16, pady=12)


# -----------------------------------------------------------------------------
# 8. BATTERY CALENDAR FILTER MODAL (USER-FRIENDLY & MODERNIZED)
# -----------------------------------------------------------------------------
class BatteryCalendarDialog(ctk.CTkToplevel):
    def __init__(self, parent, df_daily, active_date, on_date_selected_callback, on_clear_callback):
        super().__init__(parent)
        _ensure_tkcalendar()
        self.df_daily = df_daily if df_daily is not None else pd.DataFrame()
        self.active_date = active_date
        self.on_date_selected_callback = on_date_selected_callback
        self.on_clear_callback = on_clear_callback

        self.title("Filter Kalender PIR (Tanggal Masuk Unit) - Battery Shop TCC-2")
        self.geometry("600x540")
        self.minsize(520, 440)
        self.after(10, lambda: bring_window_to_front(self))

        # Index PIR dates from df_daily
        self.pir_dates = {}
        if not self.df_daily.empty and "PIR" in self.df_daily.columns:
            for _, row in self.df_daily.iterrows():
                val = row.get("PIR")
                d_obj = parse_date_flexible(val)
                if d_obj:
                    self.pir_dates.setdefault(d_obj, []).append(row)

        self._build_ui()

    def _build_ui(self):
        # Header Box
        header = ctk.CTkFrame(
            self,
            corner_radius=12,
            fg_color=("#ffffff", "#1e293b"),
            border_width=1,
            border_color=("#e2e8f0", "#334155")
        )
        header.pack(fill="x", padx=16, pady=(14, 8))

        top_row = ctk.CTkFrame(header, fg_color="transparent")
        top_row.pack(fill="x", padx=14, pady=(12, 2))

        ctk.CTkLabel(
            top_row,
            text="📅 Filter Kalender Part In Receive (PIR)",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        ).pack(side="left")

        total_units = sum(len(r) for r in self.pir_dates.values())
        ctk.CTkLabel(
            top_row,
            text=f"📊 {total_units} Unit di {len(self.pir_dates)} Tanggal",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#0284c7", "#38bdf8")
        ).pack(side="right")

        ctk.CTkLabel(
            header,
            text="Pilih tanggal kedatangan baterai untuk memfilter unit harian (WIP). Tanggal berlatar biru memiliki unit terdaftar.",
            font=ctk.CTkFont(size=11),
            text_color=("#64748b", "#94a3b8"),
            wraplength=540,
            justify="left"
        ).pack(anchor="w", padx=14, pady=(0, 10))

        # Active filter badge if exists
        if self.active_date:
            active_bar = ctk.CTkFrame(header, fg_color=("#ecfdf5", "#064e3b"), corner_radius=8, border_width=1, border_color=("#a7f3d0", "#047857"))
            active_bar.pack(fill="x", padx=14, pady=(0, 10))
            ctk.CTkLabel(
                active_bar,
                text=f"🟢 Filter Aktif Saat Ini: {self.active_date}",
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=("#065f46", "#6ee7b7")
            ).pack(side="left", padx=10, pady=4)
            ctk.CTkButton(
                active_bar,
                text="✕ Hapus Filter",
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color="#ef4444",
                hover_color="#dc2626",
                height=24,
                width=85,
                command=self.clear_filter
            ).pack(side="right", padx=6, pady=4)

        # Quick Presets Bar
        presets_bar = ctk.CTkFrame(self, fg_color="transparent")
        presets_bar.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            presets_bar,
            text="⚡ Pintasan:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#475569", "#94a3b8")
        ).pack(side="left", padx=(0, 6))

        if self.pir_dates:
            latest_d = max(self.pir_dates.keys())
            latest_str = f"{latest_d.day} {MONTHS_FULL_ID[latest_d.month - 1]}"
            ctk.CTkButton(
                presets_bar,
                text=f"🎯 Unit Terbaru ({latest_str})",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#0284c7",
                hover_color="#0369a1",
                height=30,
                command=lambda: self.select_specific_date(latest_d)
            ).pack(side="left", padx=3)

        ctk.CTkButton(
            presets_bar,
            text="📌 Hari Ini",
            font=ctk.CTkFont(size=11),
            fg_color=("#e2e8f0", "#334155"),
            text_color=("#1e293b", "#f8fafc"),
            hover_color=("#cbd5e1", "#475569"),
            height=30,
            command=lambda: self.select_specific_date(date.today())
        ).pack(side="left", padx=3)

        ctk.CTkButton(
            presets_bar,
            text="⏮️ Kemarin",
            font=ctk.CTkFont(size=11),
            fg_color=("#e2e8f0", "#334155"),
            text_color=("#1e293b", "#f8fafc"),
            hover_color=("#cbd5e1", "#475569"),
            height=30,
            command=lambda: self.select_specific_date(date.today() - timedelta(days=1))
        ).pack(side="left", padx=3)

        # Calendar Container
        cal_container = ctk.CTkFrame(
            self,
            fg_color=("#ffffff", "#1e293b"),
            border_width=1,
            border_color=("#e2e8f0", "#334155"),
            corner_radius=10
        )
        cal_container.pack(fill="both", expand=True, padx=16, pady=4)

        if HAS_TKCALENDAR:
            self.cal = Calendar(
                cal_container,
                selectmode="day",
                showweeknumbers=False,
                firstweekday="monday",
                date_pattern="y-mm-dd",
                font=("Segoe UI", 10),
                background="#0284c7",
                foreground="#ffffff",
                headersbackground="#f1f5f9",
                headersforeground="#334155",
                selectbackground="#2563eb",
                selectforeground="#ffffff",
                normalbackground="#ffffff",
                normalforeground="#0f172a",
                weekendbackground="#f8fafc",
                weekendforeground="#64748b",
                othermonthbackground="#f8fafc",
                othermonthforeground="#cbd5e1",
                othermonthwebackground="#f8fafc",
                othermonthweforeground="#cbd5e1",
                bordercolor="#e2e8f0",
                cursor="hand2"
            )
            self.cal.pack(fill="both", expand=True, padx=8, pady=8)

            # Highlight all dates that have PIR unit records
            for d_obj, rows in self.pir_dates.items():
                self.cal.calevent_create(d_obj, f"{len(rows)} Unit", "has_unit")
            self.cal.tag_config("has_unit", background="#e0f2fe", foreground="#0369a1")

            # Determine initial focused date
            init_d = None
            if self.active_date:
                init_d = parse_date_flexible(self.active_date)
            if not init_d and self.pir_dates:
                init_d = max(self.pir_dates.keys())
            if not init_d:
                init_d = date.today()

            self.cal.see(init_d)
            self.cal.selection_set(init_d)

            self.cal.bind("<<CalendarSelected>>", self.on_calendar_selected)
            self.cal.bind("<Double-1>", lambda e: self.apply_calendar())
        else:
            ctk.CTkLabel(
                cal_container,
                text="Modul tkcalendar tidak terpasang. Gunakan kolom pencarian PIR langsung di tabel.",
                font=ctk.CTkFont(size=12)
            ).pack(pady=30)

        # Interactive Live Selection Preview Card
        self.preview_card = ctk.CTkFrame(
            self,
            fg_color=("#f8fafc", "#0f172a"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            corner_radius=10
        )
        self.preview_card.pack(fill="x", padx=16, pady=(4, 6))

        p_row1 = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        p_row1.pack(fill="x", padx=12, pady=(10, 2))

        self.lbl_selected_header = ctk.CTkLabel(
            p_row1,
            text="📅 Tanggal Terpilih: -",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0f172a", "#f8fafc")
        )
        self.lbl_selected_header.pack(side="left")

        self.lbl_unit_badge = ctk.CTkLabel(
            p_row1,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        )
        self.lbl_unit_badge.pack(side="right")

        self.lbl_unit_breakdown = ctk.CTkLabel(
            self.preview_card,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#475569", "#94a3b8"),
            justify="left",
            wraplength=540
        )
        self.lbl_unit_breakdown.pack(anchor="w", padx=12, pady=(0, 6))

        # User Guide / Tip
        ctk.CTkLabel(
            self.preview_card,
            text="💡 Tip: Klik 2x pada tanggal untuk langsung memfilter tabel. Tanggal biru muda memiliki unit masuk.",
            font=ctk.CTkFont(size=10, slant="italic"),
            text_color=("#64748b", "#94a3b8")
        ).pack(anchor="w", padx=12, pady=(0, 8))

        # Footer Actions
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(4, 14))

        btn_clear = ctk.CTkButton(
            footer,
            text="🔄 Reset / Tampilkan Semua",
            font=ctk.CTkFont(size=11, weight="bold"),
            height=36,
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.clear_filter
        )
        btn_clear.pack(side="left")

        self.btn_apply = ctk.CTkButton(
            footer,
            text="✔ Terapkan Filter Tanggal",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=36,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self.apply_calendar
        )
        self.btn_apply.pack(side="right")

        btn_cancel = ctk.CTkButton(
            footer,
            text="Batal",
            font=ctk.CTkFont(size=11),
            width=80,
            height=36,
            fg_color="#64748b",
            hover_color="#475569",
            command=self.destroy
        )
        btn_cancel.pack(side="right", padx=8)

        # Trigger initial preview update
        if HAS_TKCALENDAR and hasattr(self, "cal"):
            self.on_calendar_selected()

    def select_specific_date(self, target_date: date):
        """Pintasan untuk langsung memilih tanggal tertentu di kalender."""
        if HAS_TKCALENDAR and hasattr(self, "cal"):
            try:
                self.cal.see(target_date)
                self.cal.selection_set(target_date)
                self.on_calendar_selected()
            except Exception:
                pass

    def on_calendar_selected(self, event=None):
        """Updates live selection preview card dynamically upon user clicking any date."""
        if not HAS_TKCALENDAR or not hasattr(self, "cal"):
            return

        sel_str = self.cal.get_date()
        try:
            d_obj = datetime.strptime(sel_str, "%Y-%m-%d").date()
            day_name = DAYS_FULL_ID[d_obj.weekday()]
            display_title = f"📅 {day_name}, {d_obj.day} {MONTHS_FULL_ID[d_obj.month - 1]} {d_obj.year}"
            self.lbl_selected_header.configure(text=display_title)

            rows = self.pir_dates.get(d_obj, [])
            count = len(rows)

            if count > 0:
                self.lbl_unit_badge.configure(
                    text=f"🟢 {count} Unit Masuk (PIR)",
                    text_color="#15803d"
                )

                # Breakdown customers and sample WOs
                cust_counter = Counter([str(r.get("Customer", "")).strip() for r in rows if str(r.get("Customer", "")).strip()])
                c_str = ", ".join([f"{c} ({n})" for c, n in cust_counter.most_common(3)])
                if len(cust_counter) > 3:
                    c_str += f" +{len(cust_counter) - 3} lainnya"

                wos = [str(r.get("WO", "")).strip() for r in rows if str(r.get("WO", "")).strip()]
                wo_str = ", ".join(wos[:4])
                if len(wos) > 4:
                    wo_str += f" (+{len(wos) - 4} unit)"

                self.lbl_unit_breakdown.configure(
                    text=f"• Maskapai: {c_str}\n• Contoh WO: {wo_str}"
                )
                self.btn_apply.configure(text=f"✔ Terapkan Filter ({count} Unit)")
            else:
                self.lbl_unit_badge.configure(
                    text="⚪ 0 Unit Masuk",
                    text_color="#64748b"
                )
                self.lbl_unit_breakdown.configure(
                    text="Tidak ada unit baterai tercatat masuk pada tanggal ini.\n(Tabel harian akan kosong jika memfilter tanggal ini)"
                )
                self.btn_apply.configure(text="✔ Terapkan Tanggal Ini (0 Unit)")
        except Exception:
            self.lbl_selected_header.configure(text=f"📅 Tanggal Terpilih: {sel_str}")

    def apply_calendar(self):
        if HAS_TKCALENDAR and hasattr(self, "cal"):
            sel_date_str = self.cal.get_date()
            try:
                dt = datetime.strptime(sel_date_str, "%Y-%m-%d").date()
                idx = dt.month - 1
                formatted = f"{dt.day} {MONTHS_FULL_ID[idx]} {dt.year}"
            except Exception:
                formatted = sel_date_str
            if self.on_date_selected_callback:
                self.on_date_selected_callback(formatted)
        self.destroy()

    def clear_filter(self):
        if self.on_clear_callback:
            self.on_clear_callback()
        self.destroy()



# -----------------------------------------------------------------------------
# 9. ROUTE CONFIRMATION DIALOGS (SHARED BASE)
# -----------------------------------------------------------------------------
class _BaseRouteDialog(ctk.CTkToplevel):
    """Shared base modal for routing units out of Daily Battery."""
    def __init__(self, parent, unit_data: dict, title: str, header_title: str, header_sub: str, fg_theme: tuple, border_theme: tuple, text_theme: tuple):
        super().__init__(parent)
        self.unit_data = unit_data
        self.title(title)
        self.geometry("540x460")
        self.minsize(480, 400)
        self.after(10, lambda: bring_window_to_front(self))

        # Header
        header = ctk.CTkFrame(self, fg_color=fg_theme, corner_radius=10, border_width=1, border_color=border_theme)
        header.pack(fill="x", padx=16, pady=(16, 12))
        ctk.CTkLabel(header, text=header_title, font=ctk.CTkFont(size=14, weight="bold"), text_color=text_theme[0]).pack(anchor="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(header, text=header_sub, font=ctk.CTkFont(size=11), text_color=text_theme[1]).pack(anchor="w", padx=14, pady=(0, 10))

        # Details Box
        details_box = ctk.CTkFrame(self, fg_color=("white", "#1e293b"), corner_radius=10, border_width=1, border_color=("#e2e8f0", "#334155"))
        details_box.pack(fill="x", padx=16, pady=(0, 12))

        wo = str(unit_data.get("WO", "-")).strip()
        pn = str(unit_data.get("Part Number", "-")).strip()
        sn = str(unit_data.get("Serial Number", "-")).strip()
        cust = str(unit_data.get("Customer", unit_data.get("Costumer", "-"))).strip()

        for label, val in [
            ("Nomor Work Order (WO):", wo),
            ("Part Number (P/N):", pn),
            ("Serial Number (S/N):", sn),
            ("Customer / Maskapai:", cust),
        ]:
            r_frame = ctk.CTkFrame(details_box, fg_color="transparent")
            r_frame.pack(fill="x", padx=12, pady=3)
            ctk.CTkLabel(r_frame, text=label, width=170, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
            ctk.CTkLabel(r_frame, text=val, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left")

        # Form Inputs Container
        self.input_box = ctk.CTkFrame(self, fg_color="transparent")
        self.input_box.pack(fill="x", padx=16, pady=(0, 12))

    def create_input_row(self, label: str, widget_factory):
        row = ctk.CTkFrame(self.input_box, fg_color="transparent")
        row.pack(fill="x", pady=4)
        ctk.CTkLabel(row, text=label, width=180, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        w = widget_factory(row)
        w.pack(side="left", fill="x", expand=True)
        return w

    def create_footer(self, confirm_text: str, confirm_cmd, confirm_color: str, hover_color: str):
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(4, 16))
        ctk.CTkButton(footer, text="Batal", fg_color=("gray85", "gray30"), text_color=("black", "white"), hover_color=("gray75", "gray40"), width=100, height=36, command=self.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(footer, text=confirm_text, fg_color=confirm_color, hover_color=hover_color, text_color="white", font=ctk.CTkFont(weight="bold"), height=36, command=confirm_cmd).pack(side="right")


class RouteToFSBDialog(_BaseRouteDialog):
    def __init__(self, parent, unit_data: dict, on_confirm_callback):
        super().__init__(
            parent, unit_data,
            title="⚡ Konfirmasi Perpindahan Unit ke History FSB (FINAL CHARGE)",
            header_title="⚡ Status: FINAL CHARGE (Serviceable Ready)",
            header_sub="Unit akan dipindahkan secara otomatis dari Daily Battery ke History FSB.",
            fg_theme=("#ecfdf5", "#064e3b"),
            border_theme=("#a7f3d0", "#047857"),
            text_theme=("#065f46", "#a7f3d0")
        )
        self.on_confirm_callback = on_confirm_callback

        self.ent_date = self.create_input_row(
            "Tanggal Selesai (Date FSB)*:",
            lambda p: ctk.CTkEntry(p, height=32)
        )
        self.ent_date.insert(0, format_date_with_month_name(date.today()))

        self.ent_remark = self.create_input_row(
            "Catatan / Form 1 Remark:",
            lambda p: ctk.CTkEntry(p, placeholder_text="Opsional (misal: Selesai Uji Kapasitas, Siap Pasang)", height=32)
        )
        self.ent_remark.insert(0, str(unit_data.get("Remark", "")).strip())

        self.create_footer("✔ Pindahkan ke History FSB", self.confirm_route, "#059669", "#047857")

    def confirm_route(self):
        dt_val = self.ent_date.get().strip()
        rem_val = self.ent_remark.get().strip()
        if not dt_val:
            messagebox.showerror("Validasi Gagal", "Tanggal Selesai (Date FSB) wajib diisi.", parent=self)
            return
        wo = str(self.unit_data.get("WO", "")).strip()
        self.destroy()
        if self.on_confirm_callback:
            self.on_confirm_callback(wo, dt_val, rem_val)


class RouteToSCRAPDialog(_BaseRouteDialog):
    def __init__(self, parent, unit_data: dict, on_confirm_callback):
        super().__init__(
            parent, unit_data,
            title="⚠️ Konfirmasi Perpindahan Unit ke History SCRAP / FRA",
            header_title="⚠️ Status: SCRAP / FRA (Disposisi Tidak Laik Pakai)",
            header_sub="Unit akan dihapus dari Daily Battery dan diarsipkan permanen ke History SCRAP.",
            fg_theme=("#fef2f2", "#450a0a"),
            border_theme=("#fecaca", "#991b1b"),
            text_theme=("#991b1b", "#fca5a5")
        )
        self.on_confirm_callback = on_confirm_callback

        self.ent_date = self.create_input_row(
            "Tanggal SCRAP / FRA*:",
            lambda p: ctk.CTkEntry(p, height=32)
        )
        self.ent_date.insert(0, format_date_with_month_name(date.today()))

        self.cmb_pos = self.create_input_row(
            "Disposisi Posisi*:",
            lambda p: ctk.CTkComboBox(p, values=["SCRAP", "FRA"], state="readonly", height=32)
        )
        init_pos = str(unit_data.get("Position", "SCRAP")).strip().upper()
        self.cmb_pos.set(init_pos if init_pos in ("SCRAP", "FRA") else "SCRAP")

        self.ent_reason = self.create_input_row(
            "Alasan Scrap / DMF Remark*:",
            lambda p: ctk.CTkEntry(p, placeholder_text="Alasan teknis (misal: Sel bocor, Kapasitas <80%, DMF)", height=32)
        )
        curr_rem = str(unit_data.get("Remark", "")).strip()
        self.ent_reason.insert(0, curr_rem if curr_rem else "DMF")

        self.create_footer("⚠️ Pindahkan ke History SCRAP", self.confirm_route, "#dc2626", "#b91c1c")

    def confirm_route(self):
        dt_val = self.ent_date.get().strip()
        pos_val = self.cmb_pos.get().strip().upper()
        reason_val = self.ent_reason.get().strip()
        if not dt_val:
            messagebox.showerror("Validasi Gagal", "Tanggal SCRAP wajib diisi.", parent=self)
            return
        if not reason_val:
            messagebox.showerror("Validasi Gagal", "Alasan Scrap / DMF Remark wajib diisi.", parent=self)
            return
        wo = str(self.unit_data.get("WO", "")).strip()
        self.destroy()
        if self.on_confirm_callback:
            self.on_confirm_callback(wo, dt_val, reason_val, pos_val)
