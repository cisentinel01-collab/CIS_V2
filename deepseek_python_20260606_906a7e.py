"""
نظام إدارة المخازن المؤسسي
AMERICAN MARINE SERVICES FREE ZONE
Enterprise Warehouse ERP System — Production v2.0
Full Arabic RTL | Multi-Warehouse | Complete Traceability
"""

import sys
import os
import sqlite3
import hashlib
import json
import shutil
import io
import base64
from datetime import datetime, date
from pathlib import Path
import threading
import time

# PySide6
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

# ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas as rl_canvas

# Pandas / OpenPyXL
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Fill, PatternFill, Alignment, Border, Side

# Pillow
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance

# QR Code
import qrcode

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS & PATHS
# ─────────────────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "database.db"
ASSETS_DIR = APP_DIR / "logo"
LOGO_PATH = ASSETS_DIR / "logo.png"
BACKUP_DIR = APP_DIR / "backups"
ATTACHMENTS_DIR = APP_DIR / "attachments"
BACKUP_DIR.mkdir(exist_ok=True)
ATTACHMENTS_DIR.mkdir(exist_ok=True)

COMPANY_NAME = "AMERICAN MARINE SERVICES FREE ZONE"
APP_TITLE = f"{COMPANY_NAME}\nنظام إدارة المخازن المؤسسي"

# Arabic labels
AR = {
    "dashboard": "لوحة التحكم",
    "products": "المنتجات",
    "warehouses": "المخازن",
    "suppliers": "الموردين",
    "stock_in": "إضافة للمخزون",
    "stock_out": "صرف من المخزون",
    "transfers": "التحويلات",
    "reports": "التقارير",
    "settings": "الإعدادات",
    "users": "المستخدمون",
    "notifications": "الإشعارات",
    "audit": "الجرد",
    "logout": "تسجيل الخروج",
    "save": "حفظ",
    "cancel": "إلغاء",
    "delete": "حذف",
    "edit": "تعديل",
    "add": "إضافة",
    "search": "بحث",
    "print": "طباعة",
    "export": "تصدير",
    "approve": "موافقة",
    "reject": "رفض",
    "return": "إرجاع",
    "transfer": "تحويل",
    "pending": "قيد الانتظار",
    "approved": "تم الموافقة",
    "rejected": "مرفوض",
    "delivered": "تم التسليم",
    "returned": "مُرجع",
    "cancelled": "ملغي",
    "in_warehouse": "في المخزن",
    "issued": "تم الصرف",
    "transferred": "محوّل",
    "lost": "مفقود",
    "disposed": "تم التخلص",
    "good": "جيد",
    "damaged": "تالف",
    "warning": "تحذير",
    "critical": "حرج",
    "emergency": "طارئ",
    "analytics": "التحليلات الذكية",
}

# ─────────────────────────────────────────────────────────────────────────────
# THEME
# ─────────────────────────────────────────────────────────────────────────────
COLORS = {
    "bg_dark":       "#05070A",
    "bg_nav":        "#0C0E14",
    "bg_card":       "#0C0E14",
    "bg_card2":      "#141820",
    "accent":        "#C0A060",  # Gold for a premium feel
    "accent2":       "#907040",
    "accent_green":  "#10B981",
    "accent_red":    "#EF4444",
    "accent_orange": "#F59E0B",
    "accent_purple": "#8B5CF6",
    "text_primary":  "#F1F5F9",
    "text_secondary":"#94A3B8",
    "text_muted":    "#64748B",
    "border":        "#1E293B",
    "border2":       "#334155",
    "hover":         "rgba(192, 160, 96, 0.1)",
    "selected":      "rgba(192, 160, 96, 0.2)",
    "glass":         "rgba(255, 255, 255, 0.03)",
    "warning_bg":    "#451A03",
    "critical_bg":   "#450A0A",
    "row_warning":   "#2D1B02",
    "row_critical":  "#1A0505",
}

STYLESHEET = f"""
QMainWindow, QDialog, QWidget {{
    background-color: {COLORS['bg_dark']};
    color: {COLORS['text_primary']};
    font-family: 'Cairo', 'Segoe UI', sans-serif;
    font-size: 14px;
}}
QLabel {{
    color: {COLORS['text_primary']};
    background: transparent;
}}
QPushButton {{
    background-color: {COLORS['accent2']};
    color: white;
    border: 1px solid {COLORS['accent']};
    border-radius: 12px;
    padding: 10px 24px;
    font-weight: 700;
    font-size: 14px;
    min-height: 42px;
}}
QPushButton:hover {{
    background-color: {COLORS['accent']};
    color: {COLORS['bg_dark']};
}}
QPushButton:pressed {{ background-color: #0284C7; }}
QPushButton:disabled {{ background-color: {COLORS['text_muted']}; color: {COLORS['bg_card']}; }}
QPushButton#btn_danger {{
    background-color: {COLORS['accent_red']};
}}
QPushButton#btn_danger:hover {{
    background-color: #E11D48;
}}
QPushButton#btn_success {{
    background-color: {COLORS['accent_green']};
}}
QPushButton#btn_success:hover {{
    background-color: #059669;
}}
QPushButton#btn_warning {{
    background-color: {COLORS['accent_orange']};
}}
QPushButton#btn_warning:hover {{
    background-color: #D97706;
}}
QPushButton#btn_flat {{
    background-color: transparent;
    color: {COLORS['text_secondary']};
    border: 1px solid {COLORS['border2']};
}}
QPushButton#btn_flat:hover {{
    background-color: {COLORS['bg_card2']};
    color: {COLORS['text_primary']};
    border-color: {COLORS['accent']};
}}
QLineEdit, QTextEdit, QPlainTextEdit, QDoubleSpinBox, QDateEdit, QSpinBox {{
    background-color: {COLORS['bg_card2']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border2']};
    border-radius: 12px;
    padding: 12px 18px;
    font-size: 14px;
    selection-background-color: {COLORS['accent']};
    selection-color: {COLORS['bg_dark']};
}}
QLineEdit:focus, QTextEdit:focus, QDoubleSpinBox:focus {{
    border: 1px solid {COLORS['accent']};
    background-color: {COLORS['bg_dark']};
}}
QComboBox {{
    background-color: {COLORS['bg_card2']};
    color: {COLORS['text_primary']};
    border: 2px solid {COLORS['border2']};
    border-radius: 10px;
    padding: 10px 15px;
    min-height: 40px;
}}
QComboBox:focus {{ border-color: {COLORS['accent']}; }}
QComboBox::drop-down {{ border: none; width: 35px; }}
QComboBox::down-arrow {{
    image: none;
    border-left: 6px solid transparent;
    border-right: 6px solid transparent;
    border-top: 7px solid {COLORS['text_secondary']};
    margin-right: 10px;
}}
QTableWidget {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border2']};
    border-radius: 12px;
    gridline-color: {COLORS['border2']};
    selection-background-color: {COLORS['selected']};
    alternate-background-color: {COLORS['bg_dark']};
}}
QHeaderView::section {{
    background-color: {COLORS['bg_nav']};
    color: {COLORS['accent']};
    font-weight: 800;
    padding: 12px;
    border: none;
    border-bottom: 2px solid {COLORS['accent']};
}}
QScrollBar:vertical {{
    background: transparent; width: 10px; border-radius: 5px;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['text_muted']}; border-radius: 5px;
}}
QScrollBar::handle:vertical:hover {{ background: {COLORS['accent']}; }}
QGroupBox {{
    font-weight: 800; font-size: 14px; color: {COLORS['accent']};
    border: 2px solid {COLORS['border2']}; border-radius: 12px;
    margin-top: 15px; padding-top: 20px;
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 15px; padding: 0 10px;
}}
QTabWidget::pane {{ border: 2px solid {COLORS['border2']}; border-radius: 12px; top: -1px; }}
QTabBar::tab {{
    background-color: {COLORS['bg_card2']}; color: {COLORS['text_secondary']};
    padding: 12px 25px; border-top-left-radius: 10px; border-top-right-radius: 10px;
    margin-right: 5px; font-weight: 700;
}}
QTabBar::tab:selected {{ background-color: {COLORS['bg_card']}; color: {COLORS['accent']}; }}
QFrame#card {{
    background-color: {COLORS['bg_card']}; border: 1.5px solid {COLORS['border2']}; border-radius: 15px;
}}
QFrame#separator {{ background: {COLORS['border2']}; max-height: 1px; }}
QStatusBar {{
    background: {COLORS['bg_nav']}; color: {COLORS['text_secondary']};
    font-size: 12px; border-top: 1px solid {COLORS['border2']};
}}
QMessageBox {{ background: {COLORS['bg_card']}; }}
QMessageBox QLabel {{ color: {COLORS['text_primary']}; font-size: 13px; }}
QToolTip {{
    background: {COLORS['bg_card']}; color: {COLORS['text_primary']};
    border: 1px solid {COLORS['accent']}; border-radius: 6px; padding: 6px 10px; font-size: 12px;
}}
"""

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE MANAGER — ENTERPRISE EDITION
# ─────────────────────────────────────────────────────────────────────────────
class DatabaseManager:
    def __init__(self):
        self.db_path = str(DB_PATH)
        self.init_database()
        self.run_migrations()

    def get_conn(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        return conn

    def execute(self, sql, params=()):
        with self.get_conn() as conn:
            cur = conn.execute(sql, params)
            conn.commit()
            return cur

    def executemany(self, sql, params_list):
        with self.get_conn() as conn:
            conn.executemany(sql, params_list)
            conn.commit()

    def fetchall(self, sql, params=()):
        with self.get_conn() as conn:
            return conn.execute(sql, params).fetchall()

    def fetchone(self, sql, params=()):
        with self.get_conn() as conn:
            return conn.execute(sql, params).fetchone()

    def init_database(self):
        """Create all tables using IF NOT EXISTS — safe for existing data."""
        with self.get_conn() as conn:
            conn.executescript("""
            -- USERS
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'Viewer',
                full_name TEXT,
                email TEXT,
                phone TEXT,
                department TEXT,
                last_login TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                is_active INTEGER DEFAULT 1
            );

            -- CATEGORIES
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                description TEXT,
                created_at TEXT DEFAULT (datetime('now'))
            );

            -- WAREHOUSES
            CREATE TABLE IF NOT EXISTS warehouses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                warehouse_code TEXT UNIQUE NOT NULL,
                warehouse_name TEXT NOT NULL,
                warehouse_manager TEXT,
                phone TEXT,
                address TEXT,
                notes TEXT,
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT (datetime('now'))
            );

            -- SUPPLIERS
            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                nationality TEXT,
                phone TEXT,
                email TEXT,
                company TEXT,
                country TEXT,
                address TEXT,
                notes TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                is_active INTEGER DEFAULT 1
            );

            -- PRODUCTS
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                description TEXT,
                category TEXT,
                quantity REAL DEFAULT 0,
                min_quantity REAL DEFAULT 0,
                alert_quantity REAL DEFAULT 0,
                purchase_price REAL DEFAULT 0,
                selling_price REAL DEFAULT 0,
                barcode TEXT,
                warehouse_location TEXT,
                rack_number TEXT,
                warehouse_id INTEGER,
                supplier_id INTEGER,
                image_data TEXT,
                qr_code_data TEXT,
                serial_number TEXT,
                asset_tag TEXT,
                purchase_date TEXT,
                warranty_end TEXT,
                current_owner TEXT,
                current_location TEXT,
                current_status TEXT DEFAULT 'في المخزن',
                is_asset INTEGER DEFAULT 0,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now')),
                is_active INTEGER DEFAULT 1,
                alarms_enabled INTEGER DEFAULT 1,
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id)
            );

            -- STOCK IN (original preserved)
            CREATE TABLE IF NOT EXISTS stock_in (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                supplier_id INTEGER,
                warehouse_id INTEGER,
                date TEXT NOT NULL,
                notes TEXT,
                total_value REAL DEFAULT 0,
                created_by INTEGER,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id),
                FOREIGN KEY (created_by) REFERENCES users(id),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id)
            );

            CREATE TABLE IF NOT EXISTS stock_in_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                stock_in_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL DEFAULT 0,
                total_price REAL DEFAULT 0,
                FOREIGN KEY (stock_in_id) REFERENCES stock_in(id),
                FOREIGN KEY (product_id) REFERENCES products(id)
            );

            -- STOCK OUT / ISSUE REQUESTS
            CREATE TABLE IF NOT EXISTS stock_out (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_number TEXT NOT NULL,
                warehouse_id INTEGER,
                department TEXT,
                employee TEXT,
                employee_id_number TEXT,
                employee_phone TEXT,
                date TEXT NOT NULL,
                reason TEXT,
                status TEXT DEFAULT 'معلق',
                approved_by INTEGER,
                approved_at TEXT,
                issued_by INTEGER,
                issued_at TEXT,
                notes TEXT,
                created_by INTEGER,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (created_by) REFERENCES users(id),
                FOREIGN KEY (approved_by) REFERENCES users(id),
                FOREIGN KEY (issued_by) REFERENCES users(id),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id)
            );

            CREATE TABLE IF NOT EXISTS stock_out_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                stock_out_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity REAL NOT NULL,
                returned_quantity REAL DEFAULT 0,
                FOREIGN KEY (stock_out_id) REFERENCES stock_out(id),
                FOREIGN KEY (product_id) REFERENCES products(id)
            );

            -- STOCK MOVEMENTS (full traceability ledger)
            CREATE TABLE IF NOT EXISTS stock_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                movement_number TEXT UNIQUE NOT NULL,
                movement_type TEXT NOT NULL,
                product_id INTEGER NOT NULL,
                warehouse_id INTEGER,
                quantity REAL NOT NULL,
                date TEXT NOT NULL,
                reference_number TEXT,
                issued_to_name TEXT,
                issued_to_department TEXT,
                issued_to_phone TEXT,
                issued_to_national_id TEXT,
                issued_by_user INTEGER,
                approved_by_user INTEGER,
                received_from_supplier INTEGER,
                reason TEXT,
                current_destination TEXT,
                notes TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (product_id) REFERENCES products(id),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
                FOREIGN KEY (issued_by_user) REFERENCES users(id),
                FOREIGN KEY (approved_by_user) REFERENCES users(id),
                FOREIGN KEY (received_from_supplier) REFERENCES suppliers(id)
            );

            -- RETURNS
            CREATE TABLE IF NOT EXISTS returns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                return_number TEXT UNIQUE NOT NULL,
                stock_out_id INTEGER,
                product_id INTEGER NOT NULL,
                warehouse_id INTEGER,
                quantity REAL NOT NULL,
                returned_by TEXT,
                return_reason TEXT,
                return_date TEXT NOT NULL,
                condition TEXT DEFAULT 'جيد',
                processed_by INTEGER,
                notes TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (stock_out_id) REFERENCES stock_out(id),
                FOREIGN KEY (product_id) REFERENCES products(id),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
                FOREIGN KEY (processed_by) REFERENCES users(id)
            );

            -- TRANSFERS
            CREATE TABLE IF NOT EXISTS transfers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transfer_number TEXT UNIQUE NOT NULL,
                source_warehouse_id INTEGER NOT NULL,
                dest_warehouse_id INTEGER NOT NULL,
                transferred_by INTEGER,
                approved_by INTEGER,
                reason TEXT,
                status TEXT DEFAULT 'معلق',
                date TEXT NOT NULL,
                notes TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (source_warehouse_id) REFERENCES warehouses(id),
                FOREIGN KEY (dest_warehouse_id) REFERENCES warehouses(id),
                FOREIGN KEY (transferred_by) REFERENCES users(id),
                FOREIGN KEY (approved_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS transfer_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transfer_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity REAL NOT NULL,
                FOREIGN KEY (transfer_id) REFERENCES transfers(id),
                FOREIGN KEY (product_id) REFERENCES products(id)
            );

            -- INVENTORY AUDIT
            CREATE TABLE IF NOT EXISTS inventory_audit (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                audit_number TEXT UNIQUE NOT NULL,
                warehouse_id INTEGER,
                product_id INTEGER NOT NULL,
                system_quantity REAL DEFAULT 0,
                physical_quantity REAL DEFAULT 0,
                difference REAL DEFAULT 0,
                value_difference REAL DEFAULT 0,
                audit_date TEXT NOT NULL,
                audited_by INTEGER,
                notes TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
                FOREIGN KEY (product_id) REFERENCES products(id),
                FOREIGN KEY (audited_by) REFERENCES users(id)
            );

            -- INVOICES (preserved)
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT UNIQUE NOT NULL,
                invoice_type TEXT DEFAULT 'purchase',
                supplier_id INTEGER,
                date TEXT NOT NULL,
                subtotal REAL DEFAULT 0,
                tax_rate REAL DEFAULT 5,
                tax_amount REAL DEFAULT 0,
                grand_total REAL DEFAULT 0,
                notes TEXT,
                status TEXT DEFAULT 'pending',
                created_by INTEGER,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS invoice_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_id INTEGER NOT NULL,
                product_id INTEGER,
                description TEXT,
                quantity REAL DEFAULT 1,
                unit_price REAL DEFAULT 0,
                total_price REAL DEFAULT 0,
                FOREIGN KEY (invoice_id) REFERENCES invoices(id),
                FOREIGN KEY (product_id) REFERENCES products(id)
            );

            -- AUDIT LOGS
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT,
                action TEXT NOT NULL,
                action_type TEXT,
                affected_table TEXT,
                affected_id INTEGER,
                details TEXT,
                ip_address TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (user_id) REFERENCES users(id)
            );

            -- NOTIFICATIONS
            CREATE TABLE IF NOT EXISTS notifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                message TEXT,
                type TEXT DEFAULT 'info',
                is_read INTEGER DEFAULT 0,
                created_at TEXT DEFAULT (datetime('now'))
            );

            -- ATTACHMENTS
            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                related_table TEXT NOT NULL,
                related_id INTEGER NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_type TEXT,
                uploaded_by INTEGER,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (uploaded_by) REFERENCES users(id)
            );

            -- OFFERS
            CREATE TABLE IF NOT EXISTS offers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                offer_key TEXT UNIQUE NOT NULL,
                discount_type TEXT NOT NULL, -- percentage or amount
                discount_value REAL NOT NULL,
                apply_to TEXT, -- all, category, or product_id
                apply_id TEXT,
                start_date TEXT,
                end_date TEXT,
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT (datetime('now'))
            );

            -- SETTINGS
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            """)
            conn.commit()

            # ── Default admin
            admin_pw = hashlib.sha256("admin123".encode()).hexdigest()
            conn.execute("""INSERT OR IGNORE INTO users (username, password, role, full_name, email)
                VALUES (?, ?, 'Admin', 'مدير النظام', 'admin@ams.ae')""", ("admin", admin_pw))

            for uname, pw_plain, role, fname in [
                ("manager",  "manager123",  "Warehouse Manager", "مدير المخزن"),
                ("keeper",   "keeper123",   "Store Keeper",      "أمين المخزن"),
                ("auditor",  "auditor123",  "Auditor",           "المدقق"),
                ("viewer",   "viewer123",   "Viewer",            "مستخدم عادي"),
            ]:
                h = hashlib.sha256(pw_plain.encode()).hexdigest()
                conn.execute("INSERT OR IGNORE INTO users (username, password, role, full_name) VALUES (?,?,?,?)",
                             (uname, h, role, fname))

            # ── Default categories
            for c in ["معدات بحرية", "معدات السلامة", "أدوات ومعدات", "كهربائيات",
                      "ميكانيكا", "مواد كيميائية", "قطع غيار", "مستلزمات مكتبية",
                      "تقنية المعلومات", "أجهزة شبكات"]:
                conn.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (c,))

            # ── Default warehouses
            warehouses = [
                ("WH-001", "المخزن الرئيسي",          "مدير المخزن",     "04-0000001", "المبنى الرئيسي"),
                ("WH-002", "مخزن الصيانة",             "مشرف الصيانة",    "04-0000002", "ورشة الصيانة"),
                ("WH-003", "مخزن المعدات البحرية",     "مشرف البحرية",    "04-0000003", "الرصيف البحري"),
                ("WH-004", "مخزن تقنية المعلومات",     "مدير تقنية المعلومات", "04-0000004", "مبنى الإدارة"),
            ]
            for wh in warehouses:
                conn.execute("""INSERT OR IGNORE INTO warehouses
                    (warehouse_code, warehouse_name, warehouse_manager, phone, address)
                    VALUES (?,?,?,?,?)""", wh)

            # ── Default settings
            for k, v in [("theme", "dark"), ("tax_rate", "5"), ("currency", "AED"),
                         ("auto_backup", "true"), ("alarm_enabled", "true"),
                         ("backup_interval", "daily"), ("company_name", COMPANY_NAME)]:
                conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)", (k, v))

            conn.commit()

    def run_migrations(self):
        """Safe schema migrations — adds columns/indexes that don't exist yet."""
        migrations = [
            # Add warehouse_id to stock_in if missing
            "ALTER TABLE stock_in ADD COLUMN warehouse_id INTEGER REFERENCES warehouses(id)",
            # Add new fields to stock_out if missing
            "ALTER TABLE stock_out ADD COLUMN warehouse_id INTEGER REFERENCES warehouses(id)",
            "ALTER TABLE stock_out ADD COLUMN employee_id_number TEXT",
            "ALTER TABLE stock_out ADD COLUMN employee_phone TEXT",
            "ALTER TABLE stock_out ADD COLUMN reason TEXT",
            "ALTER TABLE stock_out ADD COLUMN status TEXT DEFAULT 'معلق'",
            "ALTER TABLE stock_out ADD COLUMN approved_by INTEGER REFERENCES users(id)",
            "ALTER TABLE stock_out ADD COLUMN approved_at TEXT",
            "ALTER TABLE stock_out ADD COLUMN issued_by INTEGER REFERENCES users(id)",
            "ALTER TABLE stock_out ADD COLUMN issued_at TEXT",
            # Add new fields to products if missing
            "ALTER TABLE products ADD COLUMN warehouse_id INTEGER REFERENCES warehouses(id)",
            "ALTER TABLE products ADD COLUMN serial_number TEXT",
            "ALTER TABLE products ADD COLUMN asset_tag TEXT",
            "ALTER TABLE products ADD COLUMN purchase_date TEXT",
            "ALTER TABLE products ADD COLUMN warranty_end TEXT",
            "ALTER TABLE products ADD COLUMN current_owner TEXT",
            "ALTER TABLE products ADD COLUMN current_location TEXT",
            "ALTER TABLE products ADD COLUMN current_status TEXT DEFAULT 'في المخزن'",
            "ALTER TABLE products ADD COLUMN is_asset INTEGER DEFAULT 0",
            # Add phone/dept to users
            "ALTER TABLE users ADD COLUMN phone TEXT",
            "ALTER TABLE users ADD COLUMN department TEXT",
            "ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT '[]'",
        ]
        with self.get_conn() as conn:
            for sql in migrations:
                try:
                    conn.execute(sql)
                except Exception:
                    pass  # Column already exists or other benign error
            conn.commit()

    # ─── Settings ───────────────────────────────────────────────────────────
    def get_setting(self, key, default=""):
        row = self.fetchone("SELECT value FROM settings WHERE key=?", (key,))
        return row[0] if row else default

    def set_setting(self, key, value):
        self.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?,?)", (key, str(value)))

    # ─── Audit Log ──────────────────────────────────────────────────────────
    def log_action(self, user_id, username, action, action_type="", table="", record_id=None, details=""):
        self.execute("""INSERT INTO audit_logs
            (user_id, username, action, action_type, affected_table, affected_id, details)
            VALUES (?,?,?,?,?,?,?)""",
            (user_id, username, action, action_type, table, record_id, details))

    # ─── Number generators ──────────────────────────────────────────────────
    def generate_product_id(self):
        row = self.fetchone("SELECT MAX(CAST(SUBSTR(product_id,4) AS INTEGER)) as mx FROM products WHERE product_id LIKE 'AMS%'")
        return f"AMS{(row['mx'] or 0) + 1:05d}"

    def generate_invoice_number(self):
        now = datetime.now()
        prefix = f"INV-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM invoices WHERE invoice_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_stock_in_number(self):
        now = datetime.now()
        prefix = f"SI-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM stock_in WHERE invoice_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_stock_out_number(self):
        now = datetime.now()
        prefix = f"SO-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM stock_out WHERE request_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_movement_number(self):
        now = datetime.now()
        prefix = f"MV-{now.year}{now.month:02d}{now.day:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM stock_movements WHERE movement_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_transfer_number(self):
        now = datetime.now()
        prefix = f"TR-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM transfers WHERE transfer_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_return_number(self):
        now = datetime.now()
        prefix = f"RT-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM returns WHERE return_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    def generate_audit_number(self):
        now = datetime.now()
        prefix = f"AU-{now.year}{now.month:02d}-"
        row = self.fetchone(f"SELECT COUNT(*)+1 as n FROM inventory_audit WHERE audit_number LIKE '{prefix}%'")
        return f"{prefix}{row['n']:04d}"

    # ─── Dashboard ──────────────────────────────────────────────────────────
    def get_dashboard_stats(self):
        stats = {}
        r = self.fetchone("SELECT COUNT(*) as c, COALESCE(SUM(quantity*selling_price),0) as v FROM products WHERE is_active=1")
        stats['total_products'] = r['c']
        stats['inventory_value'] = r['v']
        r2 = self.fetchone("SELECT COUNT(*) as c FROM products WHERE is_active=1 AND quantity <= alert_quantity AND quantity > min_quantity AND alarms_enabled=1")
        stats['low_stock'] = r2['c']
        r3 = self.fetchone("SELECT COUNT(*) as c FROM products WHERE is_active=1 AND quantity <= min_quantity AND alarms_enabled=1")
        stats['critical_stock'] = r3['c']
        r4 = self.fetchone("SELECT COUNT(*) as c FROM suppliers WHERE is_active=1")
        stats['suppliers'] = r4['c']
        r5 = self.fetchone("SELECT COUNT(*) as c FROM stock_in WHERE date >= date('now', '-30 days')")
        stats['incoming'] = r5['c']
        r6 = self.fetchone("SELECT COUNT(*) as c FROM stock_out WHERE date >= date('now', '-30 days')")
        stats['outgoing'] = r6['c']
        r7 = self.fetchone("SELECT COUNT(*) as c FROM stock_out WHERE status='معلق'")
        stats['pending_requests'] = r7['c']
        r8 = self.fetchone("SELECT COUNT(*) as c FROM warehouses WHERE is_active=1")
        stats['warehouses'] = r8['c']
        # Today
        today = date.today().isoformat()
        r9 = self.fetchone("SELECT COALESCE(SUM(sii.quantity),0) as q FROM stock_in si JOIN stock_in_items sii ON si.id=sii.stock_in_id WHERE si.date=?", (today,))
        stats['stock_in_today'] = r9['q']
        r10 = self.fetchone("SELECT COALESCE(SUM(soi.quantity),0) as q FROM stock_out so JOIN stock_out_items soi ON so.id=soi.stock_out_id WHERE so.date=?", (today,))
        stats['stock_out_today'] = r10['q']
        return stats

    def get_stock_movement(self, days=30):
        in_data = self.fetchall(f"""
            SELECT date, SUM(sii.quantity) as qty
            FROM stock_in si JOIN stock_in_items sii ON si.id=sii.stock_in_id
            WHERE si.date >= date('now', '-{days} days')
            GROUP BY si.date ORDER BY si.date
        """)
        out_data = self.fetchall(f"""
            SELECT date, SUM(soi.quantity) as qty
            FROM stock_out so JOIN stock_out_items soi ON so.id=soi.stock_out_id
            WHERE so.date >= date('now', '-{days} days')
            GROUP BY so.date ORDER BY so.date
        """)
        return in_data, out_data

    def check_low_stock(self):
        return self.fetchall("""
            SELECT p.*, s.name as supplier_name, w.warehouse_name
            FROM products p
            LEFT JOIN suppliers s ON p.supplier_id = s.id
            LEFT JOIN warehouses w ON p.warehouse_id = w.id
            WHERE p.is_active=1 AND p.alarms_enabled=1
            AND p.quantity <= p.alert_quantity
            ORDER BY p.quantity ASC
        """)

    def check_critical_stock(self):
        return self.fetchall("""
            SELECT p.*, s.name as supplier_name, w.warehouse_name
            FROM products p
            LEFT JOIN suppliers s ON p.supplier_id = s.id
            LEFT JOIN warehouses w ON p.warehouse_id = w.id
            WHERE p.is_active=1 AND p.alarms_enabled=1
            AND p.quantity <= p.min_quantity
            ORDER BY p.quantity ASC
        """)

    def get_item_trace(self, product_id):
        """Full lifecycle trace for one product."""
        return self.fetchall("""
            SELECT sm.*, u1.full_name as issued_by_name,
                   u2.full_name as approved_by_name,
                   w.warehouse_name, p.name as product_name,
                   s.name as supplier_name
            FROM stock_movements sm
            LEFT JOIN users u1 ON sm.issued_by_user = u1.id
            LEFT JOIN users u2 ON sm.approved_by_user = u2.id
            LEFT JOIN warehouses w ON sm.warehouse_id = w.id
            LEFT JOIN products p ON sm.product_id = p.id
            LEFT JOIN suppliers s ON sm.received_from_supplier = s.id
            WHERE sm.product_id = ?
            ORDER BY sm.created_at DESC
        """, (product_id,))

    def get_most_used_products(self, limit=10):
        """Returns products with highest issuance frequency."""
        return self.fetchall("""
            SELECT p.name, COUNT(soi.id) as usage_frequency, SUM(soi.quantity) as total_issued
            FROM products p
            JOIN stock_out_items soi ON p.id = soi.product_id
            GROUP BY p.id
            ORDER BY usage_frequency DESC, total_issued DESC
            LIMIT ?
        """, (limit,))

    def get_product_purchase_insights(self, product_id):
        """Returns the most recent purchase price and supplier for a product."""
        return self.fetchone("""
            SELECT sii.unit_price, s.name as supplier_name, si.date, s.id as supplier_id
            FROM stock_in_items sii
            JOIN stock_in si ON sii.stock_in_id = si.id
            JOIN suppliers s ON si.supplier_id = s.id
            WHERE sii.product_id = ?
            ORDER BY si.date DESC, si.created_at DESC
            LIMIT 1
        """, (product_id,))


# Singleton DB
db = DatabaseManager()

# ─────────────────────────────────────────────────────────────────────────────
# CURRENT LOGGED USER (global)
# ─────────────────────────────────────────────────────────────────────────────
CURRENT_USER = {"id": None, "username": "system", "full_name": "System", "role": "Admin"}


def set_current_user(user_dict):
    global CURRENT_USER
    CURRENT_USER = user_dict


def log(action, table="", record_id=None, details=""):
    db.log_action(CURRENT_USER.get("id"), CURRENT_USER.get("username", ""),
                  action, "", table, record_id, details)


# ─────────────────────────────────────────────────────────────────────────────
# PERMISSION HELPER
# ─────────────────────────────────────────────────────────────────────────────
ROLE_PERMISSIONS = {
    "Admin":             {"all"},
    "Warehouse Manager": {"view", "add", "edit", "approve", "transfer", "report", "audit"},
    "Store Keeper":      {"view", "add", "stock_in", "stock_out"},
    "Auditor":           {"view", "audit", "report"},
    "Viewer":            {"view"},
}


def has_perm(perm: str) -> bool:
    # Role based
    role_perms = ROLE_PERMISSIONS.get(CURRENT_USER.get("role", "Viewer"), set())
    if "all" in role_perms or perm in role_perms:
        return True

    # User based (granular)
    user_perms_raw = CURRENT_USER.get("permissions") or "[]"
    try:
        user_perms = json.loads(user_perms_raw.replace("'", '"')) if isinstance(user_perms_raw, str) else user_perms_raw
        if isinstance(user_perms, list) and (perm in user_perms or "all" in user_perms):
            return True
        if isinstance(user_perms, str) and (perm in user_perms or "all" in user_perms):
            return True
    except:
        if perm in str(user_perms_raw): return True

    return False


# ─────────────────────────────────────────────────────────────────────────────
# SPLASH SCREEN
# ─────────────────────────────────────────────────────────────────────────────
class SplashScreen(QSplashScreen):
    def __init__(self):
        pixmap = self._create_splash()
        super().__init__(pixmap, Qt.WindowStaysOnTopHint)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self._progress = 0
        self._base_pixmap = QPixmap(pixmap)

    def _create_splash(self):
        w, h = 720, 440
        pixmap = QPixmap(w, h)
        pixmap.fill(QColor("#0A0E1A"))
        p = QPainter(pixmap)
        p.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, w, h)
        grad.setColorAt(0, QColor("#0A0E1A"))
        grad.setColorAt(0.5, QColor("#0D1226"))
        grad.setColorAt(1, QColor("#0A0E1A"))
        p.fillRect(0, 0, w, h, grad)
        p.setPen(QPen(QColor("#0EA5E920"), 1))
        p.drawEllipse(QPoint(360, 220), 310, 310)
        p.setPen(QPen(QColor("#0EA5E910"), 1))
        p.drawEllipse(QPoint(360, 220), 210, 210)
        if LOGO_PATH.exists():
            logo = QPixmap(str(LOGO_PATH)).scaled(90, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            p.drawPixmap((w - 90) // 2, 55, logo)
        p.setPen(QColor("#F1F5F9"))
        p.setFont(QFont("Segoe UI", 15, QFont.Bold))
        p.drawText(QRect(0, 160, w, 30), Qt.AlignCenter, COMPANY_NAME)
        p.setFont(QFont("Cairo", 12))
        p.setPen(QColor(COLORS['accent']))
        p.drawText(QRect(0, 192, w, 26), Qt.AlignCenter, "نظام إدارة المخازن المؤسسي")
        p.setFont(QFont("Cairo", 10))
        p.setPen(QColor(COLORS['text_muted']))
        p.drawText(QRect(0, 222, w, 22), Qt.AlignCenter, "نظام تخطيط موارد المؤسسات الذكي — AMS")
        p.setPen(QPen(QColor("#1E293B"), 1))
        p.drawLine(80, 280, 640, 280)
        p.end()
        return pixmap

    def set_progress(self, pct, msg):
        w, h = 720, 440
        overlay = QPixmap(self._base_pixmap.size())
        overlay.fill(Qt.transparent)
        p = QPainter(overlay)
        p.setRenderHint(QPainter.Antialiasing)
        bx, by, bw, bh2 = 80, 300, 560, 9
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#1E293B"))
        p.drawRoundedRect(bx, by, bw, bh2, 4, 4)
        fw = int(bw * pct / 100)
        if fw > 0:
            g = QLinearGradient(bx, 0, bx + fw, 0)
            g.setColorAt(0, QColor("#0369A1"))
            g.setColorAt(1, QColor("#38BDF8"))
            p.setBrush(g)
            p.drawRoundedRect(bx, by, fw, bh2, 4, 4)
        p.setPen(QColor("#94A3B8"))
        p.setFont(QFont("Tahoma", 9))
        p.drawText(QRect(bx, by + 18, bw, 20), Qt.AlignCenter, f"{msg}  {pct}%")
        p.end()
        combined = QPixmap(self._base_pixmap.size())
        combined.fill(Qt.transparent)
        cp = QPainter(combined)
        cp.drawPixmap(0, 0, self._base_pixmap)
        cp.drawPixmap(0, 0, overlay)
        cp.end()
        self.setPixmap(combined)
        QApplication.processEvents()


# ─────────────────────────────────────────────────────────────────────────────
# LOGIN WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class LoginWindow(QDialog):
    login_successful = Signal(dict)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{COMPANY_NAME} — تسجيل الدخول")
        self.setFixedSize(480, 640)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self._drag_pos = None
        self.setup_ui()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.LeftButton and self._drag_pos:
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    def setup_ui(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        # Title bar
        tb = QFrame()
        tb.setFixedHeight(44)
        tb.setStyleSheet(f"background: {COLORS['bg_nav']}; border-bottom: 1px solid {COLORS['border2']};")
        tb_lay = QHBoxLayout(tb)
        tb_lay.setContentsMargins(16, 0, 12, 0)
        tb_lay.addWidget(QLabel("⚓"))
        tl = QLabel("نظام إدارة المخازن — AMS")
        tl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 12px; font-weight: 600;")
        tb_lay.addWidget(tl)
        tb_lay.addStretch()
        cb = QPushButton("✕")
        cb.setFixedSize(28, 28)
        cb.setStyleSheet(f"background: transparent; color: {COLORS['text_secondary']}; border: none; font-size: 14px;")
        cb.clicked.connect(self.reject)
        tb_lay.addWidget(cb)
        main.addWidget(tb)

        body = QFrame()
        body.setStyleSheet(f"background: {COLORS['bg_dark']};")
        blay = QVBoxLayout(body)
        blay.setContentsMargins(50, 28, 50, 28)
        blay.setSpacing(0)

        if LOGO_PATH.exists():
            ll = QLabel()
            ll.setPixmap(QPixmap(str(LOGO_PATH)).scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            ll.setAlignment(Qt.AlignCenter)
            blay.addWidget(ll)
            blay.addSpacing(12)

        co = QLabel(COMPANY_NAME)
        co.setAlignment(Qt.AlignCenter)
        co.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 12px; font-weight: 700;")
        co.setWordWrap(True)
        blay.addWidget(co)

        sub = QLabel("نظام إدارة المخازن المؤسسي")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet(f"color: {COLORS['accent']}; font-size: 11px; font-weight: 600; margin-bottom: 4px;")
        blay.addWidget(sub)
        blay.addSpacing(24)

        card = QFrame()
        card.setObjectName("card")
        card.setStyleSheet(f"QFrame#card {{ background: {COLORS['bg_card']}; border: 1px solid {COLORS['border2']}; border-radius: 16px; }}")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(30, 26, 30, 26)
        cl.setSpacing(14)

        signin = QLabel("تسجيل الدخول")
        signin.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 18px; font-weight: 700;")
        signin.setAlignment(Qt.AlignCenter)
        cl.addWidget(signin)

        for lbl_text, attr, placeholder, pw in [
            ("اسم المستخدم", "username_edit", "أدخل اسم المستخدم", False),
            ("كلمة المرور",  "password_edit", "أدخل كلمة المرور",  True),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; font-weight: 600;")
            cl.addWidget(lbl)
            if pw:
                row = QHBoxLayout()
                row.setSpacing(8)
                edit = QLineEdit()
                edit.setEchoMode(QLineEdit.Password)
                edit.setPlaceholderText(placeholder)
                edit.setFixedHeight(42)
                row.addWidget(edit)
                eye = QPushButton("👁")
                eye.setFixedSize(42, 42)
                eye.setObjectName("btn_flat")
                eye.setCheckable(True)
                eye.clicked.connect(lambda c, e=edit: e.setEchoMode(QLineEdit.Normal if c else QLineEdit.Password))
                row.addWidget(eye)
                cl.addLayout(row)
            else:
                edit = QLineEdit()
                edit.setPlaceholderText(placeholder)
                edit.setFixedHeight(42)
                cl.addWidget(edit)
            setattr(self, attr, edit)

        self.login_btn = QPushButton("🔐  دخول النظام")
        self.login_btn.setFixedHeight(46)
        self.login_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {COLORS['accent']}, stop:1 {COLORS['accent2']});
                color: white; border: none; border-radius: 10px;
                font-size: 15px; font-weight: 700;
            }}
            QPushButton:hover {{ background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {COLORS['accent2']}, stop:1 {COLORS['accent']}); }}
        """)
        self.login_btn.clicked.connect(self.do_login)
        cl.addWidget(self.login_btn)

        self.error_lbl = QLabel("")
        self.error_lbl.setAlignment(Qt.AlignCenter)
        self.error_lbl.setStyleSheet(f"color: {COLORS['accent_red']}; font-size: 12px;")
        self.error_lbl.setVisible(False)
        cl.addWidget(self.error_lbl)

        blay.addWidget(card)
        blay.addStretch()

        footer = QLabel(f"© {datetime.now().year} {COMPANY_NAME}")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 10px;")
        blay.addWidget(footer)

        main.addWidget(body)
        self.password_edit.returnPressed.connect(self.do_login)
        self.username_edit.returnPressed.connect(lambda: self.password_edit.setFocus())

        remembered = db.get_setting("remembered_user")
        if remembered:
            self.username_edit.setText(remembered)
            self.password_edit.setFocus()

    def do_login(self):
        username = self.username_edit.text().strip()
        password = self.password_edit.text()
        if not username or not password:
            self._show_error("يرجى إدخال اسم المستخدم وكلمة المرور")
            return
        pw_hash = hashlib.sha256(password.encode()).hexdigest()
        user = db.fetchone("SELECT * FROM users WHERE username=? AND password=? AND is_active=1",
                           (username, pw_hash))
        if user:
            db.execute("UPDATE users SET last_login=? WHERE id=?",
                       (datetime.now().isoformat(), user['id']))
            user_dict = dict(user)
            set_current_user(user_dict)
            log(f"تسجيل دخول: {username}", "users", user['id'])
            self.login_successful.emit(user_dict)
            self.accept()
        else:
            self._show_error("اسم المستخدم أو كلمة المرور غير صحيحة")
            self.password_edit.clear()

    def _show_error(self, msg):
        self.error_lbl.setText(f"⚠ {msg}")
        self.error_lbl.setVisible(True)
        QTimer.singleShot(4000, lambda: self.error_lbl.setVisible(False))


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR BUTTON (Arabic/RTL)
# ─────────────────────────────────────────────────────────────────────────────
class SidebarButton(QPushButton):
    def __init__(self, icon, label, parent=None):
        super().__init__(parent)
        self.setText(f"  {icon}  {label}")
        self.setCheckable(True)
        self.setFixedHeight(46)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(self._style(False))
        self.toggled.connect(lambda c: self.setStyleSheet(self._style(c)))

    def _style(self, active):
        if active:
            return f"""QPushButton {{
                background: qlineargradient(x1:1,y1:0,x2:0,y2:0, stop:0 {COLORS['accent']}22, stop:1 transparent);
                color: {COLORS['accent']}; border: none; border-left: 3px solid {COLORS['accent']};
                border-radius: 0; text-align: right; padding-right: 16px; font-weight: 700; font-size: 13px;
            }}"""
        return f"""QPushButton {{
                background: transparent; color: {COLORS['text_secondary']}; border: none;
                border-left: 3px solid transparent; border-radius: 0;
                text-align: right; padding-right: 16px; font-size: 13px;
            }}
            QPushButton:hover {{ background: {COLORS['bg_card2']}; color: {COLORS['text_primary']}; }}"""


# ─────────────────────────────────────────────────────────────────────────────
# STAT CARD
# ─────────────────────────────────────────────────────────────────────────────
class StatCard(QFrame):
    def __init__(self, title, value, icon, color, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setFixedHeight(110)
        self._setup(title, value, icon, color)

    def _setup(self, title, value, icon, color):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 14, 18, 14)
        info = QVBoxLayout()
        info.setSpacing(4)
        t = QLabel(title)
        t.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 11px; font-weight: 600;")
        self.val_lbl = QLabel(str(value))
        self.val_lbl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 24px; font-weight: 800;")
        info.addWidget(t)
        info.addWidget(self.val_lbl)
        icon_frame = QFrame()
        icon_frame.setFixedSize(52, 52)
        icon_frame.setStyleSheet(f"background: {color}22; border-radius: 13px;")
        il = QVBoxLayout(icon_frame)
        il.setContentsMargins(0, 0, 0, 0)
        il2 = QLabel(icon)
        il2.setAlignment(Qt.AlignCenter)
        il2.setStyleSheet("font-size: 20px; background: transparent;")
        il.addWidget(il2)
        lay.addLayout(info)
        lay.addStretch()
        lay.addWidget(icon_frame)
        self.setStyleSheet(f"""QFrame#card {{
            background: {COLORS['bg_card']}; border: 1px solid {COLORS['border2']};
            border-radius: 14px; border-left: 4px solid {color};
        }} QFrame#card:hover {{ background: {COLORS['bg_card2']}; }}""")

    def update_value(self, v):
        self.val_lbl.setText(str(v))


# ─────────────────────────────────────────────────────────────────────────────
# BAR CHART WIDGET
# ─────────────────────────────────────────────────────────────────────────────
class BarChartWidget(QWidget):
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.title = title
        self.data_in = []
        self.data_out = []
        self.labels = []
        self.setMinimumHeight(190)

    def set_data(self, labels, data_in, data_out=None):
        self.labels = labels
        self.data_in = data_in
        self.data_out = data_out or []
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        p.fillRect(0, 0, w, h, QColor(COLORS['bg_card']))
        if not self.labels:
            p.setPen(QColor(COLORS['text_muted']))
            p.drawText(QRect(0, 0, w, h), Qt.AlignCenter, "لا توجد بيانات")
            p.end()
            return
        title_h = 28
        pad_l, pad_r, pad_b = 50, 20, 36
        cx = pad_l
        cy = title_h + 8
        cw = w - pad_l - pad_r
        ch = h - title_h - 8 - pad_b
        p.setPen(QColor(COLORS['text_primary']))
        p.setFont(QFont("Tahoma", 10, QFont.Bold))
        p.drawText(QRect(0, 0, w, title_h), Qt.AlignCenter, self.title)
        all_vals = self.data_in + self.data_out
        max_val = max(all_vals) if all_vals else 1
        n = len(self.labels)
        group_w = cw / max(n, 1)
        bar_w = group_w * 0.35 if self.data_out else group_w * 0.5
        p.setPen(QPen(QColor(COLORS['border']), 1, Qt.DotLine))
        for i in range(5):
            y = cy + ch - (i / 4) * ch
            p.drawLine(int(cx), int(y), int(cx + cw), int(y))
            p.setPen(QColor(COLORS['text_muted']))
            p.setFont(QFont("Tahoma", 8))
            p.drawText(QRect(0, int(y) - 8, pad_l - 4, 16), Qt.AlignRight | Qt.AlignVCenter, str(int(max_val * i / 4)))
            p.setPen(QPen(QColor(COLORS['border']), 1, Qt.DotLine))
        for i, label in enumerate(self.labels):
            lx = cx + i * group_w + group_w / 2
            if self.data_out:
                if i < len(self.data_in) and self.data_in[i] > 0:
                    bh2 = (self.data_in[i] / max_val) * ch
                    bx2 = lx - bar_w
                    by2 = cy + ch - bh2
                    g = QLinearGradient(0, by2, 0, by2 + bh2)
                    g.setColorAt(0, QColor(COLORS['accent2']))
                    g.setColorAt(1, QColor(COLORS['accent']))
                    p.setBrush(g)
                    p.setPen(Qt.NoPen)
                    p.drawRoundedRect(QRectF(bx2, by2, bar_w, bh2), 3, 3)
                if i < len(self.data_out) and self.data_out[i] > 0:
                    bh2 = (self.data_out[i] / max_val) * ch
                    bx2 = lx
                    by2 = cy + ch - bh2
                    g2 = QLinearGradient(0, by2, 0, by2 + bh2)
                    g2.setColorAt(0, QColor(COLORS['accent_red']))
                    g2.setColorAt(1, QColor("#B91C1C"))
                    p.setBrush(g2)
                    p.setPen(Qt.NoPen)
                    p.drawRoundedRect(QRectF(bx2, by2, bar_w, bh2), 3, 3)
            else:
                if i < len(self.data_in) and self.data_in[i] > 0:
                    bh2 = (self.data_in[i] / max_val) * ch
                    bx2 = lx - bar_w / 2
                    by2 = cy + ch - bh2
                    g = QLinearGradient(0, by2, 0, by2 + bh2)
                    g.setColorAt(0, QColor(COLORS['accent2']))
                    g.setColorAt(1, QColor(COLORS['accent']))
                    p.setBrush(g)
                    p.setPen(Qt.NoPen)
                    p.drawRoundedRect(QRectF(bx2, by2, bar_w, bh2), 3, 3)
            p.setPen(QColor(COLORS['text_muted']))
            p.setFont(QFont("Tahoma", 7))
            p.drawText(QRect(int(lx - group_w / 2), int(cy + ch + 4), int(group_w), 20), Qt.AlignCenter, str(label))
        p.end()


# ─────────────────────────────────────────────────────────────────────────────
# NOTIFICATION BADGE
# ─────────────────────────────────────────────────────────────────────────────
class NotificationBadge(QPushButton):
    def __init__(self, parent=None):
        super().__init__("🔔", parent)
        self._count = 0
        self.setFixedSize(36, 36)
        self.setObjectName("btn_flat")
        self.setCursor(Qt.PointingHandCursor)
        self.clicked.connect(self.show_notifications)

    def set_count(self, n):
        self._count = n
        self.setText(f"🔔 {n}" if n > 0 else "🔔")
        self.setStyleSheet(f"background: {'#450A0A' if n > 0 else 'transparent'}; color: {'#EF4444' if n > 0 else COLORS['text_secondary']}; border: 1px solid {COLORS['border2']}; border-radius: 8px; font-size: 11px;")

    def show_notifications(self):
        notifs = db.fetchall("SELECT * FROM notifications ORDER BY created_at DESC LIMIT 50")
        dlg = QDialog(self.window())
        dlg.setWindowTitle("الإشعارات")
        dlg.setMinimumSize(420, 500)
        lay = QVBoxLayout(dlg)
        hdr = QLabel(f"🔔  الإشعارات  ({len(notifs)})")
        hdr.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']}; padding: 8px;")
        lay.addWidget(hdr)
        if not notifs:
            lay.addWidget(QLabel("لا توجد إشعارات"))
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            container = QWidget()
            vl = QVBoxLayout(container)
            for n in notifs:
                nf = QFrame()
                colors_map = {"warning": "#451A03", "critical": "#450A0A", "info": COLORS['bg_card2'], "success": "#052E16"}
                nf.setStyleSheet(f"background: {colors_map.get(n['type'], COLORS['bg_card2'])}; border-radius: 8px; padding: 4px;")
                nl = QVBoxLayout(nf)
                nl.setSpacing(2)
                nl.addWidget(QLabel(f"<b>{n['title']}</b>"))
                nl.addWidget(QLabel(n['message'] or ""))
                tl = QLabel(n['created_at'])
                tl.setStyleSheet(f"font-size: 10px; color: {COLORS['text_muted']};")
                nl.addWidget(tl)
                vl.addWidget(nf)
            vl.addStretch()
            scroll.setWidget(container)
            lay.addWidget(scroll)
        clear_btn = QPushButton("مسح الكل")
        clear_btn.clicked.connect(lambda: (db.execute("UPDATE notifications SET is_read=1"),
                                            self.set_count(0), dlg.accept()))
        lay.addWidget(clear_btn)
        db.execute("UPDATE notifications SET is_read=1")
        self.set_count(0)
        dlg.exec()


# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD PAGE
# ─────────────────────────────────────────────────────────────────────────────
class DashboardPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(24, 20, 24, 20)
        main.setSpacing(18)

        # Header
        hdr = QHBoxLayout()
        title = QLabel("📊  لوحة التحكم")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        dl = QLabel(datetime.now().strftime("%A, %d %B %Y"))
        dl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 12px;")
        hdr.addWidget(title)
        hdr.addStretch()
        hdr.addWidget(dl)
        main.addLayout(hdr)

        # Stat cards — row 1
        grid = QGridLayout()
        grid.setSpacing(14)
        for i in range(4):
            grid.setColumnStretch(i, 1)

        self.cards = {}
        card_defs = [
            ("total_products",  "إجمالي المنتجات",    "0",      "📦", COLORS['accent']),
            ("warehouses",      "المخازن النشطة",      "0",      "🏭", COLORS['accent_purple']),
            ("suppliers",       "الموردين",             "0",      "🏢", COLORS['accent_green']),
            ("inventory_value", "قيمة المخزون",        "AED 0",  "💰", COLORS['accent_green']),
            ("stock_in_today",  "وارد اليوم",           "0",      "📥", COLORS['accent']),
            ("stock_out_today", "صادر اليوم",           "0",      "📤", COLORS['accent_red']),
            ("pending_requests","طلبات معلقة",          "0",      "⏳", COLORS['accent_orange']),
            ("critical_stock",  "مخزون حرج",           "0",      "🚨", COLORS['accent_red']),
        ]
        for i, (key, title, val, icon, color) in enumerate(card_defs):
            card = StatCard(title, val, icon, color)
            self.cards[key] = card
            grid.addWidget(card, i // 4, i % 4)
        main.addLayout(grid)

        # Charts row
        cr = QHBoxLayout()
        cr.setSpacing(14)

        chart_card = QFrame()
        chart_card.setObjectName("card")
        chart_card.setMinimumHeight(230)
        ccl = QVBoxLayout(chart_card)
        ccl.setContentsMargins(14, 14, 14, 14)
        ct = QLabel("📈  حركة المخزون (آخر 30 يوم)")
        ct.setStyleSheet(f"font-weight: 700; font-size: 13px; color: {COLORS['text_primary']};")
        ccl.addWidget(ct)
        self.movement_chart = BarChartWidget()
        ccl.addWidget(self.movement_chart)
        leg = QHBoxLayout()
        for lbl2, col in [("وارد", COLORS['accent']), ("صادر", COLORS['accent_red'])]:
            d = QLabel("●")
            d.setStyleSheet(f"color: {col}; font-size: 14px;")
            lb = QLabel(lbl2)
            lb.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 11px;")
            leg.addWidget(d); leg.addWidget(lb)
        leg.addStretch()
        ccl.addLayout(leg)
        cr.addWidget(chart_card, 2)

        # Intelligence Card
        intel_card = QFrame()
        intel_card.setObjectName("card")
        icl = QVBoxLayout(intel_card)
        icl.setContentsMargins(14, 14, 14, 14)
        it = QLabel("🧠  التحليلات الذكية")
        it.setStyleSheet(f"font-weight: 700; font-size: 13px; color: {COLORS['accent']};")
        icl.addWidget(it)
        self.intel_lbl = QLabel("جاري تحليل البيانات...")
        self.intel_lbl.setWordWrap(True)
        self.intel_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; line-height: 1.5;")
        icl.addWidget(self.intel_lbl)
        icl.addStretch()

        tt = QLabel("🏆  أكثر المنتجات استخداماً")
        tt.setStyleSheet(f"font-weight: 700; font-size: 12px; color: {COLORS['text_primary']}; margin-top: 10px;")
        icl.addWidget(tt)
        self.top_table = QTableWidget()
        self.top_table.setColumnCount(2)
        self.top_table.setHorizontalHeaderLabels(["المنتج", "التكرار"])
        self.top_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.top_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.top_table.setFixedHeight(150)
        self.top_table.verticalHeader().setVisible(False)
        self.top_table.setShowGrid(False)
        icl.addWidget(self.top_table)
        cr.addWidget(intel_card, 1)
        main.addLayout(cr)

        # Recent movements
        rm_card = QFrame()
        rm_card.setObjectName("card")
        rml = QVBoxLayout(rm_card)
        rml.setContentsMargins(14, 14, 14, 14)
        rmh = QHBoxLayout()
        rmt = QLabel("🔄  آخر حركات المخزون")
        rmt.setStyleSheet(f"font-weight: 700; font-size: 13px; color: {COLORS['text_primary']};")
        rmh.addWidget(rmt); rmh.addStretch()
        rml.addLayout(rmh)
        self.recent_table = QTableWidget()
        self.recent_table.setColumnCount(6)
        self.recent_table.setHorizontalHeaderLabels(["رقم الحركة", "النوع", "المنتج", "الكمية", "الوجهة", "التاريخ"])
        self.recent_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.recent_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.recent_table.setAlternatingRowColors(True)
        self.recent_table.verticalHeader().setVisible(False)
        self.recent_table.setMaximumHeight(200)
        rml.addWidget(self.recent_table)
        main.addWidget(rm_card)

        self.refresh_data()

    def refresh_data(self):
        stats = db.get_dashboard_stats()
        self.cards['total_products'].update_value(stats['total_products'])
        self.cards['warehouses'].update_value(stats['warehouses'])
        self.cards['suppliers'].update_value(stats['suppliers'])
        self.cards['critical_stock'].update_value(stats['critical_stock'])
        self.cards['pending_requests'].update_value(stats['pending_requests'])
        self.cards['stock_in_today'].update_value(f"{stats['stock_in_today']:.0f}")
        self.cards['stock_out_today'].update_value(f"{stats['stock_out_today']:.0f}")
        currency = db.get_setting("currency", "AED")
        val = stats['inventory_value']
        val_str = f"{currency} {val/1_000_000:.1f}M" if val >= 1_000_000 else (f"{currency} {val/1000:.1f}K" if val >= 1000 else f"{currency} {val:.0f}")
        self.cards['inventory_value'].update_value(val_str)

        in_data, out_data = db.get_stock_movement(30)
        all_dates = set()
        in_dict, out_dict = {}, {}
        for r in in_data:
            in_dict[r['date']] = r['qty']; all_dates.add(r['date'])
        for r in out_data:
            out_dict[r['date']] = r['qty']; all_dates.add(r['date'])
        sorted_dates = sorted(all_dates)[-14:]
        labels = [d[-5:] for d in sorted_dates]
        vals_in = [in_dict.get(d, 0) for d in sorted_dates]
        vals_out = [out_dict.get(d, 0) for d in sorted_dates]
        if not sorted_dates:
            labels = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو"]
            vals_in = vals_out = [0] * 7
        self.movement_chart.set_data(labels, vals_in, vals_out)

        # Intelligence Data
        most_used = db.get_most_used_products(5)
        self.top_table.setRowCount(len(most_used))
        currency = db.get_setting("currency", "AED")
        for i, row in enumerate(most_used):
            self.top_table.setItem(i, 0, QTableWidgetItem(row['name']))
            self.top_table.setItem(i, 1, QTableWidgetItem(f"{row['usage_frequency']} طلبات"))

        # Smart Insights Text
        if most_used:
            best_prod = most_used[0]['name']
            prod_row = db.fetchone("SELECT id, quantity FROM products WHERE name=?", (best_prod,))
            insights = db.get_product_purchase_insights(prod_row['id'])

            # Smart Forecast (Logic: Total Issued / 30 days = daily rate)
            total_issued = most_used[0]['total_issued']
            daily_rate = total_issued / 30
            days_left = (prod_row['quantity'] / daily_rate) if daily_rate > 0 else 999

            txt = f"المنتج الأكثر طلباً هو <b>{best_prod}</b>.<br>"
            if insights:
                txt += f"يُشترى من <b>{insights['supplier_name']}</b> بسعر <b>{currency} {insights['unit_price']:.2f}</b>.<br>"

            if days_left < 7:
                txt += f"⚠️ المخزون قد ينفذ خلال <b>{days_left:.1f} أيام</b> بناءً على معدل الاستهلاك."
            elif daily_rate > 0:
                txt += f"✅ المخزون يكفي لحوالي <b>{days_left:.0f} يوم</b>."

            self.intel_lbl.setText(txt)
        else:
            self.intel_lbl.setText("ابدأ بتسجيل عمليات الصرف للحصول على تحليلات ذكية وتوقعات المخزون.")

        recent_mvs = db.fetchall("""
            SELECT sm.movement_number, sm.movement_type, p.name as product_name,
                   sm.quantity, sm.current_destination, sm.created_at
            FROM stock_movements sm
            LEFT JOIN products p ON sm.product_id = p.id
            ORDER BY sm.created_at DESC LIMIT 15
        """)
        self.recent_table.setRowCount(len(recent_mvs))
        for i, row in enumerate(recent_mvs):
            self.recent_table.setItem(i, 0, QTableWidgetItem(row['movement_number'] or ""))
            type_item = QTableWidgetItem(row['movement_type'] or "")
            color_map = {"وارد": COLORS['accent_green'], "صادر": COLORS['accent_red'], "تحويل": COLORS['accent_orange']}
            type_item.setForeground(QColor(color_map.get(row['movement_type'], COLORS['text_primary'])))
            self.recent_table.setItem(i, 1, type_item)
            self.recent_table.setItem(i, 2, QTableWidgetItem(row['product_name'] or ""))
            self.recent_table.setItem(i, 3, QTableWidgetItem(f"{row['quantity']:.0f}"))
            self.recent_table.setItem(i, 4, QTableWidgetItem(row['current_destination'] or ""))
            self.recent_table.setItem(i, 5, QTableWidgetItem((row['created_at'] or "")[:16]))


# ─────────────────────────────────────────────────────────────────────────────
# WAREHOUSES PAGE
# ─────────────────────────────────────────────────────────────────────────────
class WarehouseDialog(QDialog):
    def __init__(self, parent=None, warehouse=None):
        super().__init__(parent)
        self.warehouse = warehouse
        self.setWindowTitle("إضافة مخزن" if not warehouse else "تعديل مخزن")
        self.setMinimumWidth(480)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        title = QLabel("🏭  " + ("إضافة مخزن جديد" if not self.warehouse else "تعديل بيانات المخزن"))
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        form = QFormLayout()
        form.setSpacing(10)

        self.code_edit = QLineEdit()
        self.name_edit = QLineEdit()
        self.manager_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.address_edit = QLineEdit()
        self.notes_edit = QTextEdit()
        self.notes_edit.setFixedHeight(70)
        self.active_cb = QCheckBox("نشط")
        self.active_cb.setChecked(True)

        form.addRow("كود المخزن *:", self.code_edit)
        form.addRow("اسم المخزن *:", self.name_edit)
        form.addRow("مدير المخزن:", self.manager_edit)
        form.addRow("الهاتف:", self.phone_edit)
        form.addRow("العنوان:", self.address_edit)
        form.addRow("ملاحظات:", self.notes_edit)
        form.addRow("", self.active_cb)
        lay.addLayout(form)

        if self.warehouse:
            self.code_edit.setText(self.warehouse['warehouse_code'])
            self.name_edit.setText(self.warehouse['warehouse_name'])
            self.manager_edit.setText(self.warehouse['warehouse_manager'] or "")
            self.phone_edit.setText(self.warehouse['phone'] or "")
            self.address_edit.setText(self.warehouse['address'] or "")
            self.notes_edit.setText(self.warehouse['notes'] or "")
            self.active_cb.setChecked(bool(self.warehouse['is_active']))

        btns = QHBoxLayout()
        btns.addStretch()
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setObjectName("btn_flat")
        cancel_btn.clicked.connect(self.reject)
        save_btn = QPushButton("💾  حفظ")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save)
        btns.addWidget(cancel_btn)
        btns.addWidget(save_btn)
        lay.addLayout(btns)

    def save(self):
        code = self.code_edit.text().strip()
        name = self.name_edit.text().strip()
        if not code or not name:
            QMessageBox.warning(self, "تنبيه", "كود المخزن والاسم مطلوبان")
            return
        data = (code, name, self.manager_edit.text().strip(),
                self.phone_edit.text().strip(), self.address_edit.text().strip(),
                self.notes_edit.toPlainText().strip(), int(self.active_cb.isChecked()))
        if self.warehouse:
            db.execute("""UPDATE warehouses SET warehouse_code=?, warehouse_name=?,
                warehouse_manager=?, phone=?, address=?, notes=?, is_active=? WHERE id=?""",
                data + (self.warehouse['id'],))
            log("تعديل مخزن", "warehouses", self.warehouse['id'], name)
        else:
            db.execute("""INSERT INTO warehouses (warehouse_code, warehouse_name, warehouse_manager,
                phone, address, notes, is_active) VALUES (?,?,?,?,?,?,?)""", data)
            log("إضافة مخزن", "warehouses", details=name)
        self.accept()


class WarehousesPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        hdr = QHBoxLayout()
        title = QLabel("🏭  المخازن")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        hdr.addWidget(title)
        hdr.addStretch()
        add_btn = QPushButton("➕  إضافة مخزن")
        add_btn.setObjectName("btn_success")
        add_btn.clicked.connect(self.add_warehouse)
        hdr.addWidget(add_btn)
        lay.addLayout(hdr)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["الكود", "الاسم", "المدير", "الهاتف", "العنوان", "الحالة", "إجراءات"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        lay.addWidget(self.table)
        self.load_data()

    def load_data(self):
        rows = db.fetchall("SELECT * FROM warehouses ORDER BY warehouse_code")
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(r['warehouse_code']))
            self.table.setItem(i, 1, QTableWidgetItem(r['warehouse_name']))
            self.table.setItem(i, 2, QTableWidgetItem(r['warehouse_manager'] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(r['phone'] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(r['address'] or ""))
            status = "✅ نشط" if r['is_active'] else "❌ غير نشط"
            si = QTableWidgetItem(status)
            si.setForeground(QColor(COLORS['accent_green'] if r['is_active'] else COLORS['accent_red']))
            self.table.setItem(i, 5, si)
            btns = QWidget()
            bl = QHBoxLayout(btns)
            bl.setContentsMargins(4, 2, 4, 2)
            bl.setSpacing(4)
            edit_btn = QPushButton("✏️")
            edit_btn.setFixedSize(30, 28)
            edit_btn.setObjectName("btn_flat")
            edit_btn.clicked.connect(lambda _, row=r: self.edit_warehouse(row))
            bl.addWidget(edit_btn)
            if has_perm("all"):
                del_btn = QPushButton("🗑")
                del_btn.setFixedSize(30, 28)
                del_btn.setObjectName("btn_danger")
                del_btn.clicked.connect(lambda _, row=r: self.delete_warehouse(row))
                bl.addWidget(del_btn)
            self.table.setCellWidget(i, 6, btns)

    def add_warehouse(self):
        dlg = WarehouseDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.load_data()

    def edit_warehouse(self, row):
        dlg = WarehouseDialog(self, dict(row))
        if dlg.exec() == QDialog.Accepted:
            self.load_data()

    def delete_warehouse(self, row):
        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد من حذف المخزن: {row['warehouse_name']}؟",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            db.execute("UPDATE warehouses SET is_active=0 WHERE id=?", (row['id'],))
            log("حذف مخزن", "warehouses", row['id'], row['warehouse_name'])
            self.load_data()


# ─────────────────────────────────────────────────────────────────────────────
# PRODUCTS PAGE (Enhanced)
# ─────────────────────────────────────────────────────────────────────────────
class ProductDialog(QDialog):
    def __init__(self, parent=None, product=None):
        super().__init__(parent)
        self.product = product
        self.setWindowTitle("إضافة منتج" if not product else "تعديل منتج")
        self.setMinimumSize(700, 700)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        tabs = QTabWidget()

        # Tab 1: Basic info
        t1 = QWidget()
        f1 = QFormLayout(t1)
        f1.setContentsMargins(20, 16, 20, 16)
        f1.setSpacing(10)

        self.name_edit = QLineEdit()
        self.desc_edit = QTextEdit()
        self.desc_edit.setFixedHeight(70)
        self.cat_combo = QComboBox()
        self.cat_combo.addItems([c['name'] for c in db.fetchall("SELECT name FROM categories ORDER BY name")])
        self.barcode_edit = QLineEdit()
        self.warehouse_combo = QComboBox()
        self.warehouse_combo.addItem("-- اختر المخزن --", None)
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.warehouse_combo.addItem(w['warehouse_name'], w['id'])
        self.location_edit = QLineEdit()
        self.rack_edit = QLineEdit()
        self.supplier_combo = QComboBox()
        self.supplier_combo.addItem("-- اختر المورد --", None)
        for s in db.fetchall("SELECT id, name FROM suppliers WHERE is_active=1 ORDER BY name"):
            self.supplier_combo.addItem(s['name'], s['id'])
        self.active_cb = QCheckBox("منتج نشط")
        self.active_cb.setChecked(True)
        self.is_asset_cb = QCheckBox("أصل ثابت (جهاز/معدة)")

        f1.addRow("اسم المنتج *:", self.name_edit)
        f1.addRow("الوصف:", self.desc_edit)
        f1.addRow("الفئة:", self.cat_combo)
        f1.addRow("الباركود:", self.barcode_edit)
        f1.addRow("المخزن:", self.warehouse_combo)
        f1.addRow("الموقع في المخزن:", self.location_edit)
        f1.addRow("رقم الرف:", self.rack_edit)
        f1.addRow("المورد:", self.supplier_combo)
        f1.addRow("", self.active_cb)
        f1.addRow("", self.is_asset_cb)
        tabs.addTab(t1, "المعلومات الأساسية")

        # Tab 2: Quantities & Prices
        t2 = QWidget()
        f2 = QFormLayout(t2)
        f2.setContentsMargins(20, 16, 20, 16)
        f2.setSpacing(10)
        self.qty_spin = QDoubleSpinBox()
        self.qty_spin.setRange(0, 9999999)
        self.qty_spin.setDecimals(2)
        self.min_qty_spin = QDoubleSpinBox()
        self.min_qty_spin.setRange(0, 9999999)
        self.alert_qty_spin = QDoubleSpinBox()
        self.alert_qty_spin.setRange(0, 9999999)
        self.purchase_price_spin = QDoubleSpinBox()
        self.purchase_price_spin.setRange(0, 9999999)
        self.purchase_price_spin.setDecimals(2)
        self.selling_price_spin = QDoubleSpinBox()
        self.selling_price_spin.setRange(0, 9999999)
        self.selling_price_spin.setDecimals(2)
        self.alarm_cb = QCheckBox("تفعيل تنبيه المخزون")
        self.alarm_cb.setChecked(True)
        f2.addRow("الكمية الحالية:", self.qty_spin)
        f2.addRow("الحد الأدنى (حرج):", self.min_qty_spin)
        f2.addRow("كمية التنبيه:", self.alert_qty_spin)
        f2.addRow("سعر الشراء:", self.purchase_price_spin)
        f2.addRow("سعر البيع:", self.selling_price_spin)
        f2.addRow("", self.alarm_cb)
        tabs.addTab(t2, "الكميات والأسعار")

        # Tab 3: Asset Tracking
        t3 = QWidget()
        f3 = QFormLayout(t3)
        f3.setContentsMargins(20, 16, 20, 16)
        f3.setSpacing(10)
        self.serial_edit = QLineEdit()
        self.asset_tag_edit = QLineEdit()
        self.purchase_date_edit = QDateEdit()
        self.purchase_date_edit.setCalendarPopup(True)
        self.purchase_date_edit.setDate(QDate.currentDate())
        self.warranty_end_edit = QDateEdit()
        self.warranty_end_edit.setCalendarPopup(True)
        self.warranty_end_edit.setDate(QDate.currentDate())
        self.current_owner_edit = QLineEdit()
        self.current_location_edit = QLineEdit()
        self.current_status_combo = QComboBox()
        self.current_status_combo.addItems(["في المخزن", "تم الصرف", "محوّل", "مُرجع", "تم التخلص", "مفقود"])
        f3.addRow("الرقم التسلسلي:", self.serial_edit)
        f3.addRow("رقم الأصل:", self.asset_tag_edit)
        f3.addRow("تاريخ الشراء:", self.purchase_date_edit)
        f3.addRow("نهاية الضمان:", self.warranty_end_edit)
        f3.addRow("المسؤول الحالي:", self.current_owner_edit)
        f3.addRow("الموقع الحالي:", self.current_location_edit)
        f3.addRow("الحالة:", self.current_status_combo)
        tabs.addTab(t3, "تتبع الأصل")

        lay.addWidget(tabs)

        if self.product:
            self._load_product()

        btns = QHBoxLayout()
        btns.setContentsMargins(20, 10, 20, 16)
        btns.addStretch()
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setObjectName("btn_flat")
        cancel_btn.clicked.connect(self.reject)
        save_btn = QPushButton("💾  حفظ")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save)
        btns.addWidget(cancel_btn)
        btns.addWidget(save_btn)
        lay.addLayout(btns)

    def _load_product(self):
        p = self.product
        self.name_edit.setText(p['name'])
        self.desc_edit.setText(p.get('description') or "")
        idx = self.cat_combo.findText(p.get('category') or "")
        if idx >= 0: self.cat_combo.setCurrentIndex(idx)
        self.barcode_edit.setText(p.get('barcode') or "")
        wid = p.get('warehouse_id')
        for j in range(self.warehouse_combo.count()):
            if self.warehouse_combo.itemData(j) == wid:
                self.warehouse_combo.setCurrentIndex(j); break
        self.location_edit.setText(p.get('warehouse_location') or "")
        self.rack_edit.setText(p.get('rack_number') or "")
        sid = p.get('supplier_id')
        for j in range(self.supplier_combo.count()):
            if self.supplier_combo.itemData(j) == sid:
                self.supplier_combo.setCurrentIndex(j); break
        self.active_cb.setChecked(bool(p.get('is_active', 1)))
        self.is_asset_cb.setChecked(bool(p.get('is_asset', 0)))
        self.qty_spin.setValue(p.get('quantity') or 0)
        self.min_qty_spin.setValue(p.get('min_quantity') or 0)
        self.alert_qty_spin.setValue(p.get('alert_quantity') or 0)
        self.purchase_price_spin.setValue(p.get('purchase_price') or 0)
        self.selling_price_spin.setValue(p.get('selling_price') or 0)
        self.alarm_cb.setChecked(bool(p.get('alarms_enabled', 1)))
        self.serial_edit.setText(p.get('serial_number') or "")
        self.asset_tag_edit.setText(p.get('asset_tag') or "")
        self.current_owner_edit.setText(p.get('current_owner') or "")
        self.current_location_edit.setText(p.get('current_location') or "")
        idx2 = self.current_status_combo.findText(p.get('current_status') or "في المخزن")
        if idx2 >= 0: self.current_status_combo.setCurrentIndex(idx2)

    def save(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "تنبيه", "اسم المنتج مطلوب")
            return
        wid = self.warehouse_combo.currentData()
        sid = self.supplier_combo.currentData()
        data = dict(
            name=name,
            description=self.desc_edit.toPlainText().strip(),
            category=self.cat_combo.currentText(),
            quantity=self.qty_spin.value(),
            min_quantity=self.min_qty_spin.value(),
            alert_quantity=self.alert_qty_spin.value(),
            purchase_price=self.purchase_price_spin.value(),
            selling_price=self.selling_price_spin.value(),
            barcode=self.barcode_edit.text().strip(),
            warehouse_location=self.location_edit.text().strip(),
            rack_number=self.rack_edit.text().strip(),
            warehouse_id=wid,
            supplier_id=sid,
            serial_number=self.serial_edit.text().strip(),
            asset_tag=self.asset_tag_edit.text().strip(),
            purchase_date=self.purchase_date_edit.date().toString("yyyy-MM-dd"),
            warranty_end=self.warranty_end_edit.date().toString("yyyy-MM-dd"),
            current_owner=self.current_owner_edit.text().strip(),
            current_location=self.current_location_edit.text().strip(),
            current_status=self.current_status_combo.currentText(),
            is_asset=int(self.is_asset_cb.isChecked()),
            is_active=int(self.active_cb.isChecked()),
            alarms_enabled=int(self.alarm_cb.isChecked()),
        )
        if self.product:
            db.execute("""UPDATE products SET
                name=:name, description=:description, category=:category,
                quantity=:quantity, min_quantity=:min_quantity, alert_quantity=:alert_quantity,
                purchase_price=:purchase_price, selling_price=:selling_price,
                barcode=:barcode, warehouse_location=:warehouse_location, rack_number=:rack_number,
                warehouse_id=:warehouse_id, supplier_id=:supplier_id,
                serial_number=:serial_number, asset_tag=:asset_tag,
                purchase_date=:purchase_date, warranty_end=:warranty_end,
                current_owner=:current_owner, current_location=:current_location,
                current_status=:current_status, is_asset=:is_asset,
                is_active=:is_active, alarms_enabled=:alarms_enabled,
                updated_at=datetime('now') WHERE id=?""", {**data, "?": self.product['id']})
            # Fix: use positional
            vals = [data[k] for k in ['name','description','category','quantity','min_quantity',
                'alert_quantity','purchase_price','selling_price','barcode','warehouse_location',
                'rack_number','warehouse_id','supplier_id','serial_number','asset_tag',
                'purchase_date','warranty_end','current_owner','current_location','current_status',
                'is_asset','is_active','alarms_enabled']] + [self.product['id']]
            db.execute("""UPDATE products SET
                name=?, description=?, category=?, quantity=?, min_quantity=?,
                alert_quantity=?, purchase_price=?, selling_price=?,
                barcode=?, warehouse_location=?, rack_number=?, warehouse_id=?, supplier_id=?,
                serial_number=?, asset_tag=?, purchase_date=?, warranty_end=?,
                current_owner=?, current_location=?, current_status=?, is_asset=?,
                is_active=?, alarms_enabled=?, updated_at=datetime('now') WHERE id=?""", vals)
            log("تعديل منتج", "products", self.product['id'], name)
        else:
            pid = db.generate_product_id()
            db.execute("""INSERT INTO products
                (product_id, name, description, category, quantity, min_quantity, alert_quantity,
                 purchase_price, selling_price, barcode, warehouse_location, rack_number,
                 warehouse_id, supplier_id, serial_number, asset_tag, purchase_date, warranty_end,
                 current_owner, current_location, current_status, is_asset, is_active, alarms_enabled)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (pid, data['name'], data['description'], data['category'],
                 data['quantity'], data['min_quantity'], data['alert_quantity'],
                 data['purchase_price'], data['selling_price'], data['barcode'],
                 data['warehouse_location'], data['rack_number'], data['warehouse_id'],
                 data['supplier_id'], data['serial_number'], data['asset_tag'],
                 data['purchase_date'], data['warranty_end'],
                 data['current_owner'], data['current_location'], data['current_status'],
                 data['is_asset'], data['is_active'], data['alarms_enabled']))
            log("إضافة منتج", "products", details=name)
        self.accept()


class ProductsPage(QWidget):
    stock_alert = Signal(list, list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        hdr = QHBoxLayout()
        title = QLabel("📦  المنتجات")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        hdr.addWidget(title)
        hdr.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍  بحث...")
        self.search_edit.setFixedWidth(220)
        self.search_edit.textChanged.connect(self.load_products)
        hdr.addWidget(self.search_edit)

        self.wh_filter = QComboBox()
        self.wh_filter.setFixedWidth(160)
        self.wh_filter.addItem("كل المخازن", None)
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.wh_filter.addItem(w['warehouse_name'], w['id'])
        self.wh_filter.currentIndexChanged.connect(self.load_products)
        hdr.addWidget(self.wh_filter)

        add_btn = QPushButton("➕  إضافة منتج")
        add_btn.setObjectName("btn_success")
        add_btn.clicked.connect(self.add_product)
        hdr.addWidget(add_btn)
        lay.addLayout(hdr)

        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels(["الكود", "الاسم", "الفئة", "المخزن", "الكمية",
                                               "الحد الأدنى", "التنبيه", "آخر سعر شراء", "المورد المفضل", "الحالة", "إجراءات"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self.view_trace)
        lay.addWidget(self.table)
        self.load_products()

    def load_products(self):
        search = self.search_edit.text().strip()
        wh_id = self.wh_filter.currentData()
        conditions = ["p.is_active=1"]
        params = []
        if search:
            conditions.append("(p.name LIKE ? OR p.product_id LIKE ? OR p.barcode LIKE ?)")
            params.extend([f"%{search}%"] * 3)
        if wh_id:
            conditions.append("p.warehouse_id=?")
            params.append(wh_id)
        where = " AND ".join(conditions)
        rows = db.fetchall(f"""
            SELECT p.*, w.warehouse_name, s.name as supplier_name
            FROM products p
            LEFT JOIN warehouses w ON p.warehouse_id = w.id
            LEFT JOIN suppliers s ON p.supplier_id = s.id
            WHERE {where} ORDER BY p.name
        """, params)
        self.table.setRowCount(len(rows))
        currency = db.get_setting("currency", "AED")
        warnings, criticals = [], []
        for i, r in enumerate(rows):
            qty = r['quantity'] or 0
            min_q = r['min_quantity'] or 0
            alert_q = r['alert_quantity'] or 0
            is_critical = qty <= min_q and r['alarms_enabled']
            is_warning = qty <= alert_q and not is_critical and r['alarms_enabled']
            if is_critical: criticals.append(dict(r))
            elif is_warning: warnings.append(dict(r))

            # Intelligence for each row
            insights = db.get_product_purchase_insights(r['id'])
            last_price = f"{currency} {insights['unit_price']:.2f}" if insights else f"{currency} {r['purchase_price']:.2f}"
            fav_supplier = insights['supplier_name'] if insights else (r['supplier_name'] or "—")

            for col, val in enumerate([r['product_id'], r['name'], r['category'] or "",
                                        r['warehouse_name'] or "", f"{qty:.0f}",
                                        f"{min_q:.0f}", f"{alert_q:.0f}",
                                        last_price, fav_supplier]):
                item = QTableWidgetItem(val)
                if is_critical:
                    item.setBackground(QColor(COLORS['row_critical']))
                elif is_warning:
                    item.setBackground(QColor(COLORS['row_warning']))
                self.table.setItem(i, col, item)

            status_item = QTableWidgetItem(r['current_status'] or "في المخزن")
            if is_critical:
                status_item.setForeground(QColor(COLORS['accent_red']))
                status_item.setText("🚨 حرج")
            elif is_warning:
                status_item.setForeground(QColor(COLORS['accent_orange']))
                status_item.setText("⚠ تنبيه")
            self.table.setItem(i, 8, status_item)

            btns = QWidget()
            bl = QHBoxLayout(btns)
            bl.setContentsMargins(2, 1, 2, 1)
            bl.setSpacing(3)
            eb = QPushButton("✏️")
            eb.setFixedSize(28, 26)
            eb.setObjectName("btn_flat")
            eb.clicked.connect(lambda _, row=dict(r): self.edit_product(row))
            tb_btn = QPushButton("📋")
            tb_btn.setFixedSize(28, 26)
            tb_btn.setObjectName("btn_flat")
            tb_btn.setToolTip("عرض السجل الكامل")
            tb_btn.clicked.connect(lambda _, row=dict(r): self.show_trace(row))
            bl.addWidget(eb)
            bl.addWidget(tb_btn)
            if has_perm("all"):
                dbb = QPushButton("🗑")
                dbb.setFixedSize(28, 26)
                dbb.setObjectName("btn_danger")
                dbb.clicked.connect(lambda _, row=dict(r): self.delete_product(row))
                bl.addWidget(dbb)
            self.table.setCellWidget(i, 10, btns)

        if warnings or criticals:
            self.stock_alert.emit(warnings, criticals)

    def add_product(self):
        dlg = ProductDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.load_products()

    def edit_product(self, row):
        dlg = ProductDialog(self, row)
        if dlg.exec() == QDialog.Accepted:
            self.load_products()

    def delete_product(self, row):
        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد من حذف المنتج: {row['name']}؟",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            db.execute("UPDATE products SET is_active=0 WHERE id=?", (row['id'],))
            log("حذف منتج", "products", row['id'], row['name'])
            self.load_products()

    def view_trace(self, index):
        row = db.fetchone("SELECT * FROM products WHERE product_id=?",
                          (self.table.item(index.row(), 0).text(),))
        if row:
            self.show_trace(dict(row))

    def show_trace(self, product):
        dlg = ItemTraceDialog(self, product)
        dlg.exec()


# ─────────────────────────────────────────────────────────────────────────────
# ITEM TRACE DIALOG — "Where did this item go?"
# ─────────────────────────────────────────────────────────────────────────────
class ItemTraceDialog(QDialog):
    def __init__(self, parent, product):
        super().__init__(parent)
        self.product = product
        self.setWindowTitle(f"تتبع المنتج: {product['name']}")
        self.setMinimumSize(800, 600)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        # Product summary card
        sc = QFrame()
        sc.setObjectName("card")
        scl = QHBoxLayout(sc)
        scl.setContentsMargins(16, 12, 16, 12)
        for lbl, val in [
            ("🆔  كود المنتج", self.product.get('product_id', '')),
            ("📦  الاسم", self.product['name']),
            ("📍  الموقع الحالي", self.product.get('current_location') or self.product.get('warehouse_location') or "—"),
            ("👤  المسؤول", self.product.get('current_owner') or "—"),
            ("⚙ الحالة", self.product.get('current_status') or "في المخزن"),
        ]:
            v = QVBoxLayout()
            v.setSpacing(2)
            tl = QLabel(lbl)
            tl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 10px;")
            vl = QLabel(str(val))
            vl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 13px; font-weight: 600;")
            v.addWidget(tl); v.addWidget(vl)
            scl.addLayout(v)
            if lbl != "⚙ الحالة":
                sep = QFrame()
                sep.setFixedWidth(1)
                sep.setStyleSheet(f"background: {COLORS['border2']};")
                scl.addWidget(sep)
        lay.addWidget(sc)

        title = QLabel("📜  السجل الكامل للحركات")
        title.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        table = QTableWidget()
        table.setColumnCount(9)
        table.setHorizontalHeaderLabels(["رقم الحركة", "النوع", "الكمية", "المخزن", "صُرف إلى",
                                          "القسم", "صُرف بواسطة", "اعتمد بواسطة", "التاريخ"])
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)

        movements = db.get_item_trace(self.product['id'])
        table.setRowCount(len(movements))
        type_colors = {"وارد": COLORS['accent_green'], "صادر": COLORS['accent_red'], "تحويل": COLORS['accent_orange'], "إرجاع": COLORS['accent_purple']}
        for i, m in enumerate(movements):
            tc = type_colors.get(m['movement_type'], COLORS['text_primary'])
            for col, val in enumerate([
                m['movement_number'] or "", m['movement_type'] or "",
                f"{m['quantity']:.0f}", m['warehouse_name'] or "",
                m['issued_to_name'] or "—", m['issued_to_department'] or "—",
                m['issued_by_name'] or "—", m['approved_by_name'] or "—",
                (m['created_at'] or "")[:16]
            ]):
                item = QTableWidgetItem(val)
                if col == 1:
                    item.setForeground(QColor(tc))
                table.setItem(i, col, item)
        lay.addWidget(table)

        close_btn = QPushButton("إغلاق")
        close_btn.setObjectName("btn_flat")
        close_btn.clicked.connect(self.accept)
        lay.addWidget(close_btn, alignment=Qt.AlignRight)


# ─────────────────────────────────────────────────────────────────────────────
# SUPPLIERS PAGE
# ─────────────────────────────────────────────────────────────────────────────
class SupplierDialog(QDialog):
    def __init__(self, parent=None, supplier=None):
        super().__init__(parent)
        self.supplier = supplier
        self.setWindowTitle("إضافة مورد" if not supplier else "تعديل مورد")
        self.setMinimumWidth(480)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)
        title = QLabel("🏢  " + ("إضافة مورد جديد" if not self.supplier else "تعديل بيانات المورد"))
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['text_primary']};")
        lay.addWidget(title)
        form = QFormLayout()
        form.setSpacing(10)
        self.name_edit = QLineEdit()
        self.company_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.email_edit = QLineEdit()
        self.country_edit = QLineEdit()
        self.address_edit = QLineEdit()
        self.notes_edit = QTextEdit()
        self.notes_edit.setFixedHeight(60)
        self.active_cb = QCheckBox("مورد نشط")
        self.active_cb.setChecked(True)
        form.addRow("اسم المورد *:", self.name_edit)
        form.addRow("الشركة:", self.company_edit)
        form.addRow("الهاتف:", self.phone_edit)
        form.addRow("البريد الإلكتروني:", self.email_edit)
        form.addRow("الدولة:", self.country_edit)
        form.addRow("العنوان:", self.address_edit)
        form.addRow("ملاحظات:", self.notes_edit)
        form.addRow("", self.active_cb)
        lay.addLayout(form)
        if self.supplier:
            self.name_edit.setText(self.supplier['name'])
            self.company_edit.setText(self.supplier.get('company') or "")
            self.phone_edit.setText(self.supplier.get('phone') or "")
            self.email_edit.setText(self.supplier.get('email') or "")
            self.country_edit.setText(self.supplier.get('country') or "")
            self.address_edit.setText(self.supplier.get('address') or "")
            self.notes_edit.setText(self.supplier.get('notes') or "")
            self.active_cb.setChecked(bool(self.supplier.get('is_active', 1)))
        btns = QHBoxLayout()
        btns.addStretch()
        cb = QPushButton("إلغاء"); cb.setObjectName("btn_flat"); cb.clicked.connect(self.reject)
        sb = QPushButton("💾  حفظ"); sb.setObjectName("btn_success"); sb.clicked.connect(self.save)
        btns.addWidget(cb); btns.addWidget(sb)
        lay.addLayout(btns)

    def save(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "تنبيه", "اسم المورد مطلوب")
            return
        data = (name, self.company_edit.text().strip(), self.phone_edit.text().strip(),
                self.email_edit.text().strip(), self.country_edit.text().strip(),
                self.address_edit.text().strip(), self.notes_edit.toPlainText().strip(),
                int(self.active_cb.isChecked()))
        if self.supplier:
            db.execute("""UPDATE suppliers SET name=?, company=?, phone=?, email=?,
                country=?, address=?, notes=?, is_active=? WHERE id=?""", data + (self.supplier['id'],))
        else:
            db.execute("""INSERT INTO suppliers (name, company, phone, email, country, address, notes, is_active)
                VALUES (?,?,?,?,?,?,?,?)""", data)
        log("حفظ مورد", "suppliers", details=name)
        self.accept()


class SuppliersPage(QWidget):
    suppliers_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)
        hdr = QHBoxLayout()
        title = QLabel("🏢  الموردين")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        hdr.addWidget(title); hdr.addStretch()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍  بحث...")
        self.search_edit.setFixedWidth(200)
        self.search_edit.textChanged.connect(self.load_data)
        hdr.addWidget(self.search_edit)
        add_btn = QPushButton("➕  إضافة مورد")
        add_btn.setObjectName("btn_success")
        add_btn.clicked.connect(self.add_supplier)
        hdr.addWidget(add_btn)
        lay.addLayout(hdr)
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["الاسم", "الشركة", "الهاتف", "البريد", "الدولة", "الحالة", "إجراءات"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        lay.addWidget(self.table)
        self.load_data()

    def load_data(self):
        search = self.search_edit.text().strip()
        params = [f"%{search}%", f"%{search}%"] if search else []
        where = "WHERE (name LIKE ? OR company LIKE ?)" if search else ""
        rows = db.fetchall(f"SELECT * FROM suppliers {where} ORDER BY name", params)
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for col, val in enumerate([r['name'], r['company'] or "", r['phone'] or "", r['email'] or "", r['country'] or ""]):
                self.table.setItem(i, col, QTableWidgetItem(val))
            si = QTableWidgetItem("✅ نشط" if r['is_active'] else "❌ غير نشط")
            si.setForeground(QColor(COLORS['accent_green'] if r['is_active'] else COLORS['accent_red']))
            self.table.setItem(i, 5, si)
            bw = QWidget(); bl = QHBoxLayout(bw); bl.setContentsMargins(2,1,2,1); bl.setSpacing(3)
            eb = QPushButton("✏️"); eb.setFixedSize(28,26); eb.setObjectName("btn_flat")
            eb.clicked.connect(lambda _, row=dict(r): self.edit_supplier(row))
            bl.addWidget(eb)
            if has_perm("all"):
                dbb = QPushButton("🗑")
                dbb.setFixedSize(28, 26)
                dbb.setObjectName("btn_danger")
                dbb.clicked.connect(lambda _, row=dict(r): self.delete_supplier(row))
                bl.addWidget(dbb)
            self.table.setCellWidget(i, 6, bw)

    def add_supplier(self):
        dlg = SupplierDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.load_data(); self.suppliers_changed.emit()

    def edit_supplier(self, row):
        dlg = SupplierDialog(self, row)
        if dlg.exec() == QDialog.Accepted:
            self.load_data(); self.suppliers_changed.emit()

    def delete_supplier(self, row):
        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد من حذف المورد: {row['name']}؟",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            db.execute("UPDATE suppliers SET is_active=0 WHERE id=?", (row['id'],))
            log("حذف مورد", "suppliers", row['id'], row['name'])
            self.load_data()
            self.suppliers_changed.emit()

    def refresh_suppliers(self):
        self.load_data()


# ─────────────────────────────────────────────────────────────────────────────
# STOCK IN PAGE (Enhanced with warehouse selection)
# ─────────────────────────────────────────────────────────────────────────────
class StockInPage(QWidget):
    stock_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.items = []
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("📥  إضافة للمخزون")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        tabs = QTabWidget()

        # Tab: New Receipt
        t1 = QWidget()
        t1l = QVBoxLayout(t1)
        t1l.setContentsMargins(16, 14, 16, 14)
        t1l.setSpacing(10)

        form = QGridLayout()
        form.setSpacing(10)
        form.setColumnStretch(1, 1); form.setColumnStretch(3, 1)

        self.inv_num_edit = QLineEdit(db.generate_stock_in_number())
        self.inv_num_edit.setReadOnly(True)
        self.warehouse_combo = QComboBox()
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.warehouse_combo.addItem(w['warehouse_name'], w['id'])
        self.supplier_combo = QComboBox()
        self.supplier_combo.addItem("-- اختر المورد --", None)
        for s in db.fetchall("SELECT id, name FROM suppliers WHERE is_active=1 ORDER BY name"):
            self.supplier_combo.addItem(s['name'], s['id'])
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.notes_edit = QLineEdit()

        form.addWidget(QLabel("رقم الفاتورة:"), 0, 0); form.addWidget(self.inv_num_edit, 0, 1)
        form.addWidget(QLabel("المخزن:"), 0, 2); form.addWidget(self.warehouse_combo, 0, 3)
        form.addWidget(QLabel("المورد:"), 1, 0); form.addWidget(self.supplier_combo, 1, 1)
        form.addWidget(QLabel("التاريخ:"), 1, 2); form.addWidget(self.date_edit, 1, 3)
        form.addWidget(QLabel("ملاحظات:"), 2, 0); form.addWidget(self.notes_edit, 2, 1, 1, 3)
        t1l.addLayout(form)

        # Items
        items_grp = QGroupBox("الأصناف")
        ig = QVBoxLayout(items_grp)
        add_item_row = QHBoxLayout()
        add_item_row.setSpacing(8)
        self.item_combo = QComboBox()
        self.item_combo.setFixedWidth(250)
        self.refresh_products()
        self.item_qty = QDoubleSpinBox()
        self.item_qty.setRange(0.01, 9999999)
        self.item_qty.setValue(1)
        self.item_price = QDoubleSpinBox()
        self.item_price.setRange(0, 9999999)
        self.item_price.setDecimals(2)
        add_item_btn = QPushButton("➕  إضافة")
        add_item_btn.setObjectName("btn_success")
        add_item_btn.setFixedWidth(100)
        add_item_btn.clicked.connect(self.add_item)
        add_item_row.addWidget(QLabel("المنتج:")); add_item_row.addWidget(self.item_combo)
        add_item_row.addWidget(QLabel("الكمية:")); add_item_row.addWidget(self.item_qty)
        add_item_row.addWidget(QLabel("السعر:")); add_item_row.addWidget(self.item_price)
        add_item_row.addWidget(add_item_btn); add_item_row.addStretch()
        ig.addLayout(add_item_row)
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(5)
        self.items_table.setHorizontalHeaderLabels(["المنتج", "الكمية", "سعر الوحدة", "الإجمالي", "حذف"])
        self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.items_table.setFixedHeight(200)
        self.items_table.verticalHeader().setVisible(False)
        ig.addWidget(self.items_table)
        t1l.addWidget(items_grp)

        total_row = QHBoxLayout()
        self.total_lbl = QLabel("الإجمالي: AED 0.00")
        self.total_lbl.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['accent_green']};")
        total_row.addStretch(); total_row.addWidget(self.total_lbl)
        t1l.addLayout(total_row)

        save_btn = QPushButton("💾  حفظ الإضافة للمخزون")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save_stock_in)
        t1l.addWidget(save_btn)
        tabs.addTab(t1, "إضافة جديدة")

        # Tab: History
        t2 = QWidget()
        t2l = QVBoxLayout(t2)
        t2l.setContentsMargins(16, 14, 16, 14)
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(6)
        self.history_table.setHorizontalHeaderLabels(["رقم الفاتورة", "المخزن", "المورد", "التاريخ", "الإجمالي", "بواسطة"])
        self.history_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setVisible(False)
        t2l.addWidget(self.history_table)
        tabs.addTab(t2, "سجل الوارد")

        lay.addWidget(tabs)
        tabs.currentChanged.connect(lambda i: self.load_history() if i == 1 else None)

    def refresh_suppliers(self):
        self.supplier_combo.clear()
        self.supplier_combo.addItem("-- اختر المورد --", None)
        for s in db.fetchall("SELECT id, name FROM suppliers WHERE is_active=1 ORDER BY name"):
            self.supplier_combo.addItem(s['name'], s['id'])

    def refresh_products(self):
        self.item_combo.clear()
        for p in db.fetchall("SELECT id, name, product_id FROM products WHERE is_active=1 ORDER BY name"):
            self.item_combo.addItem(f"[{p['product_id']}] {p['name']}", p['id'])

    def add_item(self):
        pid = self.item_combo.currentData()
        pname = self.item_combo.currentText()
        qty = self.item_qty.value()
        price = self.item_price.value()
        if not pid or qty <= 0:
            return
        total = qty * price
        self.items.append({'product_id': pid, 'name': pname, 'quantity': qty, 'unit_price': price, 'total': total})
        self._refresh_items_table()

    def _refresh_items_table(self):
        self.items_table.setRowCount(len(self.items))
        grand = 0
        for i, it in enumerate(self.items):
            self.items_table.setItem(i, 0, QTableWidgetItem(it['name']))
            self.items_table.setItem(i, 1, QTableWidgetItem(f"{it['quantity']:.2f}"))
            self.items_table.setItem(i, 2, QTableWidgetItem(f"AED {it['unit_price']:.2f}"))
            self.items_table.setItem(i, 3, QTableWidgetItem(f"AED {it['total']:.2f}"))
            grand += it['total']
            db_btn = QPushButton("🗑")
            db_btn.setFixedSize(26, 24)
            db_btn.setObjectName("btn_danger")
            db_btn.clicked.connect(lambda _, idx=i: self._remove_item(idx))
            self.items_table.setCellWidget(i, 4, db_btn)
        self.total_lbl.setText(f"الإجمالي: AED {grand:,.2f}")

    def _remove_item(self, idx):
        if 0 <= idx < len(self.items):
            self.items.pop(idx)
            self._refresh_items_table()

    def save_stock_in(self):
        if not self.items:
            QMessageBox.warning(self, "تنبيه", "يرجى إضافة صنف واحد على الأقل")
            return
        inv_num = self.inv_num_edit.text().strip()
        wh_id = self.warehouse_combo.currentData()
        sup_id = self.supplier_combo.currentData()
        dt = self.date_edit.date().toString("yyyy-MM-dd")
        notes = self.notes_edit.text().strip()
        total = sum(it['total'] for it in self.items)
        uid = CURRENT_USER.get('id')

        cur = db.execute("""INSERT INTO stock_in (invoice_number, warehouse_id, supplier_id, date, notes, total_value, created_by)
            VALUES (?,?,?,?,?,?,?)""", (inv_num, wh_id, sup_id, dt, notes, total, uid))
        si_id = cur.lastrowid

        for it in self.items:
            db.execute("""INSERT INTO stock_in_items (stock_in_id, product_id, quantity, unit_price, total_price)
                VALUES (?,?,?,?,?)""", (si_id, it['product_id'], it['quantity'], it['unit_price'], it['total']))
            db.execute("UPDATE products SET quantity = quantity + ?, updated_at=datetime('now') WHERE id=?",
                       (it['quantity'], it['product_id']))
            # Log movement
            mv_num = db.generate_movement_number()
            db.execute("""INSERT INTO stock_movements
                (movement_number, movement_type, product_id, warehouse_id, quantity, date,
                 reference_number, received_from_supplier, issued_by_user, reason, current_destination)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (mv_num, "وارد", it['product_id'], wh_id, it['quantity'], dt,
                 inv_num, sup_id, uid, "إضافة للمخزون", db.fetchone("SELECT warehouse_name FROM warehouses WHERE id=?", (wh_id,))['warehouse_name'] if wh_id else "المخزن"))

        log("إضافة وارد", "stock_in", si_id, inv_num)
        QMessageBox.information(self, "تم", f"تم حفظ الوارد بنجاح\nرقم الفاتورة: {inv_num}")
        self.items = []
        self._refresh_items_table()
        self.inv_num_edit.setText(db.generate_stock_in_number())
        self.stock_changed.emit()

    def load_history(self):
        rows = db.fetchall("""
            SELECT si.*, w.warehouse_name, s.name as supplier_name, u.full_name as created_by_name
            FROM stock_in si
            LEFT JOIN warehouses w ON si.warehouse_id = w.id
            LEFT JOIN suppliers s ON si.supplier_id = s.id
            LEFT JOIN users u ON si.created_by = u.id
            ORDER BY si.created_at DESC LIMIT 100
        """)
        self.history_table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for col, val in enumerate([r['invoice_number'], r['warehouse_name'] or "", r['supplier_name'] or "",
                                        r['date'], f"AED {r['total_value']:,.2f}", r['created_by_name'] or ""]):
                self.history_table.setItem(i, col, QTableWidgetItem(val))


# ─────────────────────────────────────────────────────────────────────────────
# STOCK OUT PAGE — Full Request/Approval Workflow
# ─────────────────────────────────────────────────────────────────────────────
class StockOutPage(QWidget):
    stock_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.items = []
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("📤  صرف من المخزون")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        tabs = QTabWidget()

        # Tab 1: New Request
        t1 = QWidget()
        t1l = QVBoxLayout(t1)
        t1l.setContentsMargins(16, 14, 16, 14)
        t1l.setSpacing(10)

        form = QGridLayout()
        form.setSpacing(10)
        form.setColumnStretch(1, 1); form.setColumnStretch(3, 1)

        self.req_num_edit = QLineEdit(db.generate_stock_out_number())
        self.req_num_edit.setReadOnly(True)
        self.warehouse_combo = QComboBox()
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.warehouse_combo.addItem(w['warehouse_name'], w['id'])
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.dept_edit = QLineEdit()
        self.dept_edit.setPlaceholderText("مثال: قسم تقنية المعلومات")
        self.emp_edit = QLineEdit()
        self.emp_edit.setPlaceholderText("اسم الموظف الذي يستلم")
        self.emp_id_edit = QLineEdit()
        self.emp_id_edit.setPlaceholderText("رقم الهوية")
        self.emp_phone_edit = QLineEdit()
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("سبب الصرف")
        self.notes_edit = QLineEdit()

        form.addWidget(QLabel("رقم الطلب:"), 0, 0); form.addWidget(self.req_num_edit, 0, 1)
        form.addWidget(QLabel("المخزن:"), 0, 2); form.addWidget(self.warehouse_combo, 0, 3)
        form.addWidget(QLabel("التاريخ:"), 1, 0); form.addWidget(self.date_edit, 1, 1)
        form.addWidget(QLabel("القسم:"), 1, 2); form.addWidget(self.dept_edit, 1, 3)
        form.addWidget(QLabel("اسم المستلم *:"), 2, 0); form.addWidget(self.emp_edit, 2, 1)
        form.addWidget(QLabel("رقم الهوية:"), 2, 2); form.addWidget(self.emp_id_edit, 2, 3)
        form.addWidget(QLabel("هاتف المستلم:"), 3, 0); form.addWidget(self.emp_phone_edit, 3, 1)
        form.addWidget(QLabel("سبب الصرف *:"), 3, 2); form.addWidget(self.reason_edit, 3, 3)
        self.offer_code_edit = QLineEdit()
        self.offer_code_edit.setPlaceholderText("أدخل كود الخصم إن وجد")
        form.addWidget(QLabel("كود الخصم:"), 4, 0); form.addWidget(self.offer_code_edit, 4, 1)
        form.addWidget(QLabel("ملاحظات:"), 4, 2); form.addWidget(self.notes_edit, 4, 3)
        t1l.addLayout(form)

        # Items
        items_grp = QGroupBox("الأصناف المطلوبة")
        ig = QVBoxLayout(items_grp)
        arow = QHBoxLayout(); arow.setSpacing(8)
        self.item_combo = QComboBox()
        self.item_combo.setFixedWidth(280)
        self.refresh_products()
        self.item_qty = QDoubleSpinBox()
        self.item_qty.setRange(0.01, 9999999)
        self.item_qty.setValue(1)
        add_item_btn = QPushButton("➕  إضافة")
        add_item_btn.setObjectName("btn_success")
        add_item_btn.setFixedWidth(100)
        add_item_btn.clicked.connect(self.add_item)
        arow.addWidget(QLabel("المنتج:")); arow.addWidget(self.item_combo)
        arow.addWidget(QLabel("الكمية:")); arow.addWidget(self.item_qty)
        arow.addWidget(add_item_btn); arow.addStretch()
        ig.addLayout(arow)
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(4)
        self.items_table.setHorizontalHeaderLabels(["المنتج", "الكمية المطلوبة", "المخزون الحالي", "حذف"])
        self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.items_table.setFixedHeight(180)
        self.items_table.verticalHeader().setVisible(False)
        ig.addWidget(self.items_table)
        t1l.addWidget(items_grp)

        btns_row = QHBoxLayout()
        btns_row.addStretch()
        save_btn = QPushButton("💾  حفظ الطلب")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save_request)
        btns_row.addWidget(save_btn)
        t1l.addLayout(btns_row)
        tabs.addTab(t1, "طلب صرف جديد")

        # Tab 2: Pending Requests (Approval)
        t2 = QWidget()
        t2l = QVBoxLayout(t2)
        t2l.setContentsMargins(16, 14, 16, 14)
        t2l.setSpacing(8)
        pending_hdr = QHBoxLayout()
        pt = QLabel("الطلبات المعلقة")
        pt.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']};")
        pending_hdr.addWidget(pt); pending_hdr.addStretch()
        refresh_btn = QPushButton("🔄  تحديث")
        refresh_btn.setObjectName("btn_flat")
        refresh_btn.clicked.connect(self.load_pending)
        pending_hdr.addWidget(refresh_btn)
        t2l.addLayout(pending_hdr)
        self.pending_table = QTableWidget()
        self.pending_table.setColumnCount(8)
        self.pending_table.setHorizontalHeaderLabels(["رقم الطلب", "المخزن", "المستلم", "القسم", "السبب", "الحالة", "التاريخ", "إجراءات"])
        self.pending_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.pending_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.pending_table.setAlternatingRowColors(True)
        self.pending_table.verticalHeader().setVisible(False)
        t2l.addWidget(self.pending_table)
        tabs.addTab(t2, "الطلبات والاعتماد")

        # Tab 3: History
        t3 = QWidget()
        t3l = QVBoxLayout(t3)
        t3l.setContentsMargins(16, 14, 16, 14)
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(8)
        self.history_table.setHorizontalHeaderLabels(["رقم الطلب", "المخزن", "المستلم", "القسم", "الحالة", "اعتمد بواسطة", "التاريخ", "إجراءات"])
        self.history_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setVisible(False)
        t3l.addWidget(self.history_table)
        tabs.addTab(t3, "سجل الصادر")

        lay.addWidget(tabs)
        tabs.currentChanged.connect(lambda i: {1: self.load_pending, 2: self.load_history}.get(i, lambda: None)())

    def refresh_products(self):
        self.item_combo.clear()
        for p in db.fetchall("SELECT id, name, product_id, quantity FROM products WHERE is_active=1 AND quantity > 0 ORDER BY name"):
            self.item_combo.addItem(f"[{p['product_id']}] {p['name']} (متاح: {p['quantity']:.0f})", p['id'])

    def add_item(self):
        pid = self.item_combo.currentData()
        pname = self.item_combo.currentText()
        qty = self.item_qty.value()
        if not pid or qty <= 0:
            return
        prod = db.fetchone("SELECT quantity FROM products WHERE id=?", (pid,))
        available = prod['quantity'] if prod else 0
        if qty > available:
            QMessageBox.warning(self, "تنبيه", f"الكمية المطلوبة ({qty:.0f}) تتجاوز المخزون المتاح ({available:.0f})")
            return
        self.items.append({'product_id': pid, 'name': pname, 'quantity': qty, 'available': available})
        self._refresh_items_table()

    def _refresh_items_table(self):
        self.items_table.setRowCount(len(self.items))
        for i, it in enumerate(self.items):
            self.items_table.setItem(i, 0, QTableWidgetItem(it['name']))
            self.items_table.setItem(i, 1, QTableWidgetItem(f"{it['quantity']:.0f}"))
            avail_item = QTableWidgetItem(f"{it['available']:.0f}")
            avail_item.setForeground(QColor(COLORS['accent_green'] if it['available'] >= it['quantity'] else COLORS['accent_red']))
            self.items_table.setItem(i, 2, avail_item)
            db_btn = QPushButton("🗑")
            db_btn.setFixedSize(26, 24)
            db_btn.setObjectName("btn_danger")
            db_btn.clicked.connect(lambda _, idx=i: self._remove_item(idx))
            self.items_table.setCellWidget(i, 3, db_btn)

    def _remove_item(self, idx):
        if 0 <= idx < len(self.items):
            self.items.pop(idx)
            self._refresh_items_table()

    def save_request(self):
        emp = self.emp_edit.text().strip()
        reason = self.reason_edit.text().strip()
        if not emp or not reason:
            QMessageBox.warning(self, "تنبيه", "اسم المستلم وسبب الصرف مطلوبان")
            return
        if not self.items:
            QMessageBox.warning(self, "تنبيه", "يرجى إضافة صنف واحد على الأقل")
            return
        req_num = self.req_num_edit.text().strip()
        wh_id = self.warehouse_combo.currentData()
        dt = self.date_edit.date().toString("yyyy-MM-dd")
        uid = CURRENT_USER.get('id')

        cur = db.execute("""INSERT INTO stock_out
            (request_number, warehouse_id, department, employee, employee_id_number,
             employee_phone, date, reason, status, notes, created_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (req_num, wh_id, self.dept_edit.text().strip(), emp,
             self.emp_id_edit.text().strip(), self.emp_phone_edit.text().strip(),
             dt, reason, "معلق", self.notes_edit.text().strip(), uid))
        so_id = cur.lastrowid

        for it in self.items:
            db.execute("INSERT INTO stock_out_items (stock_out_id, product_id, quantity) VALUES (?,?,?)",
                       (so_id, it['product_id'], it['quantity']))

        offer_code = self.offer_code_edit.text().strip()
        if offer_code:
            off = db.fetchone("SELECT * FROM offers WHERE offer_key=? AND is_active=1", (offer_code,))
            if off:
                log(f"تطبيق خصم: {offer_code}", "offers", off['id'], f"القيمة: {off['discount_value']}")
            else:
                QMessageBox.warning(self, "تنبيه", "كود الخصم غير صحيح أو منتهي")

        log("إضافة طلب صرف", "stock_out", so_id, req_num)
        # Create notification
        db.execute("INSERT INTO notifications (title, message, type) VALUES (?,?,?)",
                   ("طلب صرف جديد", f"طلب رقم {req_num} من {emp}", "info"))

        QMessageBox.information(self, "تم", f"تم حفظ طلب الصرف\nرقم الطلب: {req_num}\nحالة الطلب: معلق — يحتاج اعتماد")
        self.items = []
        self._refresh_items_table()
        self.req_num_edit.setText(db.generate_stock_out_number())
        self.stock_changed.emit()

    def load_pending(self):
        rows = db.fetchall("""
            SELECT so.*, w.warehouse_name
            FROM stock_out so
            LEFT JOIN warehouses w ON so.warehouse_id = w.id
            WHERE so.status IN ('معلق', 'معتمد')
            ORDER BY so.created_at DESC
        """)
        self.pending_table.setRowCount(len(rows))
        status_colors = {"معلق": COLORS['accent_orange'], "معتمد": COLORS['accent_green'], "مرفوض": COLORS['accent_red']}
        for i, r in enumerate(rows):
            for col, val in enumerate([r['request_number'], r['warehouse_name'] or "",
                                        r['employee'] or "", r['department'] or "", r['reason'] or ""]):
                self.pending_table.setItem(i, col, QTableWidgetItem(val))
            si = QTableWidgetItem(r['status'])
            si.setForeground(QColor(status_colors.get(r['status'], COLORS['text_primary'])))
            self.pending_table.setItem(i, 5, si)
            self.pending_table.setItem(i, 6, QTableWidgetItem(r['date']))
            bw = QWidget(); bl = QHBoxLayout(bw); bl.setContentsMargins(2,1,2,1); bl.setSpacing(3)
            if r['status'] == 'معلق' and has_perm('approve'):
                appr = QPushButton("✅")
                appr.setFixedSize(28, 26)
                appr.setObjectName("btn_success")
                appr.setToolTip("اعتماد")
                appr.clicked.connect(lambda _, rid=r['id']: self.approve_request(rid))
                rej = QPushButton("❌")
                rej.setFixedSize(28, 26)
                rej.setObjectName("btn_danger")
                rej.setToolTip("رفض")
                rej.clicked.connect(lambda _, rid=r['id']: self.reject_request(rid))
                bl.addWidget(appr); bl.addWidget(rej)
            if r['status'] == 'معتمد':
                issue_btn = QPushButton("📤")
                issue_btn.setFixedSize(28, 26)
                issue_btn.setObjectName("btn_warning")
                issue_btn.setToolTip("تنفيذ الصرف")
                issue_btn.clicked.connect(lambda _, rid=r['id']: self.execute_issue(rid))
                bl.addWidget(issue_btn)
            self.pending_table.setCellWidget(i, 7, bw)

    def approve_request(self, so_id):
        uid = CURRENT_USER.get('id')
        db.execute("UPDATE stock_out SET status='معتمد', approved_by=?, approved_at=datetime('now') WHERE id=?",
                   (uid, so_id))
        log("اعتماد طلب صرف", "stock_out", so_id)
        QMessageBox.information(self, "تم", "تم اعتماد الطلب")
        self.load_pending()
        self.stock_changed.emit()

    def reject_request(self, so_id):
        uid = CURRENT_USER.get('id')
        db.execute("UPDATE stock_out SET status='مرفوض', approved_by=?, approved_at=datetime('now') WHERE id=?",
                   (uid, so_id))
        log("رفض طلب صرف", "stock_out", so_id)
        QMessageBox.information(self, "تم", "تم رفض الطلب")
        self.load_pending()

    def execute_issue(self, so_id):
        """Deduct stock and record movements."""
        so = db.fetchone("SELECT * FROM stock_out WHERE id=?", (so_id,))
        if not so:
            return
        items = db.fetchall("SELECT * FROM stock_out_items WHERE stock_out_id=?", (so_id,))
        uid = CURRENT_USER.get('id')

        for it in items:
            prod = db.fetchone("SELECT * FROM products WHERE id=?", (it['product_id'],))
            if not prod or prod['quantity'] < it['quantity']:
                QMessageBox.critical(self, "خطأ",
                    f"كمية غير كافية للمنتج: {prod['name'] if prod else it['product_id']}")
                return

        for it in items:
            db.execute("UPDATE products SET quantity=quantity-?, current_owner=?, current_location=?, current_status='تم الصرف', updated_at=datetime('now') WHERE id=?",
                       (it['quantity'], so['employee'], so['department'], it['product_id']))
            mv_num = db.generate_movement_number()
            wh = db.fetchone("SELECT warehouse_name FROM warehouses WHERE id=?", (so['warehouse_id'],))
            db.execute("""INSERT INTO stock_movements
                (movement_number, movement_type, product_id, warehouse_id, quantity, date,
                 reference_number, issued_to_name, issued_to_department, issued_to_phone,
                 issued_to_national_id, issued_by_user, approved_by_user, reason, current_destination)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (mv_num, "صادر", it['product_id'], so['warehouse_id'], it['quantity'],
                 so['date'], so['request_number'], so['employee'], so['department'],
                 so['employee_phone'], so['employee_id_number'],
                 uid, so['approved_by'], so['reason'],
                 so['department'] or "خارج المخزن"))

        db.execute("UPDATE stock_out SET status='تم التسليم', issued_by=?, issued_at=datetime('now') WHERE id=?",
                   (uid, so_id))
        log("تنفيذ صرف", "stock_out", so_id, so['request_number'])
        QMessageBox.information(self, "تم", f"تم تنفيذ الصرف بنجاح\nرقم الطلب: {so['request_number']}")
        self.load_pending()
        self.stock_changed.emit()

    def load_history(self):
        rows = db.fetchall("""
            SELECT so.*, w.warehouse_name, u.full_name as approved_by_name
            FROM stock_out so
            LEFT JOIN warehouses w ON so.warehouse_id = w.id
            LEFT JOIN users u ON so.approved_by = u.id
            ORDER BY so.created_at DESC LIMIT 100
        """)
        self.history_table.setRowCount(len(rows))
        status_colors = {"معلق": COLORS['accent_orange'], "معتمد": COLORS['accent_green'],
                         "مرفوض": COLORS['accent_red'], "تم التسليم": COLORS['accent'],
                         "مُرجع": COLORS['accent_purple']}
        for i, r in enumerate(rows):
            for col, val in enumerate([r['request_number'], r['warehouse_name'] or "",
                                        r['employee'] or "", r['department'] or ""]):
                self.history_table.setItem(i, col, QTableWidgetItem(val))
            si = QTableWidgetItem(r['status'])
            si.setForeground(QColor(status_colors.get(r['status'], COLORS['text_primary'])))
            self.history_table.setItem(i, 4, si)
            for col, val in [(5, r['approved_by_name'] or "—"), (6, r['date'])]:
                self.history_table.setItem(i, col, QTableWidgetItem(val))
            bw = QWidget(); bl = QHBoxLayout(bw); bl.setContentsMargins(2,1,2,1); bl.setSpacing(3)
            if r['status'] == 'تم التسليم':
                pr_btn = QPushButton("🖨️")
                pr_btn.setFixedSize(28, 26)
                pr_btn.setObjectName("btn_flat")
                pr_btn.setToolTip("طباعة الفاتورة")
                pr_btn.clicked.connect(lambda _, rid=r['id']: self.print_invoice(rid))
                bl.addWidget(pr_btn)
                ret_btn = QPushButton("↩️")
                ret_btn.setFixedSize(28, 26)
                ret_btn.setObjectName("btn_warning")
                ret_btn.setToolTip("إرجاع")
                ret_btn.clicked.connect(lambda _, rid=r['id']: self.return_items(rid))
                bl.addWidget(ret_btn)
            self.history_table.setCellWidget(i, 7, bw)

    def print_invoice(self, so_id):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        import arabic_reshaper
        from bidi.algorithm import get_display

        def ar(txt):
            if not txt: return ""
            reshaped_text = arabic_reshaper.reshape(str(txt))
            return get_display(reshaped_text)

        so = db.fetchone("SELECT so.*, w.warehouse_name FROM stock_out so LEFT JOIN warehouses w ON so.warehouse_id=w.id WHERE so.id=?", (so_id,))
        items = db.fetchall("SELECT soi.*, p.name, p.product_id FROM stock_out_items soi JOIN products p ON soi.product_id=p.id WHERE soi.stock_out_id=?", (so_id,))

        path, _ = QFileDialog.getSaveFileName(self, "حفظ الفاتورة", f"فاتورة_{so['request_number']}.pdf", "PDF Files (*.pdf)")
        if not path: return

        doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()

        # Header
        title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=18, alignment=TA_CENTER, textColor=colors.HexColor(COLORS['accent']))
        elements.append(Paragraph(ar(COMPANY_NAME), title_style))
        if LOGO_PATH.exists():
            elements.append(RLImage(str(LOGO_PATH), width=50*mm, height=20*mm))
        elements.append(Spacer(1, 10))

        # Info Table
        info_data = [
            [ar(f"التاريخ: {so['date']}"), ar(f"رقم الطلب: {so['request_number']}")],
            [ar(f"المستلم: {so['employee']}"), ar(f"المخزن: {so['warehouse_name']}")],
            [ar(f"القسم: {so['department']}"), ar(f"السبب: {so['reason']}")]
        ]
        it = Table(info_data, colWidths=[270, 270])
        it.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
            ('SIZE', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(it)
        elements.append(Spacer(1, 20))

        # Items Table
        header = [ar("الإجمالي"), ar("الكمية"), ar("اسم الصنف"), ar("كود الصنف")]
        table_data = [header]
        for itm in items:
            table_data.append(["—", f"{itm['quantity']:.0f}", ar(itm['name']), itm['product_id']])

        t = Table(table_data, colWidths=[100, 80, 260, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(COLORS['bg_nav'])),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('SIZE', (0, 1), (-1, -1), 10),
        ]))
        elements.append(t)

        # Footer
        elements.append(Spacer(1, 40))
        elements.append(Paragraph(ar("توقيع المستلم: ..........................                توقيع المسؤول: .........................."), styles['Normal']))

        doc.build(elements)
        QMessageBox.information(self, "تم", f"تم إنشاء الفاتورة بنجاح:\n{path}")

    def return_items(self, so_id):
        dlg = ReturnDialog(self, so_id)
        if dlg.exec() == QDialog.Accepted:
            self.load_history()
            self.stock_changed.emit()


# ─────────────────────────────────────────────────────────────────────────────
# RETURN DIALOG
# ─────────────────────────────────────────────────────────────────────────────
class ReturnDialog(QDialog):
    def __init__(self, parent, so_id):
        super().__init__(parent)
        self.so_id = so_id
        self.so = db.fetchone("SELECT * FROM stock_out WHERE id=?", (so_id,))
        self.setWindowTitle("إرجاع أصناف")
        self.setMinimumWidth(500)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        title = QLabel(f"↩️  إرجاع من طلب: {self.so['request_number']}")
        title.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        form = QFormLayout()
        self.returned_by_edit = QLineEdit(self.so['employee'] or "")
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("سبب الإرجاع")
        self.condition_combo = QComboBox()
        self.condition_combo.addItems(["جيد", "تالف", "مفقود"])
        self.return_date_edit = QDateEdit(QDate.currentDate())
        self.return_date_edit.setCalendarPopup(True)
        form.addRow("مُرجع بواسطة:", self.returned_by_edit)
        form.addRow("سبب الإرجاع:", self.reason_edit)
        form.addRow("حالة الصنف:", self.condition_combo)
        form.addRow("تاريخ الإرجاع:", self.return_date_edit)
        lay.addLayout(form)

        items = db.fetchall("""
            SELECT soi.*, p.name, p.quantity as current_qty
            FROM stock_out_items soi JOIN products p ON soi.product_id = p.id
            WHERE soi.stock_out_id=?
        """, (self.so_id,))
        self.qty_spins = []
        grp = QGroupBox("الأصناف المُرجعة")
        gl = QFormLayout(grp)
        for it in items:
            spin = QDoubleSpinBox()
            spin.setRange(0, it['quantity'])
            spin.setValue(it['quantity'])
            spin.setDecimals(2)
            spin.setProperty("item", dict(it))
            gl.addRow(f"{it['name']} (المصروف: {it['quantity']:.0f}):", spin)
            self.qty_spins.append(spin)
        lay.addWidget(grp)

        btns = QHBoxLayout(); btns.addStretch()
        cb = QPushButton("إلغاء"); cb.setObjectName("btn_flat"); cb.clicked.connect(self.reject)
        sb = QPushButton("💾  تأكيد الإرجاع"); sb.setObjectName("btn_success"); sb.clicked.connect(self.save)
        btns.addWidget(cb); btns.addWidget(sb)
        lay.addLayout(btns)

    def save(self):
        returned_by = self.returned_by_edit.text().strip()
        reason = self.reason_edit.text().strip()
        condition = self.condition_combo.currentText()
        return_date = self.return_date_edit.date().toString("yyyy-MM-dd")
        uid = CURRENT_USER.get('id')
        wh_id = self.so['warehouse_id']

        for spin in self.qty_spins:
            qty = spin.value()
            it = spin.property("item")
            if qty <= 0:
                continue
            ret_num = db.generate_return_number()
            db.execute("""INSERT INTO returns
                (return_number, stock_out_id, product_id, warehouse_id, quantity,
                 returned_by, return_reason, return_date, condition, processed_by)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (ret_num, self.so_id, it['product_id'], wh_id, qty,
                 returned_by, reason, return_date, condition, uid))
            # Add back to stock
            db.execute("UPDATE products SET quantity=quantity+?, current_status='مُرجع', updated_at=datetime('now') WHERE id=?",
                       (qty, it['product_id']))
            mv_num = db.generate_movement_number()
            db.execute("""INSERT INTO stock_movements
                (movement_number, movement_type, product_id, warehouse_id, quantity,
                 date, reference_number, issued_to_name, issued_by_user, reason, current_destination)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (mv_num, "إرجاع", it['product_id'], wh_id, qty, return_date,
                 ret_num, returned_by, uid, reason, "إعادة للمخزن"))

        log("إرجاع أصناف", "returns", details=self.so['request_number'])
        QMessageBox.information(self, "تم", "تم تسجيل الإرجاع وإضافة الكميات للمخزون")
        self.accept()


# ─────────────────────────────────────────────────────────────────────────────
# TRANSFERS PAGE
# ─────────────────────────────────────────────────────────────────────────────
class TransfersPage(QWidget):
    stock_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.items = []
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("🔄  التحويلات بين المخازن")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        tabs = QTabWidget()

        # Tab 1: New Transfer
        t1 = QWidget()
        t1l = QVBoxLayout(t1)
        t1l.setContentsMargins(16, 14, 16, 14)
        t1l.setSpacing(10)

        form = QGridLayout(); form.setSpacing(10)
        form.setColumnStretch(1, 1); form.setColumnStretch(3, 1)

        self.tr_num_edit = QLineEdit(db.generate_transfer_number())
        self.tr_num_edit.setReadOnly(True)
        warehouses = db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1")
        self.src_combo = QComboBox()
        self.dst_combo = QComboBox()
        for w in warehouses:
            self.src_combo.addItem(w['warehouse_name'], w['id'])
            self.dst_combo.addItem(w['warehouse_name'], w['id'])
        if len(warehouses) > 1:
            self.dst_combo.setCurrentIndex(1)
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.reason_edit = QLineEdit()

        form.addWidget(QLabel("رقم التحويل:"), 0, 0); form.addWidget(self.tr_num_edit, 0, 1)
        form.addWidget(QLabel("التاريخ:"), 0, 2); form.addWidget(self.date_edit, 0, 3)
        form.addWidget(QLabel("من مخزن:"), 1, 0); form.addWidget(self.src_combo, 1, 1)
        form.addWidget(QLabel("إلى مخزن:"), 1, 2); form.addWidget(self.dst_combo, 1, 3)
        form.addWidget(QLabel("سبب التحويل:"), 2, 0); form.addWidget(self.reason_edit, 2, 1, 1, 3)
        t1l.addLayout(form)

        items_grp = QGroupBox("الأصناف المحوّلة")
        ig = QVBoxLayout(items_grp)
        arow = QHBoxLayout(); arow.setSpacing(8)
        self.item_combo = QComboBox(); self.item_combo.setFixedWidth(280)
        self.refresh_products()
        self.item_qty = QDoubleSpinBox()
        self.item_qty.setRange(0.01, 9999999)
        self.item_qty.setValue(1)
        add_btn = QPushButton("➕  إضافة")
        add_btn.setObjectName("btn_success")
        add_btn.setFixedWidth(100)
        add_btn.clicked.connect(self.add_item)
        arow.addWidget(QLabel("المنتج:")); arow.addWidget(self.item_combo)
        arow.addWidget(QLabel("الكمية:")); arow.addWidget(self.item_qty)
        arow.addWidget(add_btn); arow.addStretch()
        ig.addLayout(arow)
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(3)
        self.items_table.setHorizontalHeaderLabels(["المنتج", "الكمية", "حذف"])
        self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.items_table.setFixedHeight(180)
        self.items_table.verticalHeader().setVisible(False)
        ig.addWidget(self.items_table)
        t1l.addWidget(items_grp)

        save_btn = QPushButton("🔄  تنفيذ التحويل")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save_transfer)
        t1l.addWidget(save_btn, alignment=Qt.AlignRight)
        tabs.addTab(t1, "تحويل جديد")

        # Tab 2: History
        t2 = QWidget()
        t2l = QVBoxLayout(t2)
        t2l.setContentsMargins(16, 14, 16, 14)
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(7)
        self.history_table.setHorizontalHeaderLabels(["رقم التحويل", "من مخزن", "إلى مخزن", "السبب", "الحالة", "بواسطة", "التاريخ"])
        self.history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setVisible(False)
        t2l.addWidget(self.history_table)
        tabs.addTab(t2, "سجل التحويلات")

        lay.addWidget(tabs)
        tabs.currentChanged.connect(lambda i: self.load_history() if i == 1 else None)

    def refresh_products(self):
        self.item_combo.clear()
        for p in db.fetchall("SELECT id, name, product_id, quantity FROM products WHERE is_active=1 AND quantity > 0 ORDER BY name"):
            self.item_combo.addItem(f"[{p['product_id']}] {p['name']} ({p['quantity']:.0f})", p['id'])

    def add_item(self):
        pid = self.item_combo.currentData()
        pname = self.item_combo.currentText()
        qty = self.item_qty.value()
        if not pid or qty <= 0:
            return
        self.items.append({'product_id': pid, 'name': pname, 'quantity': qty})
        self.items_table.setRowCount(len(self.items))
        for i, it in enumerate(self.items):
            self.items_table.setItem(i, 0, QTableWidgetItem(it['name']))
            self.items_table.setItem(i, 1, QTableWidgetItem(f"{it['quantity']:.0f}"))
            db_btn = QPushButton("🗑"); db_btn.setFixedSize(26,24); db_btn.setObjectName("btn_danger")
            db_btn.clicked.connect(lambda _, idx=i: (self.items.pop(idx), self.add_item.__func__(self)))
            self.items_table.setCellWidget(i, 2, db_btn)

    def save_transfer(self):
        if not self.items:
            QMessageBox.warning(self, "تنبيه", "أضف صنفاً على الأقل")
            return
        src_id = self.src_combo.currentData()
        dst_id = self.dst_combo.currentData()
        if src_id == dst_id:
            QMessageBox.warning(self, "تنبيه", "المخزن المصدر والوجهة لا يمكن أن يكونا متطابقين")
            return
        tr_num = self.tr_num_edit.text()
        dt = self.date_edit.date().toString("yyyy-MM-dd")
        reason = self.reason_edit.text().strip()
        uid = CURRENT_USER.get('id')

        cur = db.execute("""INSERT INTO transfers (transfer_number, source_warehouse_id, dest_warehouse_id,
            transferred_by, reason, status, date) VALUES (?,?,?,?,?,?,?)""",
            (tr_num, src_id, dst_id, uid, reason, "مكتمل", dt))
        tr_id = cur.lastrowid

        dst_name = self.dst_combo.currentText()
        for it in self.items:
            db.execute("INSERT INTO transfer_items (transfer_id, product_id, quantity) VALUES (?,?,?)",
                       (tr_id, it['product_id'], it['quantity']))
            db.execute("UPDATE products SET quantity=quantity-?, updated_at=datetime('now') WHERE id=?",
                       (it['quantity'], it['product_id']))
            db.execute("UPDATE products SET warehouse_id=?, current_status='محوّل', updated_at=datetime('now') WHERE id=?",
                       (dst_id, it['product_id']))
            mv_num = db.generate_movement_number()
            db.execute("""INSERT INTO stock_movements
                (movement_number, movement_type, product_id, warehouse_id, quantity,
                 date, reference_number, issued_by_user, reason, current_destination)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (mv_num, "تحويل", it['product_id'], src_id, it['quantity'],
                 dt, tr_num, uid, reason, dst_name))

        log("تحويل مخزون", "transfers", tr_id, tr_num)
        QMessageBox.information(self, "تم", f"تم التحويل بنجاح\nرقم التحويل: {tr_num}")
        self.items = []
        self.items_table.setRowCount(0)
        self.tr_num_edit.setText(db.generate_transfer_number())
        self.stock_changed.emit()

    def load_history(self):
        rows = db.fetchall("""
            SELECT t.*, w1.warehouse_name as src_name, w2.warehouse_name as dst_name,
                   u.full_name as by_name
            FROM transfers t
            LEFT JOIN warehouses w1 ON t.source_warehouse_id = w1.id
            LEFT JOIN warehouses w2 ON t.dest_warehouse_id = w2.id
            LEFT JOIN users u ON t.transferred_by = u.id
            ORDER BY t.created_at DESC LIMIT 100
        """)
        self.history_table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for col, val in enumerate([r['transfer_number'], r['src_name'] or "", r['dst_name'] or "",
                                        r['reason'] or "", r['status'], r['by_name'] or "", r['date']]):
                self.history_table.setItem(i, col, QTableWidgetItem(val))


# ─────────────────────────────────────────────────────────────────────────────
# REPORTS PAGE
# ─────────────────────────────────────────────────────────────────────────────
class ReportsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("📊  التقارير")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        tabs = QTabWidget()

        # ── Movement Report
        t_mv = QWidget()
        t_mv_l = QVBoxLayout(t_mv)
        t_mv_l.setContentsMargins(14, 12, 14, 12)
        t_mv_l.setSpacing(10)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("من:"))
        self.mv_from = QDateEdit(QDate.currentDate().addDays(-30))
        self.mv_from.setCalendarPopup(True)
        filter_row.addWidget(self.mv_from)
        filter_row.addWidget(QLabel("إلى:"))
        self.mv_to = QDateEdit(QDate.currentDate())
        self.mv_to.setCalendarPopup(True)
        filter_row.addWidget(self.mv_to)
        self.mv_type_combo = QComboBox()
        self.mv_type_combo.addItems(["الكل", "وارد", "صادر", "تحويل", "إرجاع"])
        self.mv_type_combo.setFixedWidth(120)
        filter_row.addWidget(QLabel("النوع:")); filter_row.addWidget(self.mv_type_combo)
        filter_row.addStretch()
        load_btn = QPushButton("🔄  تحميل"); load_btn.clicked.connect(self.load_movements)
        export_btn = QPushButton("📊  تصدير Excel"); export_btn.setObjectName("btn_success")
        export_btn.clicked.connect(self.export_movements)
        filter_row.addWidget(load_btn); filter_row.addWidget(export_btn)
        t_mv_l.addLayout(filter_row)
        self.mv_table = QTableWidget()
        self.mv_table.setColumnCount(9)
        self.mv_table.setHorizontalHeaderLabels(["رقم الحركة", "النوع", "المنتج", "المخزن",
                                                   "الكمية", "صُرف إلى", "القسم", "بواسطة", "التاريخ"])
        self.mv_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.mv_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.mv_table.setAlternatingRowColors(True)
        self.mv_table.verticalHeader().setVisible(False)
        t_mv_l.addWidget(self.mv_table)
        tabs.addTab(t_mv, "حركات المخزون")

        # ── Inventory Report
        t_inv = QWidget()
        t_inv_l = QVBoxLayout(t_inv)
        t_inv_l.setContentsMargins(14, 12, 14, 12)
        t_inv_l.setSpacing(10)
        inv_filter = QHBoxLayout()
        self.inv_wh_combo = QComboBox()
        self.inv_wh_combo.addItem("كل المخازن", None)
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.inv_wh_combo.addItem(w['warehouse_name'], w['id'])
        self.inv_low_cb = QCheckBox("مخزون منخفض فقط")
        inv_filter.addWidget(QLabel("المخزن:")); inv_filter.addWidget(self.inv_wh_combo)
        inv_filter.addWidget(self.inv_low_cb); inv_filter.addStretch()
        load_inv_btn = QPushButton("🔄  تحميل"); load_inv_btn.clicked.connect(self.load_inventory_report)
        export_inv_btn = QPushButton("📊  تصدير Excel"); export_inv_btn.setObjectName("btn_success")
        export_inv_btn.clicked.connect(self.export_inventory)
        inv_filter.addWidget(load_inv_btn); inv_filter.addWidget(export_inv_btn)
        t_inv_l.addLayout(inv_filter)
        self.inv_table = QTableWidget()
        self.inv_table.setColumnCount(9)
        self.inv_table.setHorizontalHeaderLabels(["الكود", "الاسم", "الفئة", "المخزن", "الكمية",
                                                    "الحد الأدنى", "سعر الشراء", "القيمة الكلية", "الحالة"])
        self.inv_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.inv_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.inv_table.setAlternatingRowColors(True)
        self.inv_table.verticalHeader().setVisible(False)
        t_inv_l.addWidget(self.inv_table)
        tabs.addTab(t_inv, "تقرير المخزون")

        # ── Issue Report
        t_iss = QWidget()
        t_iss_l = QVBoxLayout(t_iss)
        t_iss_l.setContentsMargins(14, 12, 14, 12)
        t_iss_l.setSpacing(10)
        iss_filter = QHBoxLayout()
        iss_filter.addWidget(QLabel("من:"))
        self.iss_from = QDateEdit(QDate.currentDate().addDays(-30)); self.iss_from.setCalendarPopup(True)
        self.iss_to = QDateEdit(QDate.currentDate()); self.iss_to.setCalendarPopup(True)
        iss_filter.addWidget(self.iss_from)
        iss_filter.addWidget(QLabel("إلى:")); iss_filter.addWidget(self.iss_to)
        iss_filter.addStretch()
        load_iss_btn = QPushButton("🔄  تحميل"); load_iss_btn.clicked.connect(self.load_issue_report)
        iss_filter.addWidget(load_iss_btn)
        t_iss_l.addLayout(iss_filter)
        self.iss_table = QTableWidget()
        self.iss_table.setColumnCount(8)
        self.iss_table.setHorizontalHeaderLabels(["رقم الطلب", "المخزن", "المستلم",
                                                    "القسم", "السبب", "الحالة", "اعتمد بواسطة", "التاريخ"])
        self.iss_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.iss_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.iss_table.setAlternatingRowColors(True)
        self.iss_table.verticalHeader().setVisible(False)
        t_iss_l.addWidget(self.iss_table)
        tabs.addTab(t_iss, "تقرير الصادر")

        # ── Audit Log Report
        t_log = QWidget()
        t_log_l = QVBoxLayout(t_log)
        t_log_l.setContentsMargins(14, 12, 14, 12)
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(6)
        self.log_table.setHorizontalHeaderLabels(["المستخدم", "الإجراء", "الجدول", "رقم السجل", "التفاصيل", "التاريخ"])
        self.log_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.log_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.log_table.setAlternatingRowColors(True)
        self.log_table.verticalHeader().setVisible(False)
        t_log_l.addWidget(self.log_table)
        load_log_btn = QPushButton("🔄  تحميل سجل الأنشطة"); load_log_btn.clicked.connect(self.load_audit_log)
        t_log_l.addWidget(load_log_btn)
        tabs.addTab(t_log, "سجل الأنشطة")

        lay.addWidget(tabs)

    def load_movements(self):
        frm = self.mv_from.date().toString("yyyy-MM-dd")
        to = self.mv_to.date().toString("yyyy-MM-dd")
        mv_type = self.mv_type_combo.currentText()
        params = [frm, to]
        type_clause = "" if mv_type == "الكل" else " AND sm.movement_type=?"
        if mv_type != "الكل":
            params.append(mv_type)
        rows = db.fetchall(f"""
            SELECT sm.*, p.name as product_name, w.warehouse_name,
                   u.full_name as issued_by_name
            FROM stock_movements sm
            LEFT JOIN products p ON sm.product_id = p.id
            LEFT JOIN warehouses w ON sm.warehouse_id = w.id
            LEFT JOIN users u ON sm.issued_by_user = u.id
            WHERE sm.date >= ? AND sm.date <= ? {type_clause}
            ORDER BY sm.created_at DESC
        """, params)
        self.mv_table.setRowCount(len(rows))
        type_colors = {"وارد": COLORS['accent_green'], "صادر": COLORS['accent_red'],
                       "تحويل": COLORS['accent_orange'], "إرجاع": COLORS['accent_purple']}
        for i, r in enumerate(rows):
            tc = type_colors.get(r['movement_type'], COLORS['text_primary'])
            for col, val in enumerate([r['movement_number'] or "", "", r['product_name'] or "",
                                        r['warehouse_name'] or "", f"{r['quantity']:.0f}",
                                        r['issued_to_name'] or "—", r['issued_to_department'] or "—",
                                        r['issued_by_name'] or "—", (r['created_at'] or "")[:16]]):
                item = QTableWidgetItem(val)
                self.mv_table.setItem(i, col, item)
            type_item = QTableWidgetItem(r['movement_type'] or "")
            type_item.setForeground(QColor(tc))
            self.mv_table.setItem(i, 1, type_item)

    def export_movements(self):
        self.load_movements()
        path, _ = QFileDialog.getSaveFileName(self, "حفظ كـ Excel", "تقرير_الحركات.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
        rows = []
        for i in range(self.mv_table.rowCount()):
            row = [self.mv_table.item(i, j).text() if self.mv_table.item(i, j) else "" for j in range(self.mv_table.columnCount())]
            rows.append(row)
        df = pd.DataFrame(rows, columns=["رقم الحركة", "النوع", "المنتج", "المخزن", "الكمية", "صُرف إلى", "القسم", "بواسطة", "التاريخ"])
        df.to_excel(path, index=False)
        QMessageBox.information(self, "تم", f"تم التصدير إلى:\n{path}")

    def load_inventory_report(self):
        wh_id = self.inv_wh_combo.currentData()
        low_only = self.inv_low_cb.isChecked()
        conds = ["p.is_active=1"]
        params = []
        if wh_id:
            conds.append("p.warehouse_id=?"); params.append(wh_id)
        if low_only:
            conds.append("p.quantity <= p.alert_quantity")
        rows = db.fetchall(f"""
            SELECT p.*, w.warehouse_name
            FROM products p
            LEFT JOIN warehouses w ON p.warehouse_id = w.id
            WHERE {' AND '.join(conds)}
            ORDER BY p.name
        """, params)
        self.inv_table.setRowCount(len(rows))
        total_val = 0
        for i, r in enumerate(rows):
            val = (r['quantity'] or 0) * (r['purchase_price'] or 0)
            total_val += val
            qty = r['quantity'] or 0
            min_q = r['min_quantity'] or 0
            status = "🚨 حرج" if qty <= min_q else ("⚠ تنبيه" if qty <= (r['alert_quantity'] or 0) else "✅ جيد")
            for col, val2 in enumerate([r['product_id'], r['name'], r['category'] or "",
                                         r['warehouse_name'] or "", f"{qty:.0f}",
                                         f"{min_q:.0f}", f"AED {r['purchase_price']:.2f}",
                                         f"AED {val:,.2f}", status]):
                item = QTableWidgetItem(val2)
                if "🚨" in status: item.setBackground(QColor(COLORS['row_critical']))
                elif "⚠" in status: item.setBackground(QColor(COLORS['row_warning']))
                self.inv_table.setItem(i, col, item)

    def export_inventory(self):
        self.load_inventory_report()
        path, _ = QFileDialog.getSaveFileName(self, "حفظ كـ Excel", "تقرير_المخزون.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
        rows = []
        for i in range(self.inv_table.rowCount()):
            rows.append([self.inv_table.item(i, j).text() if self.inv_table.item(i, j) else "" for j in range(self.inv_table.columnCount())])
        df = pd.DataFrame(rows, columns=["الكود","الاسم","الفئة","المخزن","الكمية","الحد الأدنى","سعر الشراء","القيمة الكلية","الحالة"])
        df.to_excel(path, index=False)
        QMessageBox.information(self, "تم", f"تم التصدير:\n{path}")

    def load_issue_report(self):
        frm = self.iss_from.date().toString("yyyy-MM-dd")
        to = self.iss_to.date().toString("yyyy-MM-dd")
        rows = db.fetchall("""
            SELECT so.*, w.warehouse_name, u.full_name as appr_name
            FROM stock_out so
            LEFT JOIN warehouses w ON so.warehouse_id = w.id
            LEFT JOIN users u ON so.approved_by = u.id
            WHERE so.date >= ? AND so.date <= ?
            ORDER BY so.date DESC
        """, (frm, to))
        self.iss_table.setRowCount(len(rows))
        status_colors = {"معلق": COLORS['accent_orange'], "معتمد": COLORS['accent_green'],
                         "مرفوض": COLORS['accent_red'], "تم التسليم": COLORS['accent']}
        for i, r in enumerate(rows):
            for col, val in enumerate([r['request_number'], r['warehouse_name'] or "",
                                        r['employee'] or "", r['department'] or "", r['reason'] or ""]):
                self.iss_table.setItem(i, col, QTableWidgetItem(val))
            si = QTableWidgetItem(r['status'])
            si.setForeground(QColor(status_colors.get(r['status'], COLORS['text_primary'])))
            self.iss_table.setItem(i, 5, si)
            for col, val in [(6, r['appr_name'] or "—"), (7, r['date'])]:
                self.iss_table.setItem(i, col, QTableWidgetItem(val))

    def load_audit_log(self):
        rows = db.fetchall("SELECT * FROM audit_logs ORDER BY created_at DESC LIMIT 500")
        self.log_table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for col, val in enumerate([r['username'] or "", r['action'] or "",
                                        r['affected_table'] or "", str(r['affected_id'] or ""),
                                        r['details'] or "", (r['created_at'] or "")[:16]]):
                self.log_table.setItem(i, col, QTableWidgetItem(val))


# ─────────────────────────────────────────────────────────────────────────────
# INVENTORY AUDIT PAGE
# ─────────────────────────────────────────────────────────────────────────────
class AuditPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("📋  الجرد")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        hdr_row = QHBoxLayout()
        self.wh_combo = QComboBox()
        self.wh_combo.addItem("كل المخازن", None)
        for w in db.fetchall("SELECT id, warehouse_name FROM warehouses WHERE is_active=1"):
            self.wh_combo.addItem(w['warehouse_name'], w['id'])
        hdr_row.addWidget(QLabel("المخزن:")); hdr_row.addWidget(self.wh_combo)
        hdr_row.addStretch()
        load_btn = QPushButton("🔄  تحميل للجرد"); load_btn.clicked.connect(self.load_audit)
        save_btn = QPushButton("💾  حفظ الجرد"); save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save_audit)
        hdr_row.addWidget(load_btn); hdr_row.addWidget(save_btn)
        lay.addLayout(hdr_row)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["الكود", "الاسم", "المخزن", "كمية النظام",
                                               "العدد الفعلي", "الفرق", "قيمة الفرق"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        lay.addWidget(self.table)

        self.summary_lbl = QLabel("")
        self.summary_lbl.setStyleSheet(f"font-size: 13px; font-weight: 600; color: {COLORS['accent_orange']};")
        lay.addWidget(self.summary_lbl)
        self._rows = []

    def load_audit(self):
        wh_id = self.wh_combo.currentData()
        params = [wh_id] if wh_id else []
        where = "WHERE p.warehouse_id=?" if wh_id else ""
        self._rows = db.fetchall(f"""
            SELECT p.*, w.warehouse_name
            FROM products p
            LEFT JOIN warehouses w ON p.warehouse_id = w.id
            {where} ORDER BY p.name
        """, params)
        self.table.setRowCount(len(self._rows))
        self._physical_spins = []
        for i, r in enumerate(self._rows):
            self.table.setItem(i, 0, QTableWidgetItem(r['product_id']))
            self.table.setItem(i, 1, QTableWidgetItem(r['name']))
            self.table.setItem(i, 2, QTableWidgetItem(r['warehouse_name'] or ""))
            sys_qty = r['quantity'] or 0
            self.table.setItem(i, 3, QTableWidgetItem(f"{sys_qty:.0f}"))
            spin = QDoubleSpinBox()
            spin.setRange(0, 9999999)
            spin.setValue(sys_qty)
            spin.setDecimals(2)
            spin.valueChanged.connect(lambda v, row=i, sys=sys_qty, price=r['purchase_price'] or 0: self._update_diff(row, sys, price))
            self.table.setCellWidget(i, 4, spin)
            self._physical_spins.append(spin)
            diff = 0
            diff_item = QTableWidgetItem(f"{diff:.0f}")
            self.table.setItem(i, 5, diff_item)
            self.table.setItem(i, 6, QTableWidgetItem(f"AED {diff * (r['purchase_price'] or 0):.2f}"))
        self.summary_lbl.setText("")

    def _update_diff(self, row, sys_qty, price):
        if row < self.table.rowCount() and row < len(self._physical_spins):
            spin = self._physical_spins[row]
            diff = spin.value() - sys_qty
            diff_item = self.table.item(row, 5)
            val_item = self.table.item(row, 6)
            if diff_item:
                diff_item.setText(f"{diff:+.0f}")
                diff_item.setForeground(QColor(COLORS['accent_green'] if diff >= 0 else COLORS['accent_red']))
            if val_item:
                val_item.setText(f"AED {diff * price:+,.2f}")

    def save_audit(self):
        if not self._rows:
            QMessageBox.warning(self, "تنبيه", "حمّل البيانات أولاً")
            return
        audit_date = date.today().isoformat()
        uid = CURRENT_USER.get('id')
        wh_id = self.wh_combo.currentData()
        total_diff = 0
        for i, r in enumerate(self._rows):
            if i >= len(self._physical_spins):
                break
            sys_qty = r['quantity'] or 0
            phys_qty = self._physical_spins[i].value()
            diff = phys_qty - sys_qty
            total_diff += abs(diff)
            val_diff = diff * (r['purchase_price'] or 0)
            au_num = db.generate_audit_number()
            db.execute("""INSERT INTO inventory_audit
                (audit_number, warehouse_id, product_id, system_quantity, physical_quantity,
                 difference, value_difference, audit_date, audited_by)
                VALUES (?,?,?,?,?,?,?,?,?)""",
                (au_num, wh_id or r['warehouse_id'], r['id'], sys_qty, phys_qty, diff, val_diff, audit_date, uid))
            if diff != 0:
                db.execute("UPDATE products SET quantity=? WHERE id=?", (phys_qty, r['id']))
        log("حفظ الجرد", "inventory_audit", details=f"فروق: {total_diff:.0f}")
        self.summary_lbl.setText(f"✅ تم حفظ الجرد — إجمالي الفروق: {total_diff:.0f} وحدة")
        QMessageBox.information(self, "تم", "تم حفظ الجرد وتحديث الكميات")


# ─────────────────────────────────────────────────────────────────────────────
# USERS PAGE
# ─────────────────────────────────────────────────────────────────────────────
class UsersPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        hdr = QHBoxLayout()
        title = QLabel("👥  المستخدمون")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        hdr.addWidget(title); hdr.addStretch()
        if has_perm("all"):
            add_btn = QPushButton("➕  إضافة مستخدم")
            add_btn.setObjectName("btn_success")
            add_btn.clicked.connect(self.add_user)
            hdr.addWidget(add_btn)
        lay.addLayout(hdr)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["اسم المستخدم", "الاسم الكامل", "الصلاحية", "القسم", "البريد", "آخر دخول", "إجراءات"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        lay.addWidget(self.table)
        self.load_data()

    def load_data(self):
        rows = db.fetchall("SELECT * FROM users ORDER BY full_name")
        self.table.setRowCount(len(rows))
        role_colors = {"Admin": COLORS['accent_red'], "Warehouse Manager": COLORS['accent_orange'],
                       "Store Keeper": COLORS['accent_green'], "Auditor": COLORS['accent_purple'], "Viewer": COLORS['text_muted']}
        role_map_inv = {"Admin": "مدير نظام", "Warehouse Manager": "مدير مخازن", "Store Keeper": "أمين مخزن", "Auditor": "مدقق", "Viewer": "مشاهد"}

        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(r['username']))
            self.table.setItem(i, 1, QTableWidgetItem(r['full_name'] or ""))
            role_display = role_map_inv.get(r['role'], r['role'])
            role_item = QTableWidgetItem(role_display)
            role_item.setForeground(QColor(role_colors.get(r['role'], COLORS['text_primary'])))
            self.table.setItem(i, 2, role_item)
            self.table.setItem(i, 3, QTableWidgetItem(r['department'] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(r['email'] or ""))
            self.table.setItem(i, 5, QTableWidgetItem((r['last_login'] or "")[:16]))
            if has_perm("all"):
                bw = QWidget(); bl = QHBoxLayout(bw); bl.setContentsMargins(2,1,2,1); bl.setSpacing(3)
                eb = QPushButton("✏️"); eb.setFixedSize(28,26); eb.setObjectName("btn_flat")
                eb.clicked.connect(lambda _, row=dict(r): self.edit_user(row))
                bl.addWidget(eb)
                dbb = QPushButton("🗑"); dbb.setFixedSize(28,26); dbb.setObjectName("btn_danger")
                dbb.clicked.connect(lambda _, row=dict(r): self.delete_user(row))
                bl.addWidget(dbb)
                self.table.setCellWidget(i, 6, bw)

    def add_user(self):
        dlg = UserDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.load_data()

    def edit_user(self, row):
        dlg = UserDialog(self, row)
        if dlg.exec() == QDialog.Accepted:
            self.load_data()

    def delete_user(self, row):
        if row['username'] == 'admin':
            QMessageBox.critical(self, "خطأ", "لا يمكن حذف مدير النظام الرئيسي")
            return
        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد من حذف المستخدم: {row['username']}؟",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            db.execute("UPDATE users SET is_active=0 WHERE id=?", (row['id'],))
            log("حذف مستخدم", "users", row['id'], row['username'])
            self.load_data()


class UserDialog(QDialog):
    def __init__(self, parent=None, user=None):
        super().__init__(parent)
        self.user = user
        self.setWindowTitle("إضافة مستخدم" if not user else "تعديل مستخدم")
        self.setMinimumWidth(420)
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)
        title = QLabel("👤  " + ("إضافة مستخدم جديد" if not self.user else "تعديل بيانات المستخدم"))
        title.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']};")
        lay.addWidget(title)
        form = QFormLayout(); form.setSpacing(10)
        self.username_edit = QLineEdit()
        self.fullname_edit = QLineEdit()
        self.password_edit = QLineEdit(); self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("اتركه فارغاً للإبقاء على كلمة المرور الحالية" if self.user else "")
        self.role_combo = QComboBox()
        self.role_combo.addItems(["مدير نظام", "مدير مخازن", "أمين مخزن", "مدقق", "مشاهد"])
        self.email_edit = QLineEdit()
        self.dept_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.active_cb = QCheckBox("مستخدم نشط"); self.active_cb.setChecked(True)
        self.perms_edit = QTextEdit()
        self.perms_edit.setPlaceholderText("مثال: view,add,edit,delete")
        self.perms_edit.setFixedHeight(60)
        form.addRow("اسم المستخدم *:", self.username_edit)
        form.addRow("الاسم الكامل:", self.fullname_edit)
        form.addRow("كلمة المرور:", self.password_edit)
        form.addRow("الصلاحية:", self.role_combo)
        form.addRow("القسم:", self.dept_edit)
        form.addRow("البريد الإلكتروني:", self.email_edit)
        form.addRow("الهاتف:", self.phone_edit)
        form.addRow("صلاحيات إضافية (JSON):", self.perms_edit)
        form.addRow("", self.active_cb)
        lay.addLayout(form)
        if self.user:
            self.username_edit.setText(self.user['username'])
            self.fullname_edit.setText(self.user.get('full_name') or "")
            role_map_inv = {"Admin": "مدير نظام", "Warehouse Manager": "مدير مخازن", "Store Keeper": "أمين مخزن", "Auditor": "مدقق", "Viewer": "مشاهد"}
            idx = self.role_combo.findText(role_map_inv.get(self.user['role'], self.user['role']))
            if idx >= 0: self.role_combo.setCurrentIndex(idx)
            self.email_edit.setText(self.user.get('email') or "")
            self.dept_edit.setText(self.user.get('department') or "")
            self.phone_edit.setText(self.user.get('phone') or "")
            self.active_cb.setChecked(bool(self.user.get('is_active', 1)))
            self.perms_edit.setText(self.user.get('permissions') or "[]")
        btns = QHBoxLayout(); btns.addStretch()
        cb = QPushButton("إلغاء"); cb.setObjectName("btn_flat"); cb.clicked.connect(self.reject)
        sb = QPushButton("💾  حفظ"); sb.setObjectName("btn_success"); sb.clicked.connect(self.save)
        btns.addWidget(cb); btns.addWidget(sb)
        lay.addLayout(btns)

    def save(self):
        uname = self.username_edit.text().strip()
        if not uname:
            QMessageBox.warning(self, "تنبيه", "اسم المستخدم مطلوب")
            return
        pw = self.password_edit.text()
        role_map = {"مدير نظام": "Admin", "مدير مخازن": "Warehouse Manager", "أمين مخزن": "Store Keeper", "مدقق": "Auditor", "مشاهد": "Viewer"}
        role = role_map.get(self.role_combo.currentText(), "Viewer")

        perms = self.perms_edit.toPlainText().strip() or "[]"
        if self.user:
            if pw:
                pw_hash = hashlib.sha256(pw.encode()).hexdigest()
                db.execute("UPDATE users SET username=?, full_name=?, password=?, role=?, email=?, department=?, phone=?, is_active=?, permissions=? WHERE id=?",
                           (uname, self.fullname_edit.text().strip(), pw_hash, role,
                            self.email_edit.text().strip(), self.dept_edit.text().strip(),
                            self.phone_edit.text().strip(), int(self.active_cb.isChecked()), perms, self.user['id']))
            else:
                db.execute("UPDATE users SET username=?, full_name=?, role=?, email=?, department=?, phone=?, is_active=?, permissions=? WHERE id=?",
                           (uname, self.fullname_edit.text().strip(), role,
                            self.email_edit.text().strip(), self.dept_edit.text().strip(),
                            self.phone_edit.text().strip(), int(self.active_cb.isChecked()), perms, self.user['id']))
        else:
            if not pw:
                QMessageBox.warning(self, "تنبيه", "كلمة المرور مطلوبة")
                return
            pw_hash = hashlib.sha256(pw.encode()).hexdigest()
            db.execute("""INSERT INTO users (username, password, full_name, role, email, department, phone, is_active, permissions)
                VALUES (?,?,?,?,?,?,?,?,?)""",
                (uname, pw_hash, self.fullname_edit.text().strip(), role,
                 self.email_edit.text().strip(), self.dept_edit.text().strip(),
                 self.phone_edit.text().strip(), int(self.active_cb.isChecked()), perms))
        log("حفظ مستخدم", "users", details=uname)
        self.accept()


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS PAGE
# ─────────────────────────────────────────────────────────────────────────────
class SettingsPage(QWidget):
    def __init__(self, user, parent=None):
        super().__init__(parent)
        self.user = user
        self.setup_ui()

    def setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("⚙️  الإعدادات")
        title.setStyleSheet(f"font-size: 22px; font-weight: 800; color: {COLORS['text_primary']};")
        lay.addWidget(title)

        tabs = QTabWidget()

        # General Settings
        t1 = QWidget()
        f1 = QFormLayout(t1)
        f1.setContentsMargins(20, 16, 20, 16)
        f1.setSpacing(12)

        self.company_edit = QLineEdit(db.get_setting("company_name", COMPANY_NAME))
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["AED", "USD", "EUR", "SAR", "QAR"])
        self.currency_combo.setCurrentText(db.get_setting("currency", "AED"))
        self.tax_spin = QDoubleSpinBox()
        self.tax_spin.setRange(0, 100)
        self.tax_spin.setValue(float(db.get_setting("tax_rate", "0")))
        self.tax_spin.setVisible(False) # Hide as requested
        self.alarm_cb = QCheckBox("تفعيل تنبيهات المخزون")
        self.alarm_cb.setChecked(db.get_setting("alarm_enabled", "true") == "true")
        self.auto_backup_cb = QCheckBox("النسخ الاحتياطي التلقائي")
        self.auto_backup_cb.setChecked(db.get_setting("auto_backup", "true") == "true")
        self.notif_sound_cb = QCheckBox("تفعيل صوت التنبيه")
        self.notif_sound_cb.setChecked(db.get_setting("notif_sound", "true") == "true")

        f1.addRow("اسم الشركة:", self.company_edit)
        f1.addRow("العملة:", self.currency_combo)
        # f1.addRow("نسبة الضريبة (%):", self.tax_spin)
        f1.addRow("", self.alarm_cb)
        f1.addRow("", self.auto_backup_cb)
        f1.addRow("", self.notif_sound_cb)

        save_btn = QPushButton("💾  حفظ الإعدادات")
        save_btn.setObjectName("btn_success")
        save_btn.clicked.connect(self.save_settings)
        f1.addRow("", save_btn)
        tabs.addTab(t1, "إعدادات عامة")

        # Offers & Discounts
        if CURRENT_USER.get("role") in ["Admin", "Warehouse Manager"]:
            t_off = QWidget()
            t_off_l = QVBoxLayout(t_off)
            t_off_l.setContentsMargins(20, 16, 20, 16)

            ot = QLabel("🎟️  إدارة العروض والخصومات")
            ot.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COLORS['accent']};")
            t_off_l.addWidget(ot)

            oform = QFormLayout()
            self.offer_key = QLineEdit()
            self.offer_key.setPlaceholderText("مثال: RAMADAN2024")
            gen_key_btn = QPushButton("🔄 توليد كود")
            gen_key_btn.clicked.connect(self.generate_key)
            kh = QHBoxLayout(); kh.addWidget(self.offer_key); kh.addWidget(gen_key_btn)

            self.offer_type = QComboBox()
            self.offer_type.addItems(["نسبة مئوية (%)", "مبلغ ثابت"])
            self.offer_val = QDoubleSpinBox()
            self.offer_val.setRange(0, 1000000)

            self.offer_apply = QComboBox()
            self.offer_apply.addItems(["كل المنتجات", "فئة محددة", "منتج محدد"])

            oform.addRow("كود الخصم:", kh)
            oform.addRow("نوع الخصم:", self.offer_type)
            oform.addRow("قيمة الخصم:", self.offer_val)
            oform.addRow("ينطبق على:", self.offer_apply)

            add_off_btn = QPushButton("➕ إضافة العرض")
            add_off_btn.setObjectName("btn_success")
            add_off_btn.clicked.connect(self.save_offer)
            oform.addRow("", add_off_btn)
            t_off_l.addLayout(oform)

            self.offers_table = QTableWidget()
            self.offers_table.setColumnCount(5)
            self.offers_table.setHorizontalHeaderLabels(["الكود", "النوع", "القيمة", "التطبيق", "إجراءات"])
            t_off_l.addWidget(self.offers_table)
            tabs.addTab(t_off, "العروض والخصومات")
            self.load_offers()

        # Backup
        t2 = QWidget()
        t2l = QVBoxLayout(t2)
        t2l.setContentsMargins(20, 16, 20, 16)
        t2l.setSpacing(12)
        backup_grp = QGroupBox("النسخ الاحتياطي")
        bg = QVBoxLayout(backup_grp)
        backup_btn = QPushButton("💾  إنشاء نسخة احتياطية الآن")
        backup_btn.setObjectName("btn_success")
        backup_btn.clicked.connect(self.do_backup)
        restore_btn = QPushButton("🔄  استعادة نسخة احتياطية")
        restore_btn.setObjectName("btn_warning")
        restore_btn.clicked.connect(self.do_restore)
        bg.addWidget(backup_btn); bg.addWidget(restore_btn)
        t2l.addWidget(backup_grp)
        db_info = QGroupBox("معلومات قاعدة البيانات")
        di = QVBoxLayout(db_info)
        try:
            size = DB_PATH.stat().st_size / 1024
            di.addWidget(QLabel(f"مسار قاعدة البيانات: {DB_PATH}"))
            di.addWidget(QLabel(f"الحجم: {size:.1f} KB"))
        except:
            pass
        t2l.addWidget(db_info)
        t2l.addStretch()
        tabs.addTab(t2, "النسخ الاحتياطي")

        lay.addWidget(tabs)

    def generate_key(self):
        import random, string
        k = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
        self.offer_key.setText(k)

    def save_offer(self):
        key = self.offer_key.text().strip()
        if not key: return
        db.execute("INSERT INTO offers (offer_key, discount_type, discount_value, apply_to) VALUES (?,?,?,?)",
                   (key, self.offer_type.currentText(), self.offer_val.value(), self.offer_apply.currentText()))
        self.load_offers()
        self.offer_key.clear()

    def load_offers(self):
        rows = db.fetchall("SELECT * FROM offers WHERE is_active=1")
        self.offers_table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.offers_table.setItem(i, 0, QTableWidgetItem(r['offer_key']))
            self.offers_table.setItem(i, 1, QTableWidgetItem(r['discount_type']))
            self.offers_table.setItem(i, 2, QTableWidgetItem(str(r['discount_value'])))
            self.offers_table.setItem(i, 3, QTableWidgetItem(r['apply_to']))
            dbb = QPushButton("🗑")
            dbb.clicked.connect(lambda _, rid=r['id']: (db.execute("UPDATE offers SET is_active=0 WHERE id=?", (rid,)), self.load_offers()))
            self.offers_table.setCellWidget(i, 4, dbb)

    def save_settings(self):
        db.set_setting("company_name", self.company_edit.text().strip())
        db.set_setting("currency", self.currency_combo.currentText())
        db.set_setting("tax_rate", str(self.tax_spin.value()))
        db.set_setting("alarm_enabled", "true" if self.alarm_cb.isChecked() else "false")
        db.set_setting("auto_backup", "true" if self.auto_backup_cb.isChecked() else "false")
        db.set_setting("notif_sound", "true" if self.notif_sound_cb.isChecked() else "false")
        log("تعديل الإعدادات")
        QMessageBox.information(self, "تم", "تم حفظ الإعدادات بنجاح")

    def do_backup(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"backup_{ts}.db"
        try:
            shutil.copy2(str(DB_PATH), str(backup_path))
            log("نسخ احتياطي", details=str(backup_path))
            QMessageBox.information(self, "تم", f"تم إنشاء النسخة الاحتياطية:\n{backup_path}")
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل النسخ الاحتياطي:\n{e}")

    def do_restore(self):
        path, _ = QFileDialog.getOpenFileName(self, "اختر ملف النسخة الاحتياطية",
                                               str(BACKUP_DIR), "Database (*.db)")
        if not path:
            return
        reply = QMessageBox.question(self, "تأكيد",
            "سيتم استبدال قاعدة البيانات الحالية. هل أنت متأكد؟",
            QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                shutil.copy2(path, str(DB_PATH))
                QMessageBox.information(self, "تم", "تمت الاستعادة. أعد تشغيل التطبيق.")
            except Exception as e:
                QMessageBox.critical(self, "خطأ", str(e))


# ─────────────────────────────────────────────────────────────────────────────
# ALARM MANAGER
# ─────────────────────────────────────────────────────────────────────────────
class AlarmManager(QObject):
    warning_signal = Signal(list)
    critical_signal = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._notified_ids = set()

    def check(self):
        if db.get_setting("alarm_enabled", "true") != "true":
            return
        criticals = [dict(r) for r in db.check_critical_stock()]
        warnings = [dict(r) for r in db.check_low_stock() if r['id'] not in [c['id'] for c in criticals]]
        new_critical = [r for r in criticals if r['id'] not in self._notified_ids]
        new_warning = [r for r in warnings if r['id'] not in self._notified_ids]
        if new_critical:
            self.critical_signal.emit(new_critical)
            for r in new_critical:
                self._notified_ids.add(r['id'])
                insights = db.get_product_purchase_insights(r['id'])
                msg = f"الكمية: {r['quantity']:.0f}"
                if insights:
                    msg += f" | المورد المقترح: {insights['supplier_name']} | السعر: {insights['unit_price']:.2f}"
                db.execute("INSERT INTO notifications (title, message, type) VALUES (?,?,?)",
                           (f"🚨 مخزون حرج: {r['name']}", msg, "critical"))
        if new_warning:
            self.warning_signal.emit(new_warning)
            for r in new_warning:
                self._notified_ids.add(r['id'])
                insights = db.get_product_purchase_insights(r['id'])
                msg = f"الكمية: {r['quantity']:.0f}"
                if insights:
                    msg += f" | المورد المقترح: {insights['supplier_name']}"
                db.execute("INSERT INTO notifications (title, message, type) VALUES (?,?,?)",
                           (f"⚠️ تنبيه مخزون: {r['name']}", msg, "warning"))


# ─────────────────────────────────────────────────────────────────────────────
# MAIN WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self, user):
        super().__init__()
        set_current_user(user)
        self.user = user
        self.setWindowTitle(f"{COMPANY_NAME} — نظام إدارة المخازن المؤسسي")
        self.setMinimumSize(1280, 780)
        self.resize(1440, 900)

        if LOGO_PATH.exists():
            self.setWindowIcon(QIcon(str(LOGO_PATH)))

        self.alarm_mgr = AlarmManager(self)
        self.alarm_mgr.warning_signal.connect(self.show_warning_alarm)
        self.alarm_mgr.critical_signal.connect(self.show_critical_alarm)

        self.alarm_timer = QTimer(self)
        self.alarm_timer.timeout.connect(self.alarm_mgr.check)
        self.alarm_timer.start(60000)

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_clock)
        self.clock_timer.start(1000)

        self.setup_ui()
        self.navigate_to(0)

        QTimer.singleShot(2000, self.alarm_mgr.check)
        if db.get_setting("auto_backup", "true") == "true":
            QTimer.singleShot(5000, self._auto_backup)

    def _auto_backup(self):
        try:
            ts = datetime.now().strftime("%Y%m%d")
            bp = BACKUP_DIR / f"auto_backup_{ts}.db"
            if not bp.exists():
                shutil.copy2(str(DB_PATH), str(bp))
        except:
            pass

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ── Sidebar (RTL: right side) ──
        sidebar = QFrame()
        sidebar.setFixedWidth(230)
        sidebar.setStyleSheet(f"QFrame {{ background: {COLORS['bg_nav']}; border-right: 1px solid {COLORS['border2']}; }}")
        sidebar_lay = QVBoxLayout(sidebar)
        sidebar_lay.setContentsMargins(0, 0, 0, 0)
        sidebar_lay.setSpacing(0)

        # Logo area
        logo_area = QFrame()
        logo_area.setFixedHeight(90)
        logo_area.setStyleSheet(f"background: {COLORS['bg_dark']}; border-bottom: 1px solid {COLORS['border2']};")
        la_lay = QHBoxLayout(logo_area)
        la_lay.setContentsMargins(12, 10, 12, 10)
        la_lay.setSpacing(10)
        if LOGO_PATH.exists():
            ll = QLabel()
            ll.setPixmap(QPixmap(str(LOGO_PATH)).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            la_lay.addWidget(ll)
        lt = QVBoxLayout(); lt.setSpacing(2)
        cl = QLabel("AMS")
        cl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 18px; font-weight: 900; letter-spacing: 2px;")
        sl = QLabel("نظام إدارة المخازن")
        sl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 9px;")
        lt.addWidget(cl); lt.addWidget(sl)
        la_lay.addLayout(lt)
        sidebar_lay.addWidget(logo_area)

        # User info
        ua = QFrame()
        ua.setStyleSheet(f"background: {COLORS['bg_card']}; border-bottom: 1px solid {COLORS['border2']};")
        ua.setFixedHeight(68)
        ul = QHBoxLayout(ua)
        ul.setContentsMargins(12, 8, 12, 8)
        ul.setSpacing(10)
        name_initial = self.user.get('full_name', 'م')
        if not name_initial: name_initial = 'م'
        avatar = QLabel(name_initial[0].upper())
        avatar.setFixedSize(40, 40)
        avatar.setAlignment(Qt.AlignCenter)
        avatar.setStyleSheet(f"background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 {COLORS['accent']}, stop:1 {COLORS['accent2']}); border-radius: 20px; color: white; font-size: 16px; font-weight: 800;")
        ui2 = QVBoxLayout(); ui2.setSpacing(2)
        nl = QLabel(self.user.get('full_name', self.user.get('username', 'مستخدم')))
        nl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 12px; font-weight: 600;")
        role_map_inv = {"Admin": "مدير نظام", "Warehouse Manager": "مدير مخازن", "Store Keeper": "أمين مخزن", "Auditor": "مدقق", "Viewer": "مشاهد"}
        rl = QLabel(role_map_inv.get(self.user.get('role'), "مشاهد"))
        rl.setStyleSheet(f"color: {COLORS['accent']}; font-size: 10px;")
        ui2.addWidget(nl); ui2.addWidget(rl)
        ul.addWidget(avatar); ul.addLayout(ui2)
        sidebar_lay.addWidget(ua)

        nav_label = QLabel("القائمة الرئيسية")
        nav_label.setContentsMargins(16, 12, 0, 4)
        nav_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 10px; font-weight: 700; letter-spacing: 1px;")
        sidebar_lay.addWidget(nav_label)

        nav_items = [
            ("🏠", AR['dashboard']),
            ("📦", AR['products']),
            ("🏭", AR['warehouses']),
            ("🏢", AR['suppliers']),
            ("📥", AR['stock_in']),
            ("📤", AR['stock_out']),
            ("🔄", AR['transfers']),
            ("📊", AR['reports']),
            ("📋", AR['audit']),
            ("👥", AR['users']),
            ("⚙️", AR['settings']),
        ]

        self.nav_btn_group = QButtonGroup(self)
        self.nav_btn_group.setExclusive(True)

        for i, (icon, label) in enumerate(nav_items):
            btn = SidebarButton(icon, label)
            self.nav_btn_group.addButton(btn, i)
            btn.clicked.connect(lambda _, idx=i: self.navigate_to(idx))
            sidebar_lay.addWidget(btn)

        sidebar_lay.addStretch()

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background: {COLORS['border2']};")
        sidebar_lay.addWidget(sep)

        logout_btn = QPushButton(f"  🚪  {AR['logout']}")
        logout_btn.setStyleSheet(f"""
            QPushButton {{ background: transparent; color: {COLORS['accent_red']}; border: none;
                text-align: right; padding: 12px 16px; font-size: 13px; font-weight: 600; }}
            QPushButton:hover {{ background: {COLORS['warning_bg']}; }}
        """)
        logout_btn.setCursor(Qt.PointingHandCursor)
        logout_btn.clicked.connect(self.logout)
        sidebar_lay.addWidget(logout_btn)

        # ── Right/Main area ──
        right_area = QVBoxLayout()
        right_area.setContentsMargins(0, 0, 0, 0)
        right_area.setSpacing(0)

        # Top bar
        topbar = QFrame()
        topbar.setFixedHeight(56)
        topbar.setStyleSheet(f"QFrame {{ background: {COLORS['bg_nav']}; border-bottom: 1px solid {COLORS['border2']}; }}")
        tb_lay = QHBoxLayout(topbar)
        tb_lay.setContentsMargins(20, 0, 20, 0)
        tb_lay.setSpacing(12)

        self.page_title_lbl = QLabel(AR['dashboard'])
        self.page_title_lbl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 16px; font-weight: 700;")
        tb_lay.addWidget(self.page_title_lbl)
        tb_lay.addStretch()

        global_search = QLineEdit()
        global_search.setPlaceholderText("🔍  بحث سريع...")
        global_search.setFixedWidth(220)
        global_search.setFixedHeight(34)
        global_search.returnPressed.connect(lambda: self.global_search(global_search.text()))
        tb_lay.addWidget(global_search)

        self.clock_lbl = QLabel()
        self.clock_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; min-width: 160px;")
        self.update_clock()
        tb_lay.addWidget(self.clock_lbl)

        self.notif_badge = NotificationBadge()
        tb_lay.addWidget(self.notif_badge)

        quick_add_btn = QPushButton("➕ إضافة سريعة")
        quick_add_btn.setFixedSize(120, 34)
        quick_add_btn.setObjectName("btn_success")
        quick_add_btn.clicked.connect(lambda: self.navigate_to(4))
        tb_lay.addWidget(quick_add_btn)

        right_area.addWidget(topbar)

        # Pages
        self.stack = QStackedWidget()
        self.stack.setStyleSheet(f"background: {COLORS['bg_dark']};")

        self.dashboard_page = DashboardPage()
        self.products_page = ProductsPage()
        self.products_page.stock_alert.connect(self.handle_stock_alert)
        self.warehouses_page = WarehousesPage()
        self.suppliers_page = SuppliersPage()
        self.stock_in_page = StockInPage()
        self.stock_in_page.stock_changed.connect(self.on_stock_changed)
        self.stock_out_page = StockOutPage()
        self.stock_out_page.stock_changed.connect(self.on_stock_changed)
        self.transfers_page = TransfersPage()
        self.transfers_page.stock_changed.connect(self.on_stock_changed)
        self.reports_page = ReportsPage()
        self.audit_page = AuditPage()
        self.users_page = UsersPage()
        self.settings_page = SettingsPage(self.user)

        for page in [self.dashboard_page, self.products_page, self.warehouses_page,
                     self.suppliers_page, self.stock_in_page, self.stock_out_page,
                     self.transfers_page, self.reports_page, self.audit_page,
                     self.users_page, self.settings_page]:
            self.stack.addWidget(page)

        scroll_area = QScrollArea()
        scroll_area.setWidget(self.stack)
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet(f"background: {COLORS['bg_dark']}; border: none;")
        right_area.addWidget(scroll_area)

        right_widget = QWidget()
        right_widget.setLayout(right_area)

        # RTL: sidebar on right, content on left
        main_layout.addWidget(sidebar)
        main_layout.addWidget(right_widget)

        self.statusBar().showMessage(
            f"  ✅  مرحباً، {self.user.get('full_name', '')}  |  جهة: {COMPANY_NAME}")
        self.statusBar().setFixedHeight(28)
        self.update_notification_count()

    def navigate_to(self, idx):
        self.stack.setCurrentIndex(idx)
        titles = [AR['dashboard'], AR['products'], AR['warehouses'], AR['suppliers'],
                  AR['stock_in'], AR['stock_out'], AR['transfers'], AR['reports'],
                  AR['audit'], AR['users'], AR['settings']]
        self.page_title_lbl.setText(titles[idx] if idx < len(titles) else "")
        btn = self.nav_btn_group.button(idx)
        if btn:
            btn.setChecked(True)
        refresh_map = {
            0: lambda: self.dashboard_page.refresh_data(),
            1: lambda: self.products_page.load_products(),
            2: lambda: self.warehouses_page.load_data(),
            3: lambda: self.suppliers_page.load_data(),
        }
        if idx in refresh_map:
            refresh_map[idx]()

    def on_stock_changed(self):
        self.dashboard_page.refresh_data()
        self.products_page.load_products()
        self.alarm_mgr.check()
        self.update_notification_count()

    def handle_stock_alert(self, warnings, criticals):
        self.update_notification_count()

    def show_warning_alarm(self, items):
        self.update_notification_count()
        if items:
            msg = QMessageBox(self)
            msg.setWindowTitle("⚠️ تنبيه مخزون منخفض")
            msg.setIcon(QMessageBox.Warning)
            il = "\n".join([f"• {r['name']}: {r['quantity']:.0f} وحدة (التنبيه: {r['alert_quantity']:.0f})" for r in items[:5]])
            if len(items) > 5:
                il += f"\n... و {len(items) - 5} منتجات أخرى"
            msg.setText(f"تنبيه مخزون منخفض!\n\n{il}")
            msg.exec()

    def show_critical_alarm(self, items):
        self.update_notification_count()
        if items:
            msg = QMessageBox(self)
            msg.setWindowTitle("🚨 تنبيه حرج — مخزون تحت الحد الأدنى")
            msg.setIcon(QMessageBox.Critical)
            il = "\n".join([f"• {r['name']}: {r['quantity']:.0f} وحدة (الحد الأدنى: {r['min_quantity']:.0f})" for r in items[:5]])
            msg.setText(f"مخزون حرج! يرجى اتخاذ إجراء فوري:\n\n{il}")
            msg.exec()

    def update_notification_count(self):
        count = db.fetchone("SELECT COUNT(*) as c FROM notifications WHERE is_read=0")
        self.notif_badge.set_count(count['c'] if count else 0)

    def update_clock(self):
        now = datetime.now()
        h = now.hour
        suffix = "صباحاً" if h < 12 else "مساءً"
        h12 = h if h <= 12 else h - 12
        if h12 == 0: h12 = 12
        self.clock_lbl.setText(f"📅 {now.strftime('%Y/%m/%d')}  🕐 {h12}:{now.strftime('%M:%S')} {suffix}")

    def global_search(self, text):
        if not text.strip():
            return
        self.navigate_to(1)
        self.products_page.search_edit.setText(text)

    def logout(self):
        reply = QMessageBox.question(self, "تسجيل الخروج",
            "هل تريد تسجيل الخروج؟", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            log("تسجيل خروج", "users", self.user.get('id'))
            self.close()
            self.login_again()

    def login_again(self):
        login = LoginWindow()
        login.login_successful.connect(lambda u: self._restart_with_user(u))
        login.exec()

    def _restart_with_user(self, user):
        new_window = MainWindow(user)
        new_window.show()
        self._new_window = new_window


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("AMS Enterprise Warehouse ERP")
    app.setOrganizationName(COMPANY_NAME)
    app.setStyleSheet(STYLESHEET)

    # Enable RTL globally
    app.setLayoutDirection(Qt.RightToLeft)

    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # Splash
    splash = SplashScreen()
    splash.show()
    QApplication.processEvents()

    steps = [
        (10, "جارٍ تحميل الإعدادات..."),
        (25, "تهيئة قاعدة البيانات..."),
        (45, "تحديث الجداول..."),
        (60, "فحص المخزون..."),
        (80, "تجهيز الواجهة..."),
        (95, "اكتمال التحضير..."),
        (100, "جاهز!"),
    ]
    for pct, msg in steps:
        splash.set_progress(pct, msg)
        QApplication.processEvents()
        time.sleep(0.15)
    time.sleep(0.3)
    splash.close()

    login = LoginWindow()
    main_window_holder = [None]

    def on_login(user):
        main_window_holder[0] = MainWindow(user)
        main_window_holder[0].show()

    login.login_successful.connect(on_login)
    result = login.exec()

    if result != QDialog.Accepted:
        sys.exit(0)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()