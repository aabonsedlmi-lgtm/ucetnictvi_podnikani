#!/usr/bin/env python3
from __future__ import annotations

import base64
from collections import defaultdict
from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile
import traceback
from typing import Any
import unicodedata
import zlib
import zipfile

from flask import Flask, flash, g, jsonify, redirect, render_template_string, request, send_file, url_for

from invoice_backend import (
    DB_FILE,
    add_customer,
    add_expense_category,
    add_expense,
    add_item,
    add_issuer,
    compute_totals,
    connect_db,
    copy_file_to_drive,
    clear_all_business_data,
    delete_expense,
    delete_expense_category_if_unused,
    delete_customer_if_unused,
    delete_issuer_if_unused,
    delete_invoice,
    ensure_recurring_expenses,
    expense_is_review_complete,
    export_data_json,
    export_invoice_html,
    export_invoice_pdf,
    fetch_ares_subject,
    get_app_setting,
    get_invoice_detail,
    import_data_json,
    import_invoices_from_excel,
    init_db,
    list_customers,
    list_expense_categories,
    list_expenses,
    list_invoices,
    list_issuers,
    mark_paid,
    monthly_expense_total,
    new_invoice,
    parse_date,
    propagate_recurring_amounts,
    rebuild_recurring_future_expenses,
    replace_invoice_items,
    recurring_period_label,
    set_app_setting,
    split_invoice_description,
    invoice_item_is_note,
    get_expense,
    sync_recurring_expense_series,
    update_invoice,
    update_customer,
    update_expense,
    update_expense_category,
    update_issuer,
    year_month_overview,
    yearly_expense_overview,
    yearly_overview,
)

EXPORT_DIR = Path("exports")
INCOME_DOCS_DIR = Path("Prijem")
EXPENSE_DOCS_DIR = Path("Vydej")
TEMP_EXPENSE_UPLOAD_DIR = Path("uploads") / "expenses" / "_temp"
ERROR_LOG_FILE = Path("app_errors.log")

app = Flask(__name__)
app.config["SECRET_KEY"] = "fakturace-studio-local"
_DB_INIT_DONE = False

BASE_TEMPLATE = """
<!doctype html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }} | Fakturace Studio</title>
  <style>
    :root{--bg:#081311;--panel:#0f1e1b;--panel-2:#142825;--line:#223a35;--text:#e7f3ef;--muted:#95ada7;--accent:#14b8a6;--accent-soft:#113a35;--danger:#f87171;--danger-soft:#3a1c20;--shadow:0 24px 60px rgba(0,0,0,.34)}
    *{box-sizing:border-box} body{margin:0;font-family:"Segoe UI",Tahoma,sans-serif;color:var(--text);background:radial-gradient(circle at top left,#0f3a34 0,transparent 28%),radial-gradient(circle at top right,#1b2b3d 0,transparent 22%),linear-gradient(180deg,#081311 0%,#0b1715 100%)} a{color:inherit;text-decoration:none}
    .shell{display:grid;grid-template-columns:290px 1fr;min-height:100vh;gap:20px;padding:20px}
    .sidebar,.hero,.card,.table-card,.content-card{background:rgba(15,30,27,.92);border:1px solid var(--line);box-shadow:var(--shadow);border-radius:24px}
    .sidebar{padding:24px 18px;position:sticky;top:20px;align-self:start}.brand-kicker{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:#5eead4;font-weight:700}.brand-title{margin:10px 0 6px;font-size:28px;font-weight:800;line-height:1.05}.brand-copy{margin:0 0 20px;color:var(--muted);font-size:14px;line-height:1.5}.active-issuer-box{margin:0 0 18px;padding:12px 14px;border-radius:18px;background:var(--panel-2);border:1px solid var(--line)}.active-issuer-box .label{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-weight:700}.active-issuer-box .value{margin-top:6px;font-size:15px;font-weight:800;line-height:1.3;color:var(--text)}.sidebar-cta{display:block;margin-top:18px;padding:14px 16px;border-radius:18px;background:var(--accent-soft);color:#7df3e4;font-weight:800;text-align:center;border:1px solid #1e4d47}.nav{display:grid;gap:10px;margin-top:18px}.nav a{padding:14px 16px;border-radius:18px;color:var(--text);font-weight:700}.nav a:hover{background:var(--panel-2)}.nav a.active{background:linear-gradient(135deg,var(--accent),#0f766e);color:#041311;box-shadow:0 14px 32px rgba(20,184,166,.22)}
    .content{display:grid;gap:18px;align-content:start}.hero{padding:24px 28px;display:flex;justify-content:space-between;gap:20px;align-items:flex-start}.hero h1{margin:0;font-size:34px;line-height:1.05;letter-spacing:-.03em}.hero p{margin:10px 0 0;color:var(--muted);max-width:720px;line-height:1.6}.hero-badge{padding:10px 14px;border-radius:999px;background:var(--accent-soft);color:var(--accent);font-weight:800;white-space:nowrap}
    .flash-stack{display:grid;gap:10px}.flash{padding:14px 16px;border-radius:18px;border:1px solid var(--line);background:var(--panel-2);font-weight:600}.flash.error{background:var(--danger-soft);border-color:#7f1d1d;color:#fecaca}
    .grid-3,.grid-2,.form-grid,.form-grid-3,.invoice-top{display:grid;gap:18px}.grid-3{grid-template-columns:repeat(3,minmax(0,1fr))}.grid-2,.invoice-top,.form-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.form-grid-3{grid-template-columns:repeat(3,minmax(0,1fr))}
    .card,.table-card,.content-card{padding:20px}.metric-label{color:var(--muted);font-weight:700;font-size:13px;text-transform:uppercase;letter-spacing:.08em}.metric-value{margin-top:8px;font-size:34px;font-weight:800;letter-spacing:-.04em}.metric-sub,.section-copy,.muted{color:var(--muted)}.metric-sub{margin-top:8px;font-size:14px}.section-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:14px}.section-head h2{margin:0;font-size:20px;letter-spacing:-.02em}
    table{width:100%;border-collapse:collapse;font-size:14px} th,td{padding:14px 12px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top} th{color:var(--muted);text-transform:uppercase;letter-spacing:.08em;font-size:12px} tr:last-child td{border-bottom:none}
    .pill{display:inline-flex;padding:7px 12px;border-radius:999px;font-size:12px;font-weight:800;background:var(--panel-2)}.pill.paid,.pill.ok{background:#123327;color:#86efac}.pill.draft{background:#102b3b;color:#7dd3fc}.pill.sent{background:#3b2f10;color:#fcd34d}.pill.overdue{background:#3b181d;color:#fca5a5}.pill.pending{background:#3b3014;color:#fde68a}.pill.review{background:#182f49;color:#93c5fd}.pill-button{border:none;cursor:pointer}.pill-button.sent:hover,.pill-button.overdue:hover{filter:brightness(1.12);transform:translateY(-1px)}
    .toolbar{display:flex;flex-wrap:wrap;gap:12px;align-items:center}.button,.button-secondary,.button-danger{border:none;cursor:pointer;border-radius:16px;padding:12px 16px;font-weight:800;font-size:14px}.button{background:linear-gradient(135deg,var(--accent),#0f766e);color:#041311;box-shadow:0 14px 32px rgba(20,184,166,.22)}.button-secondary{background:var(--panel-2);color:var(--text);border:1px solid var(--line)}.button-danger{background:var(--danger-soft);color:#fecaca;border:1px solid #7f1d1d}.button-link{padding:10px 14px;border-radius:14px;background:var(--panel-2);font-weight:700;display:inline-flex;align-items:center;justify-content:center;border:1px solid var(--line)}
    form{margin:0}.stack,.service-table,.field{display:grid;gap:8px}.field label{font-size:13px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em} input,select,textarea{width:100%;border:1px solid var(--line);border-radius:16px;padding:13px 14px;background:#0b1715;color:var(--text);font:inherit} textarea{min-height:110px;resize:vertical} textarea[data-autogrow]{min-height:140px;overflow:hidden;white-space:pre-wrap}.service-row{display:grid;grid-template-columns:2.5fr .8fr 1fr .8fr auto;gap:10px;align-items:end}.service-row.no-vat{grid-template-columns:2.5fr .8fr 1fr auto}.detail-total{font-size:30px;font-weight:900;letter-spacing:-.04em}.lookup-row{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:end}.lookup-status{font-size:12px;color:var(--muted);min-height:18px}
    @media (max-width:1100px){.shell{grid-template-columns:1fr}.sidebar{position:static}.grid-3,.grid-2,.form-grid,.form-grid-3,.invoice-top{grid-template-columns:1fr}.service-row{grid-template-columns:1fr}.hero{flex-direction:column}}
  </style>
</head>
<body>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand-kicker">Studio</div>
      <div class="brand-title">Fakturace<br>ve webu</div>
      <p class="brand-copy">Úplně nová podoba aplikace v Pythonu. Jedno rozhraní pro faktury, firmy, odběratele a exporty.</p>
      <div class="active-issuer-box">
        <div class="label">Aktivní firma</div>
        <div class="value">{{ active_issuer_name }}</div>
      </div>
      <a class="sidebar-cta" href="{{ url_for('new_invoice_page') }}">Vystavit novou fakturu</a>
      <nav class="nav">
        <a href="{{ url_for('dashboard') }}" class="{{ 'active' if active == 'dashboard' else '' }}">Dashboard</a>
        <a href="{{ url_for('invoices_page') }}" class="{{ 'active' if active == 'invoices' else '' }}">Příjmy</a>
        <a href="{{ url_for('expenses_page') }}" class="{{ 'active' if active == 'expenses' else '' }}">Výdaje</a>
        <a href="{{ url_for('year_report_page') }}" class="{{ 'active' if active == 'year-report' else '' }}">Roční přehled</a>
        <a href="{{ url_for('guides_page') }}" class="{{ 'active' if active == 'guides' else '' }}">Návody</a>
        <a href="{{ url_for('settings_page') }}" class="{{ 'active' if active == 'settings' else '' }}">Nastavení</a>
      </nav>
    </aside>
    <main class="content">
      <section class="hero"><div><h1>{{ title }}</h1><p>{{ subtitle }}</p></div></section>
      {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}<div class="flash-stack">{% for category, message in messages %}<div class="flash {{ category }}">{{ message }}</div>{% endfor %}</div>{% endif %}{% endwith %}
      {{ body|safe }}
    </main>
  </div>
  <script>
    function autoGrowTextareas() {
      document.querySelectorAll('textarea[data-autogrow]').forEach(function(el) {
        const resize = function() {
          el.style.height = 'auto';
          el.style.height = Math.max(el.scrollHeight, 140) + 'px';
        };
        if (!el.dataset.autogrowBound) {
          el.addEventListener('input', resize);
          el.dataset.autogrowBound = '1';
        }
        resize();
      });
    }
    function startBackupDownload(link) {
      if (!link || link.dataset.loading === '1') return false;
      link.dataset.loading = '1';
      link.style.pointerEvents = 'none';
      link.style.opacity = '0.78';
      const idle = link.querySelector('.backup-button-idle');
      const busy = link.querySelector('.backup-button-busy');
      if (idle) idle.style.display = 'none';
      if (busy) busy.style.display = 'inline';
      const status = document.getElementById('backupStatusMessage');
      if (status) status.style.display = 'block';
      return true;
    }
    document.addEventListener('DOMContentLoaded', autoGrowTextareas);
  </script>
  {% if ares_script %}<script>{{ ares_script|safe }}</script>{% endif %}
</body>
</html>
"""

ARES_LOOKUP_SCRIPT = """
async function lookupAres(kind) {
  const map = {
    customer: {
      ico: document.getElementById('customer-ico'),
      name: document.getElementById('customer-name'),
      dic: document.getElementById('customer-dic'),
      address: document.getElementById('customer-address'),
      status: document.getElementById('customer-ares-status'),
      vat: null
    },
    issuer: {
      ico: document.getElementById('issuer-ico'),
      name: document.getElementById('issuer-name'),
      dic: document.getElementById('issuer-dic'),
      address: document.getElementById('issuer-address'),
      status: document.getElementById('issuer-ares-status'),
      vat: document.getElementById('issuer-vat-payer')
    }
  };

  const cfg = map[kind];
  if (!cfg || !cfg.ico) return;

  const ico = (cfg.ico.value || '').trim();
  if (!ico) {
    if (cfg.status) cfg.status.textContent = 'Zadej IČO.';
    return;
  }

  if (cfg.status) cfg.status.textContent = 'Načítám ARES...';

  try {
    const response = await fetch(`/api/ares?ico=${encodeURIComponent(ico)}`);
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || 'ARES nevrátil data.');
    }

    const data = payload.data || {};
    if (cfg.name && data.name) cfg.name.value = data.name;
    if (cfg.dic && data.dic) cfg.dic.value = data.dic;
    if (cfg.address && data.address) cfg.address.value = data.address;
    if (cfg.ico && data.ico) cfg.ico.value = data.ico;
    if (cfg.vat && typeof data.vat_payer === 'boolean') cfg.vat.checked = data.vat_payer;
    if (cfg.status) cfg.status.textContent = 'Údaje byly doplněny z ARES.';
  } catch (error) {
    if (cfg.status) cfg.status.textContent = error.message || 'Načtení z ARES se nepodařilo.';
  }
}
"""


def open_db():
    global _DB_INIT_DONE
    conn = getattr(g, "_db_conn", None)
    if conn is None:
        conn = connect_db(Path(DB_FILE))
        g._db_conn = conn
    _migrate_document_directories(conn)
    if not _DB_INIT_DONE:
        try:
            init_db(conn)
        except Exception as exc:
            if "locked" not in str(exc).lower():
                raise
        _DB_INIT_DONE = True
    return conn


@app.teardown_appcontext
def close_db(exception=None):
    conn = getattr(g, "_db_conn", None)
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass
        g._db_conn = None


def try_ensure_recurring_expenses(conn) -> None:
    try:
        ensure_recurring_expenses(conn)
    except Exception as exc:
        if "locked" not in str(exc).lower():
            raise


def expense_category_options(conn: sqlite3.Connection) -> list[str]:
    rows = list_expense_categories(conn)
    names = [str(row["name"] or "").strip() for row in rows if str(row["name"] or "").strip()]
    if not names:
        return ["Ostatní"]
    if "Ostatní" not in names:
        names.append("Ostatní")
    return names


def status_pill(status: str) -> str:
    label_map = {"draft": "Rozpracovaná", "sent": "Odeslaná", "paid": "Zaplacená", "overdue": "Po splatnosti", "pending": "Čeká na úhradu", "review": "Dolož fakturu", "ok": "OK"}
    return f"<span class='pill {status}'>{label_map.get(status, status)}</span>"


def invoice_status_cell(status: str, invoice_id: int) -> str:
    normalized = str(status or "").strip().lower()
    if normalized in {"draft", "sent", "overdue"}:
        label = {"draft": "Rozpracovaná", "sent": "Odeslaná", "overdue": "Po splatnosti"}.get(normalized, normalized)
        return (
            f"<form method='post' class='invoice-paid-form' data-invoice-id='{invoice_id}' action='{url_for('mark_invoice_paid_from_list_page', invoice_id=invoice_id)}' "
            f"onsubmit=\"return confirm('Označit fakturu jako zaplacenou?');\">"
            f"<button class='pill pill-button {normalized}' type='submit' title='Kliknutím označíš jako zaplacenou'>{label}</button>"
            f"</form>"
        )
    return status_pill(normalized)


def format_date_cz(value: Any, default: str = "-") -> str:
    text = str(value or "").strip()
    if not text:
        return default
    try:
        return parse_date(text).strftime("%d.%m.%Y")
    except ValueError:
        return text


def parse_amount_text(value: Any) -> float:
    text = str(value or "").strip()
    clean = re.sub(r"[^0-9,.\- ]", "", text).replace(" ", "").replace(",", ".")
    if not clean:
        return 0.0
    try:
        return float(clean)
    except ValueError:
        return 0.0


def parse_int_field(value: Any, default: int = 0) -> int:
    text = str(value or "").strip()
    if not text:
        return default
    try:
        return int(float(text.replace(",", ".")))
    except (TypeError, ValueError):
        return default


def parse_float_field(value: Any, default: float = 0.0) -> float:
    text = str(value or "").strip()
    if not text:
        return default
    try:
        return float(text.replace(",", "."))
    except (TypeError, ValueError):
        return default


def log_app_error(context: str, exc: Exception) -> None:
    try:
        ERROR_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with ERROR_LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] {context}: {type(exc).__name__}: {exc}\n")
            handle.write(traceback.format_exc())
            handle.write("\n")
    except Exception:
        pass


def _bundle_document_dirs() -> list[Path]:
    return [INCOME_DOCS_DIR, EXPENSE_DOCS_DIR]


def _create_backup_bundle(conn: sqlite3.Connection, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workspace_root = Path.cwd().resolve()
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        backup_json = temp_root / "backup.json"
        export_data_json(conn, backup_json)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(backup_json, "backup.json")
            for doc_dir in _bundle_document_dirs():
                if not doc_dir.exists():
                    continue
                for file_path in doc_dir.rglob("*"):
                    if not file_path.is_file():
                        continue
                    resolved_path = file_path.resolve()
                    try:
                        archive_name = str(resolved_path.relative_to(workspace_root))
                    except ValueError:
                        archive_name = str(file_path)
                    archive.write(resolved_path, archive_name)
    return output_path


def _copy_tree_contents(source_dir: Path, target_dir: Path) -> None:
    if not source_dir.exists():
        return
    for path in source_dir.rglob("*"):
        if path.is_dir():
            continue
        relative = path.relative_to(source_dir)
        destination = target_dir / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, destination)


def _migrate_document_directories(conn: sqlite3.Connection | None = None) -> None:
    if getattr(_migrate_document_directories, "_done", False):
        return

    legacy_map = {
        Path("Příjem"): INCOME_DOCS_DIR,
        Path("Výdej"): EXPENSE_DOCS_DIR,
    }

    for legacy_dir, new_dir in legacy_map.items():
        try:
            if legacy_dir.resolve() == new_dir.resolve():
                continue
        except Exception:
            pass
        if legacy_dir.exists():
            new_dir.mkdir(parents=True, exist_ok=True)
            _copy_tree_contents(legacy_dir, new_dir)
            for path in sorted(legacy_dir.rglob("*"), reverse=True):
                if path.is_file():
                    path.unlink(missing_ok=True)
                elif path.is_dir():
                    try:
                        path.rmdir()
                    except OSError:
                        pass
            try:
                legacy_dir.rmdir()
            except OSError:
                pass

    if conn is not None:
        replacements = [
            ("\\Příjem\\", "\\Prijem\\"),
            ("/Příjem/", "/Prijem/"),
            ("\\Výdej\\", "\\Vydej\\"),
            ("/Výdej/", "/Vydej/"),
            ("Příjem\\", "Prijem\\"),
            ("Výdej\\", "Vydej\\"),
            ("Příjem/", "Prijem/"),
            ("Výdej/", "Vydej/"),
        ]
        for old, new in replacements:
            conn.execute(
                "UPDATE expenses SET attachment_path = REPLACE(attachment_path, ?, ?) WHERE attachment_path LIKE ?",
                (old, new, f"%{old}%"),
            )
        conn.commit()

    _migrate_document_directories._done = True


def _clear_business_files() -> dict[str, int]:
    deleted = {
        "income_docs": 0,
        "expense_docs": 0,
        "temp_uploads": 0,
        "exports": 0,
    }

    def remove_tree_files(root: Path, counter_key: str) -> None:
        if not root.exists():
            return
        for path in sorted(root.rglob("*"), reverse=True):
            if path.is_file():
                path.unlink(missing_ok=True)
                deleted[counter_key] += 1
            elif path.is_dir():
                try:
                    path.rmdir()
                except OSError:
                    pass

    remove_tree_files(INCOME_DOCS_DIR, "income_docs")
    remove_tree_files(EXPENSE_DOCS_DIR, "expense_docs")
    remove_tree_files(TEMP_EXPENSE_UPLOAD_DIR, "temp_uploads")

    remove_tree_files(EXPORT_DIR, "exports")

    return deleted


def _restore_backup_bundle(conn: sqlite3.Connection, input_path: Path) -> tuple[int, int, int, int]:
    suffix = input_path.suffix.lower()
    if suffix == ".json":
        return import_data_json(conn, input_path)

    if suffix != ".zip":
        raise ValueError("Podporovaná je JSON nebo ZIP záloha.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        with zipfile.ZipFile(input_path, "r") as archive:
            archive.extractall(temp_root)

        backup_json = temp_root / "backup.json"
        if not backup_json.exists():
            raise ValueError("ZIP záloha neobsahuje soubor backup.json.")

        raw = backup_json.read_text(encoding="utf-8")
        payload = json.loads(raw)
        workspace_root = Path.cwd().resolve()
        for expense in payload.get("expenses", []):
            attachment_path = str(expense.get("attachment_path", "") or "").strip()
            if not attachment_path:
                continue
            attachment_candidate = Path(attachment_path)
            relative_candidate: Path | None = None
            if attachment_candidate.is_absolute():
                try:
                    relative_candidate = attachment_candidate.resolve().relative_to(workspace_root)
                except Exception:
                    relative_candidate = None
            else:
                relative_candidate = attachment_candidate
            if relative_candidate:
                bundled_file = temp_root / relative_candidate
                if bundled_file.exists():
                    expense["attachment_path"] = str((workspace_root / relative_candidate).resolve())

        normalized_json = temp_root / "backup_normalized.json"
        normalized_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        restored = import_data_json(conn, normalized_json)

        for doc_dir in _bundle_document_dirs():
            source_dir = temp_root / doc_dir
            target_dir = workspace_root / doc_dir
            _copy_tree_contents(source_dir, target_dir)

        return restored


def render_page(title: str, subtitle: str, body_template: str, active: str, **context: Any) -> str:
    active_aliases = {
        "invoice-new": "invoices",
        "issuers": "settings",
        "expense-categories": "expenses",
        "customers": "invoices",
        "data-tools": "settings",
    }
    active = active_aliases.get(active, active)
    active_issuer_name = context.get("active_issuer_name")
    if not active_issuer_name:
        try:
            conn = open_db()
            issuers = list_issuers(conn)
            active_issuer = next((row for row in issuers if int(row["is_default"] or 0) == 1), issuers[0] if issuers else None)
            active_issuer_name = str(active_issuer["company_name"] or "-") if active_issuer else "-"
        except Exception:
            active_issuer_name = "-"
    body = render_template_string(body_template, **context)
    return render_template_string(BASE_TEMPLATE, title=title, subtitle=subtitle, body=body, active=active, active_issuer_name=active_issuer_name, ares_script=context.get("ares_script", ""))


def expense_status_options() -> list[tuple[str, str]]:
    return [
        ("review", "Dolož fakturu"),
        ("ok", "OK"),
    ]


def normalize_expense_status(status: Any, attachment_verified: bool = False, price_confirmed: bool = False) -> str:
    value = str(status or "").strip().lower()
    if value == "review":
        return "review"
    if value in {"ok", "paid", "pending", "overdue"}:
        return "ok"
    if attachment_verified and price_confirmed:
        return "ok"
    return "review"


def recurring_period_options() -> list[tuple[str, str]]:
    return [
        ("monthly", "Měsíčně"),
        ("quarterly", "Kvartálně"),
        ("half_yearly", "Půlročně"),
        ("yearly", "Ročně"),
    ]


def coerce_checkbox(value: Any) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "on", "yes"}


def review_flags_from_form(form, fallback_attachment: str = "") -> tuple[bool, bool, bool]:
    attachment_path = str(form.get("attachment_path", fallback_attachment) or "").strip()
    attachment_verified = bool(attachment_path)
    price_confirmed = coerce_checkbox(form.get("price_confirmed")) or attachment_verified
    price_manual_override = coerce_checkbox(form.get("price_manual_override"))
    return attachment_verified, price_confirmed, price_manual_override


def expense_review_meta(row: sqlite3.Row | dict[str, Any]) -> dict[str, Any]:
    attachment = str(row["attachment_path"] if hasattr(row, "keys") and "attachment_path" in row.keys() else row.get("attachment_path", "") or "").strip()
    recurring = int(row["recurring"] if hasattr(row, "keys") and "recurring" in row.keys() else row.get("recurring", 0) or 0)
    period = str(row["recurring_period"] if hasattr(row, "keys") and "recurring_period" in row.keys() else row.get("recurring_period", "") or "")
    price_confirmed = int(row["price_confirmed"] if hasattr(row, "keys") and "price_confirmed" in row.keys() else row.get("price_confirmed", 0) or 0)
    attachment_verified = int(row["attachment_verified"] if hasattr(row, "keys") and "attachment_verified" in row.keys() else row.get("attachment_verified", 0) or 0)
    complete = expense_is_review_complete(row)
    return {
        "attachment_path": attachment,
        "attachment_label": Path(attachment).name if attachment else "Chybí PDF doklad",
        "attachment_verified": attachment_verified,
        "price_confirmed": price_confirmed,
        "review_complete": complete,
        "recurring": recurring,
        "period_label": recurring_period_label(period),
    }


def maybe_copy_export(conn, file_path: Path) -> Path | None:
    folder = get_app_setting(conn, "google_drive_folder", "").strip()
    auto = get_app_setting(conn, "google_drive_auto_export", "0") == "1"
    if not folder or not auto:
        return None
    return copy_file_to_drive(file_path, Path(folder))


def _sanitize_filename(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).name)
    return safe or "expense.pdf"


def _store_uploaded_pdf(upload) -> Path:
    TEMP_EXPENSE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = TEMP_EXPENSE_UPLOAD_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{_sanitize_filename(upload.filename)}"
    upload.save(target)
    return target


def _relocate_expense_pdf(source_path: str | Path, expense_date: date) -> Path:
    source = Path(source_path)
    target_dir = EXPENSE_DOCS_DIR / str(expense_date.year)
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / _sanitize_filename(source.name)
    if target.exists():
        target = target_dir / f"{target.stem}_{datetime.now().strftime('%H%M%S')}{target.suffix}"
    if source.resolve() != target.resolve():
        target.write_bytes(source.read_bytes())
        try:
            source.unlink()
        except OSError:
            pass
    return target


def _relocate_income_import_pdf(source_path: str | Path, issue_date: date, invoice_number: str) -> Path:
    source = Path(source_path)
    target_dir = INCOME_DOCS_DIR / str(issue_date.year)
    target_dir.mkdir(parents=True, exist_ok=True)
    safe_number = _sanitize_filename(invoice_number or source.stem or "importovana_faktura")
    target = target_dir / f"import_{safe_number}{source.suffix or '.pdf'}"
    if target.exists():
        target = target_dir / f"import_{safe_number}_{datetime.now().strftime('%H%M%S')}{source.suffix or '.pdf'}"
    if source.resolve() != target.resolve():
        target.write_bytes(source.read_bytes())
        try:
            source.unlink()
        except OSError:
            pass
    return target


def _invoice_export_path(invoice_number: str, service_safe: str, issue_date_value: Any, fmt: str) -> Path:
    issue_date = parse_date(str(issue_date_value), default=date.today())
    target_dir = INCOME_DOCS_DIR / str(issue_date.year)
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"{invoice_number}_{service_safe}.{fmt}"


def _decode_pdf_literal_string(raw: bytes) -> str:
    result = bytearray()
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == 92 and i + 1 < len(raw):  # backslash
            i += 1
            esc = raw[i]
            mapping = {
                ord("n"): b"\n",
                ord("r"): b"\r",
                ord("t"): b"\t",
                ord("b"): b"\b",
                ord("f"): b"\f",
                ord("("): b"(",
                ord(")"): b")",
                ord("\\"): b"\\",
            }
            if esc in mapping:
                result.extend(mapping[esc])
            elif 48 <= esc <= 55:
                oct_digits = bytes([esc])
                for _ in range(2):
                    if i + 1 < len(raw) and 48 <= raw[i + 1] <= 55:
                        i += 1
                        oct_digits += bytes([raw[i]])
                    else:
                        break
                result.append(int(oct_digits, 8))
            else:
                result.append(esc)
        else:
            result.append(ch)
        i += 1
    return result.decode("utf-8", errors="ignore") or result.decode("latin1", errors="ignore")


def _extract_pdf_stream_filters(header: bytes) -> list[str]:
    match = re.search(rb"/Filter\s*(\[(.*?)\]|/\w+)", header, re.S)
    if not match:
        return []
    source = match.group(1)
    return [item.decode("ascii", errors="ignore").lstrip("/") for item in re.findall(rb"/([A-Za-z0-9]+)", source)]


def _decode_pdf_stream(chunk: bytes, filters: list[str]) -> bytes:
    decoded = chunk.strip()
    for filter_name in filters:
        if filter_name == "ASCII85Decode":
            if not decoded.endswith(b"~>"):
                decoded += b"~>"
            decoded = base64.a85decode(decoded, adobe=True)
        elif filter_name == "FlateDecode":
            decoded = zlib.decompress(decoded)
    return decoded


def _clean_extracted_pdf_lines(lines: list[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for line in lines:
        text = re.sub(r"\s+", " ", str(line or "").replace("\x00", " ")).strip()
        if not text:
            continue
        printable_ratio = sum(ch.isprintable() for ch in text) / max(len(text), 1)
        if printable_ratio < 0.85:
            continue
        if not re.search(r"[A-Za-zÀ-ž0-9]", text):
            continue
        if len(text) > 220:
            continue
        lower = text.lower()
        if lower in seen:
            continue
        seen.add(lower)
        cleaned.append(text)
    return cleaned


def _looks_like_bad_pdf_text(text: str) -> bool:
    if not text.strip():
        return True
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return True
    suspicious = 0
    for line in lines:
        tokens = line.split()
        single_char_tokens = [token for token in tokens if len(token) == 1]
        if re.search(r"(?:\b\S\b\s+){6,}", line):
            suspicious += 1
        if len(tokens) >= 8 and len(single_char_tokens) / max(len(tokens), 1) >= 0.6:
            suspicious += 1
        if line.count("(") >= 6 or line.count(")") >= 6:
            suspicious += 1
        if re.search(r"[^\x09\x0A\x0D\x20-\x7EÀ-ž]", line):
            suspicious += 1
        if sum(not (ch.isalnum() or ch.isspace() or ch in ".,:;/%+-()[]") for ch in line) > max(6, len(line) // 4):
            suspicious += 1
        if len(line) >= 20 and sum(ch.isupper() for ch in line) >= 8 and len(single_char_tokens) >= 6:
            suspicious += 1
    return suspicious >= max(2, len(lines) // 4)


def _ocr_pdf_text(pdf_path: Path) -> str:
    try:
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
    except Exception:
        return ""

    ocr = RapidOCR()
    doc = pdfium.PdfDocument(str(pdf_path))
    ocr_lines: list[str] = []
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            bitmap = page.render(scale=2.0)
            image = bitmap.to_numpy()
            result, _ = ocr(image)
            if not result:
                continue
            for item in result:
                if not item or len(item) < 3:
                    continue
                text = str(item[1] or "").strip()
                confidence = float(item[2] or 0)
                if text and confidence >= 0.35:
                    ocr_lines.append(text)
    finally:
        doc.close()

    cleaned = _clean_extracted_pdf_lines(ocr_lines)
    return "\n".join(cleaned)


def _extract_pdf_text(pdf_path: Path) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        extracted_pages: list[str] = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            if page_text.strip():
                extracted_pages.append(page_text)
        combined = "\n".join(extracted_pages)
        cleaned_pages = _clean_extracted_pdf_lines(combined.splitlines())
        if cleaned_pages and not _looks_like_bad_pdf_text("\n".join(cleaned_pages)):
            return "\n".join(cleaned_pages)
    except Exception:
        pass

    raw = pdf_path.read_bytes()
    parts: list[str] = []
    for match in re.finditer(rb"<<(.*?)>>\s*stream\r?\n(.*?)\r?\nendstream", raw, re.S):
        header = match.group(1)
        chunk = match.group(2)
        filters = _extract_pdf_stream_filters(header)
        try:
            decoded = _decode_pdf_stream(chunk, filters)
        except Exception:
            decoded = chunk

        for bit in re.findall(rb"\((.*?)(?<!\\)\)", decoded, re.S):
            text = _decode_pdf_literal_string(bit)
            if text:
                parts.append(text)

    cleaned = _clean_extracted_pdf_lines(parts)
    joined = "\n".join(cleaned)
    if cleaned and not _looks_like_bad_pdf_text(joined):
        return joined

    ocr_text = _ocr_pdf_text(pdf_path)
    if ocr_text:
        return ocr_text
    return "Text se z tohoto PDF nepodařilo spolehlivě načíst. PDF pravděpodobně obsahuje hlavně grafiku nebo vložené fonty. Údaje doplň ručně a zkontroluj je před uložením."


def _guess_expense_from_pdf_text(text: str) -> dict[str, str]:
    if text.startswith("Text se z tohoto PDF nepodařilo spolehlivě načíst."):
        return {
            "title": "PDF doklad",
            "supplier_name": "",
            "document_number": "",
            "expense_date": date.today().isoformat(),
            "due_date": date.today().isoformat(),
            "amount": "0",
            "amount_with_vat": "0",
            "amount_without_vat": "0",
            "vat_rate": "0",
            "currency": "CZK",
            "document_type": "Přijatá faktura",
            "status": "review",
            "category": "Ostatní",
            "payment_method": "Převodem",
            "project_code": "",
            "note": "PDF se nepodařilo automaticky přečíst. Údaje doplň ručně a zkontroluj je.",
        }

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    labels_upper = {
        "VYSTAVENO:",
        "DATUM SPLATNOSTI:",
        "DATUM USKUTECNENI ZDANITELNEHO PLNENI:",
        "DATUM USKUTECNENI ZDANITELNEHO PLNENI:",
        "DODAVATEL:",
        "ODBERATEL - SIDLO:",
        "ODBĚRATEL - SÍDLO:",
        "POSTOVNI ADRESA:",
        "POŠTOVNÍ ADRESA:",
        "MISTO URCENI:",
        "MÍSTO URČENÍ:",
        "IBAN:",
        "BANKOVNI UCET:",
        "BANKOVNÍ ÚČET:",
        "VAR. SYM.:",
        "KONST. SYM.:",
        "SPEC. SYM.:",
        "FORMA UHRADY:",
        "FORMA ÚHRADY:",
        "ZPUSOB DOPRAVY:",
        "ZPŮSOB DOPRAVY:",
        "CISLO SMLOUVY:",
        "ČÍSLO SMLOUVY:",
        "OBJEDNAVKA:",
        "OBJEDNÁVKA:",
        "ZAKAZKA:",
        "ZAKÁZKA:",
        "BIC:",
        "BANKA:",
        "WWW:",
        "E-MAIL:",
        "MOBIL:",
        "FAX:",
        "TELEFON:",
        "QR PLATBA",
        "FAKTURUJEME VAM NASLEDUJICI POLOZKY",
        "FAKTURUJEME VÁM NÁSLEDUJÍCÍ POLOŽKY",
        "OZNACENI DODAVKY",
        "OZNAČENÍ DODÁVKY",
        "CELKEM:",
        "CELKEM K UHRADE",
        "CELKEM K ÚHRADĚ",
        "ZALOHY",
        "ZÁLOHY",
        "ZBYVA UHRADIT [KC]",
        "ZBÝVÁ UHRADIT [KČ]",
        "RAZITKO A PODPIS",
        "RAZÍTKO A PODPIS",
        "REKAPITULACE DPH V KC",
        "REKAPITULACE DPH V KČ",
        "REGISTRACE:",
    }


def _guess_invoice_from_pdf_text(text: str) -> dict[str, str]:
    if text.startswith("Text se z tohoto PDF nepodařilo spolehlivě načíst."):
        return {
            "invoice_number": "",
            "issue_date": date.today().isoformat(),
            "due_date": date.today().isoformat(),
            "currency": "CZK",
            "issuer_name": "",
            "issuer_ico": "",
            "customer_name": "",
            "customer_email": "",
            "customer_phone": "",
            "customer_ico": "",
            "customer_dic": "",
            "customer_address": "",
            "item_description": "Služby dle PDF",
            "quantity": "1",
            "price": "0",
            "vat_rate": "21",
            "note": "PDF se nepodařilo automaticky přečíst. Údaje doplň ručně a zkontroluj je.",
        }

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    def norm_label(value: str) -> str:
        raw = unicodedata.normalize("NFKD", str(value or ""))
        ascii_text = raw.encode("ascii", "ignore").decode("ascii")
        ascii_text = ascii_text.upper().replace("?", " ")
        ascii_text = re.sub(r"[^A-Z0-9]+", " ", ascii_text)
        return re.sub(r"\s+", " ", ascii_text).strip()

    labels_upper = {norm_label(item) for item in [
        "Dodavatel", "Odběratel", "Odběratel - sídlo", "Poštovní adresa", "Místo určení",
        "Datum vystavení", "Vystaveno", "Datum splatnosti", "Datum zdan. plnění",
        "Označení dodávky", "Množství", "Jedn.cena", "Cena za MJ", "Cena za M.J.",
        "Celkem", "Celkem k úhradě", "Zbývá uhradit", "Rekapitulace DPH", "Razítko a podpis",
    ]}

    def parse_any_date(value: str) -> str:
        raw_value = str(value or "").strip()
        if not raw_value:
            return date.today().isoformat()
        match = re.search(r"(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})", raw_value)
        if not match:
            return date.today().isoformat()
        normalized = match.group(1)
        for fmt_in in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(normalized, fmt_in).date().isoformat()
            except ValueError:
                continue
        return date.today().isoformat()

    def normalize_amount(value: str) -> str:
        text_value = str(value or "").strip().replace("\u00a0", " ")
        text_value = re.sub(r"[^\d,.\s-]", "", text_value).strip()
        if not text_value:
            return "0"
        if "," not in text_value and "." not in text_value:
            return text_value.replace(" ", "") + ".00"
        return text_value.replace(" ", "").replace(",", ".")

    def block_after_label(*labels: str, max_lines: int = 7) -> list[str]:
        target_labels = {norm_label(label) for label in labels}
        for index, line in enumerate(lines):
            if norm_label(line) in target_labels:
                block: list[str] = []
                for look_ahead in range(1, max_lines + 1):
                    if index + look_ahead >= len(lines):
                        break
                    candidate = lines[index + look_ahead].strip()
                    candidate_label = norm_label(candidate)
                    if candidate_label in labels_upper:
                        break
                    if re.search(r"^(IČO|ICO|DIČ|DIC|IBAN|BIC|WWW|E-mail|Mobil|Fax|Telefon|Banka|Var\. sym\.|Konst\. sym\.)", candidate, re.I):
                        block.append(candidate)
                        continue
                    block.append(candidate)
                if block:
                    return block
        return []

    def first_line_after_label(*labels: str) -> str:
        block = block_after_label(*labels, max_lines=4)
        return block[0] if block else ""

    def top_dates() -> list[str]:
        found: list[str] = []
        for line in lines[:16]:
            for match in re.findall(r"(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})", line):
                if match not in found:
                    found.append(match)
        return found

    def find_amount(*patterns: str) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, re.I)
            if match:
                return normalize_amount(match.group(1))
        amounts = re.findall(r"(\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.]\d{2})?)", text)
        return normalize_amount(amounts[-1]) if amounts else "0"

    def parse_company_block(block: list[str]) -> dict[str, str]:
        name = ""
        address_lines: list[str] = []
        ico = ""
        dic = ""
        email = ""
        phone = ""
        for line in block:
            stripped = line.strip()
            if not stripped:
                continue
            if re.match(r"^(IČO|ICO)\s*[: ]", stripped, re.I):
                ico = re.sub(r"\D", "", stripped)
                continue
            if re.match(r"^(DIČ|DIC)\s*[: ]", stripped, re.I):
                dic = stripped.split(":", 1)[-1].strip() if ":" in stripped else re.sub(r"^(DIČ|DIC)\s*", "", stripped, flags=re.I).strip()
                continue
            if re.match(r"^(E-mail|Email)\s*[: ]", stripped, re.I):
                email = stripped.split(":", 1)[-1].strip() if ":" in stripped else stripped
                continue
            if re.match(r"^(Telefon|Tel\.|Mobil)\s*[: ]", stripped, re.I):
                phone = stripped.split(":", 1)[-1].strip() if ":" in stripped else stripped
                continue
            if not name:
                name = stripped
            else:
                address_lines.append(stripped)
        return {
            "name": name[:160],
            "address": "\n".join(address_lines[:4]).strip(),
            "ico": re.sub(r"\D", "", ico)[:8],
            "dic": dic[:32],
            "email": email[:160],
            "phone": phone[:64],
        }

    invoice_number = ""
    for pattern in [
        r"Faktura\s*-\s*daňový doklad\s*([A-Z0-9./_-]+)",
        r"(?:Číslo faktury|Cislo faktury|Faktura)\s*[:#]?\s*([A-Z0-9./_-]{4,})",
    ]:
        match = re.search(pattern, text, re.I)
        if match:
            invoice_number = match.group(1).strip()
            break

    header_dates = top_dates()
    issue_date = parse_any_date(header_dates[0]) if len(header_dates) >= 1 else parse_any_date(first_line_after_label("Vystaveno", "Datum vystavení"))
    due_date = parse_any_date(header_dates[1]) if len(header_dates) >= 2 else parse_any_date(first_line_after_label("Datum splatnosti", "Den splatnosti"))

    issuer_block = parse_company_block(block_after_label("Dodavatel", "Dodavatel:"))
    customer_block = parse_company_block(block_after_label("Odběratel", "Odběratel - sídlo", "Odběratel:"))

    amount = find_amount(
        r"Celkem k úhradě\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
        r"Zbývá uhradit(?:\s*\[[^\]]+\])?\s*([0-9\s]+(?:[,.][0-9]{2})?)",
        r"Celkem\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
    )
    vat_rate_match = re.search(r"(?:Sazba DPH|DPH)\s*[: ]\s*([0-9]{1,2}(?:[.,][0-9]+)?)", text, re.I)
    vat_rate = vat_rate_match.group(1).replace(",", ".") if vat_rate_match else "21"

    item_description = ""
    for index, line in enumerate(lines):
        normalized = norm_label(line)
        if normalized in {norm_label("Označení dodávky"), norm_label("Popis"), norm_label("OZNACENI DODAVKY")}:
            for look_ahead in range(1, 6):
                if index + look_ahead >= len(lines):
                    break
                candidate = lines[index + look_ahead].strip()
                if not candidate or norm_label(candidate) in labels_upper:
                    continue
                if re.fullmatch(r"[0-9., ]+", candidate):
                    continue
                item_description = candidate
                break
            if item_description:
                break
    if not item_description:
        item_description = "Služby dle PDF"

    quantity = "1"
    for line in lines:
        if item_description and item_description in line:
            continue
        if re.fullmatch(r"\d+(?:[,.]\d+)?", line.strip()):
            quantity = line.strip().replace(",", ".")
            break

    return {
        "invoice_number": invoice_number[:80],
        "issue_date": issue_date,
        "due_date": due_date,
        "currency": "CZK",
        "issuer_name": issuer_block["name"],
        "issuer_ico": issuer_block["ico"],
        "customer_name": customer_block["name"],
        "customer_email": customer_block["email"],
        "customer_phone": customer_block["phone"],
        "customer_ico": customer_block["ico"],
        "customer_dic": customer_block["dic"],
        "customer_address": customer_block["address"],
        "item_description": item_description[:240],
        "quantity": quantity,
        "price": amount,
        "vat_rate": vat_rate,
        "note": "Načteno z PDF - zkontroluj údaje před uložením.",
    }


def _guess_invoice_items_from_pdf_text(text: str, fallback: dict[str, str]) -> list[dict[str, str]]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return [{
            "description": fallback.get("item_description", "Služby dle PDF"),
            "quantity": fallback.get("quantity", "1"),
            "price": fallback.get("price", "0"),
            "vat_rate": fallback.get("vat_rate", "21"),
        }]

    def norm_label(value: str) -> str:
        raw = unicodedata.normalize("NFKD", str(value or ""))
        ascii_text = raw.encode("ascii", "ignore").decode("ascii")
        ascii_text = ascii_text.upper().replace("?", " ")
        ascii_text = re.sub(r"[^A-Z0-9]+", " ", ascii_text)
        return re.sub(r"\s+", " ", ascii_text).strip()

    start_index = -1
    for index, line in enumerate(lines):
        normalized = norm_label(line)
        if "OZNACENI DODAVKY" in normalized or normalized == "POPIS":
            start_index = index
            break

    if start_index < 0:
        return [{
            "description": fallback.get("item_description", "Služby dle PDF"),
            "quantity": fallback.get("quantity", "1"),
            "price": fallback.get("price", "0"),
            "vat_rate": fallback.get("vat_rate", "21"),
        }]

    stop_labels = {
        "CELKEM", "CELKEM K UHRADE", "ZBYVA UHRADIT", "ZALOHY", "RAZITKO A PODPIS",
        "REKAPITULACE DPH", "REGISTRACE",
    }
    item_lines: list[str] = []
    for line in lines[start_index + 1:]:
        normalized = norm_label(line)
        if not normalized:
            continue
        if normalized in stop_labels or any(normalized.startswith(label) for label in stop_labels):
            break
        item_lines.append(line)

    items: list[dict[str, str]] = []
    pending_description: list[str] = []

    for line in item_lines:
        if re.search(r"\b\d+(?:[,.]\d+)?\b", line) and re.search(r"\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.]\d{2})", line):
            amounts = re.findall(r"\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.]\d{2})", line)
            qty_match = re.search(r"\b(\d+(?:[,.]\d+)?)\b", line)
            vat_match = re.search(r"\b([0-9]{1,2}(?:[,.][0-9]+)?)\s*%?\b", line)
            price_value = amounts[-1] if amounts else fallback.get("price", "0")
            quantity_value = qty_match.group(1).replace(",", ".") if qty_match else "1"
            vat_value = fallback.get("vat_rate", "21")
            if vat_match and vat_match.group(1) not in {quantity_value, price_value}:
                vat_value = vat_match.group(1).replace(",", ".")
            description = " ".join(pending_description).strip() or fallback.get("item_description", "Služby dle PDF")
            items.append(
                {
                    "description": description[:240],
                    "quantity": quantity_value,
                    "price": parse_amount_text(price_value).__format__(".2f"),
                    "vat_rate": vat_value,
                }
            )
            pending_description = []
            continue
        pending_description.append(line)

    if not items:
        return [{
            "description": fallback.get("item_description", "Služby dle PDF"),
            "quantity": fallback.get("quantity", "1"),
            "price": fallback.get("price", "0"),
            "vat_rate": fallback.get("vat_rate", "21"),
        }]

    return items


def _find_matching_issuer_id(issuers: list[sqlite3.Row], guessed: dict[str, str]) -> int:
    guessed_ico = re.sub(r"\D", "", guessed.get("issuer_ico", "") or "")
    guessed_name = str(guessed.get("issuer_name", "") or "").strip().lower()
    for row in issuers:
        row_ico = re.sub(r"\D", "", str(row["ico"] or ""))
        if guessed_ico and row_ico == guessed_ico:
            return int(row["id"])
    for row in issuers:
        if guessed_name and guessed_name == str(row["company_name"] or "").strip().lower():
            return int(row["id"])
    for row in issuers:
        if int(row["is_default"] or 0):
            return int(row["id"])
    return int(issuers[0]["id"]) if issuers else 0


def _find_or_create_customer_for_invoice_pdf(
    conn: sqlite3.Connection,
    *,
    name: str,
    email: str,
    phone: str,
    ico: str,
    dic: str,
    address: str,
) -> int:
    clean_ico = re.sub(r"\D", "", ico or "")
    if clean_ico:
        row = conn.execute("SELECT id, name, email, phone, ico, dic, address FROM customers WHERE replace(coalesce(ico,''), ' ', '') = ? LIMIT 1", (clean_ico,)).fetchone()
        if row:
            update_customer(
                conn,
                int(row["id"]),
                name or str(row["name"] or ""),
                email or str(row["email"] or ""),
                address or str(row["address"] or ""),
                phone or str(row["phone"] or ""),
                clean_ico or str(row["ico"] or ""),
                dic or str(row["dic"] or ""),
            )
            return int(row["id"])
    row = conn.execute(
        "SELECT id, name, email, phone, ico, dic, address FROM customers WHERE lower(name)=lower(?) AND lower(coalesce(address,''))=lower(?) LIMIT 1",
        (name.strip(), address.strip()),
    ).fetchone()
    if row:
        update_customer(
            conn,
            int(row["id"]),
            name or str(row["name"] or ""),
            email or str(row["email"] or ""),
            address or str(row["address"] or ""),
            phone or str(row["phone"] or ""),
            clean_ico or str(row["ico"] or ""),
            dic or str(row["dic"] or ""),
        )
        return int(row["id"])
    return add_customer(conn, name, email, address, phone, clean_ico, dic)

    def norm_label(value: str) -> str:
        raw = unicodedata.normalize("NFKD", str(value or ""))
        ascii_text = raw.encode("ascii", "ignore").decode("ascii")
        ascii_text = ascii_text.upper().replace("?", " ")
        ascii_text = re.sub(r"[^A-Z0-9]+", " ", ascii_text)
        return re.sub(r"\s+", " ", ascii_text).strip()

    def block_after_label(*labels: str, max_lines: int = 6) -> list[str]:
        label_set = {norm_label(label) for label in labels}
        for index, line in enumerate(lines):
            if norm_label(line) in label_set:
                block: list[str] = []
                for look_ahead in range(1, max_lines + 1):
                    if index + look_ahead >= len(lines):
                        break
                    candidate = lines[index + look_ahead].strip()
                    candidate_label = norm_label(candidate)
                    if candidate_label in labels_upper:
                        break
                    if re.match(r"^(IČO|ICO|DIČ|DIC|IBAN|BIC|WWW|E-mail|Mobil|Fax|Telefon|Banka|Var\. sym\.|Konst\. sym\.)", candidate, re.I):
                        continue
                    if re.search(r"(?:Datum|Vystaveno|splatnosti)", candidate, re.I):
                        break
                    block.append(candidate)
                if block:
                    return block
        return []

    def block_before_label(*labels: str, max_lines: int = 6) -> list[str]:
        label_set = {norm_label(label) for label in labels}
        for index, line in enumerate(lines):
            if norm_label(line) in label_set:
                block: list[str] = []
                for look_back in range(1, max_lines + 1):
                    prev_index = index - look_back
                    if prev_index < 0:
                        break
                    candidate = lines[prev_index].strip()
                    candidate_label = norm_label(candidate)
                    if candidate_label in labels_upper:
                        break
                    if re.match(r"^(IČO|ICO|DIČ|DIC|IBAN|BIC|WWW|E-mail|Mobil|Fax|Telefon|Banka|Var\. sym\.|Konst\. sym\.)", candidate, re.I):
                        continue
                    block.append(candidate)
                if block:
                    block.reverse()
                    return block
        return []

    def first_line_after_label(*labels: str) -> str:
        block = block_after_label(*labels, max_lines=4)
        return block[0] if block else ""

    def first_line_before_label(*labels: str) -> str:
        block = block_before_label(*labels, max_lines=4)
        return block[-1] if block else ""

    def top_dates() -> list[str]:
        found: list[str] = []
        for line in lines[:12]:
            for match in re.findall(r"(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})", line):
                if match not in found:
                    found.append(match)
        return found

    def normalize_amount(value: str) -> str:
        text_value = str(value or "").strip().replace("\u00a0", " ")
        text_value = re.sub(r"[^\d,.\s]", "", text_value).strip()
        if not text_value:
            return "0"
        if "," not in text_value and "." not in text_value:
            return text_value.replace(" ", "") + ".00"
        return text_value.replace(" ", "").replace(",", ".")

    def find_amount(*patterns: str) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, re.I)
            if match:
                return normalize_amount(match.group(1))
        candidates = re.findall(r"(\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.]\d{2})?)", text)
        return normalize_amount(candidates[-1]) if candidates else "0"

    def find_amount_near_labels(*labels: str) -> str:
        labels_lower = [norm_label(label) for label in labels]
        for index, line in enumerate(lines):
            current_lower = norm_label(line)
            if any(current_lower.startswith(label) or current_lower == label for label in labels_lower):
                current_match = re.search(r"(\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.][0-9]{2})?)", line)
                if current_match:
                    return normalize_amount(current_match.group(1))
                for look_back in range(1, 3):
                    if index - look_back >= 0:
                        prev_line = lines[index - look_back]
                        prev_match = re.search(r"(\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.][0-9]{2})?)", prev_line)
                        if prev_match:
                            return normalize_amount(prev_match.group(1))
                for look_ahead in range(1, 3):
                    if index + look_ahead >= len(lines):
                        break
                    next_line = lines[index + look_ahead]
                    next_match = re.search(r"(\d{1,3}(?:[ \u00a0]\d{3})*(?:[,.][0-9]{2})?)", next_line)
                    if next_match:
                        return normalize_amount(next_match.group(1))
        return "0"

    def find_standalone_amount_after_label(*labels: str) -> str:
        labels_lower = [norm_label(label) for label in labels]
        for index, line in enumerate(lines):
            current_lower = norm_label(line)
            if any(current_lower.startswith(label) or current_lower == label for label in labels_lower):
                for look_ahead in range(1, 4):
                    if index + look_ahead >= len(lines):
                        break
                    candidate = lines[index + look_ahead].strip()
                    if re.fullmatch(r"[0-9\s]+(?:[,.][0-9]{2})?", candidate):
                        return normalize_amount(candidate)
                    match = re.search(r"([0-9]{1,3}(?:[ \u00a0][0-9]{3})*(?:[,.][0-9]{2})?)", candidate)
                    if match and len(re.findall(r"[0-9]{1,3}(?:[ \u00a0][0-9]{3})*(?:[,.][0-9]{2})?", candidate)) == 1:
                        return normalize_amount(match.group(1))
        return "0"

    def find_amount_below_keyword(keyword: str) -> str:
        keyword_norm = norm_label(keyword)
        for index, line in enumerate(lines):
            current_lower = norm_label(line)
            if keyword_norm in current_lower:
                for look_ahead in range(1, 4):
                    if index + look_ahead >= len(lines):
                        break
                    candidate = lines[index + look_ahead].strip()
                    if re.fullmatch(r"[0-9\s]+(?:[,.][0-9]{2})?", candidate):
                        return normalize_amount(candidate)
        return "0"

    def extract_vat_summary_amounts() -> tuple[str, str]:
        summary_patterns = [
            r"Celkem základ\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*Celkem DPH",
            r"Celkem základ\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*DPH\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Základ\s*21%\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*DPH",
            r"Zaklad\s*21%\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*([0-9\s]+(?:[,.][0-9]{2})?)\s*DPH",
        ]
        for pattern in summary_patterns:
            match = re.search(pattern, text, re.I)
            if match:
                return normalize_amount(match.group(1)), normalize_amount(match.group(2))
        return "0", "0"

    def find_inline_value(*labels: str) -> str:
        label_patterns = []
        for label in labels:
            escaped = re.escape(label)
            label_patterns.append(escaped.replace("\\ ", r"\s+"))
        pattern = rf"(?:{'|'.join(label_patterns)})\s*[: ]\s*([^\n]+)"
        match = re.search(pattern, text, re.I)
        return match.group(1).strip() if match else ""

    def find_date(*patterns: str) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, re.I)
            if match:
                value = match.group(1).strip()
                for fmt_in in ("%d.%m.%Y", "%Y-%m-%d"):
                    try:
                        return datetime.strptime(value, fmt_in).date().isoformat()
                    except ValueError:
                        continue
        generic = re.search(r"(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})", text)
        if generic:
            value = generic.group(1)
            for fmt_in in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(value, fmt_in).date().isoformat()
                except ValueError:
                    continue
        return date.today().isoformat()

    def parse_any_date(value: str) -> str:
        raw_value = str(value or "").strip()
        if not raw_value:
            return date.today().isoformat()
        match = re.search(r"(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})", raw_value)
        if not match:
            return date.today().isoformat()
        normalized = match.group(1)
        for fmt_in in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(normalized, fmt_in).date().isoformat()
            except ValueError:
                continue
        return date.today().isoformat()

    supplier = first_line_after_label("Dodavatel:", "DODAVATEL")
    if not supplier or supplier == ":":
        supplier_block = block_before_label("Dodavatel:", "DODAVATEL", max_lines=5)
        supplier = supplier_block[0] if supplier_block else ""
    supplier_match = re.search(r"Dodavatel\s*(.+?)(?:IČ|IC|DIČ|DIC|Datum|Odběratel)", text, re.I | re.S)
    if supplier_match:
        candidate = supplier_match.group(1).splitlines()[0].strip()
        if len(candidate) > len(supplier):
            supplier = candidate
    if not supplier:
        first_lines = [line for line in text.splitlines() if len(line.strip()) > 3]
        supplier = first_lines[0][:120] if first_lines else ""

    document_number = ""
    header_match = re.search(r"Faktura\s*-\s*daňový doklad\s*([A-Z0-9./_-]+)", text, re.I)
    if header_match:
        document_number = header_match.group(1).strip()
    for pattern in [r"(?:Faktura|Číslo faktury|Cislo faktury|Doklad)\s*[:#]?\s*([A-Z0-9_/-]{4,})", r"VS\s*[:#]?\s*([0-9]{4,})"]:
        match = re.search(pattern, text, re.I)
        if match:
            document_number = match.group(1).strip()
            break

    amount_with_vat = find_amount_below_keyword("uhradit")
    if amount_with_vat == "0":
        amount_with_vat = find_standalone_amount_after_label("Zbývá uhradit [Kč]", "Zbyva uhradit [Kc]", "Zbývá uhradit", "Zbyva uhradit", "Celkem k úhradě", "Celkem k uhrade")
    if amount_with_vat == "0":
        amount_with_vat = find_amount_near_labels("Zbývá uhradit", "Zbyva uhradit", "Celkem k úhradě", "Celkem k uhrade")
    if amount_with_vat == "0":
        amount_with_vat = find_amount(
            r"Zbývá uhradit\s*\[?[A-ZČKčkc]*\]?\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Zbyva uhradit\s*\[?[A-ZČKčkc]*\]?\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Celkem k úhradě\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Celkem\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
        )
    summary_base, summary_vat = extract_vat_summary_amounts()
    if summary_base == "0":
        for line in lines:
            normalized_line = norm_label(line)
            if (("CELKEM" in normalized_line and "KLAD" in normalized_line) or normalized_line.startswith("ZAKLAD 21")):
                amounts = re.findall(r"([0-9]{1,3}(?:[ \u00a0][0-9]{3})*(?:[,.][0-9]{2})?)", line)
                if amounts:
                    summary_base = normalize_amount(amounts[0])
                if len(amounts) > 1:
                    summary_vat = normalize_amount(amounts[1])
                break
    amount_without_vat = summary_base if summary_base != "0" else find_amount_near_labels("Celkem základ", "Celkem zaklad", "Bez DPH", "Základ", "Zaklad")
    if amount_without_vat == "0":
        amount_without_vat = find_amount(
            r"Celkem základ\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Celkem zaklad\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Bez DPH\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Základ\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
            r"Zaklad\s*[: ]\s*([0-9\s]+(?:[,.][0-9]{2})?)",
        )
    if amount_with_vat == "0" and amount_without_vat != "0":
        amount_with_vat = amount_without_vat
    if amount_with_vat != "0":
        amount = amount_with_vat
    else:
        amount = amount_without_vat
    vat_rate_match = re.search(r"(?:Sazba DPH|DPH)\s*[: ]\s*([0-9]{1,2}(?:[.,][0-9]+)?)", text, re.I)
    vat_rate = vat_rate_match.group(1).replace(",", ".") if vat_rate_match else ("21" if summary_vat != "0" else "0")
    variable_symbol = re.sub(r"\D", "", find_inline_value("Var. sym.", "Variabilní symbol", "Variabilni symbol"))[:20]

    header_dates = top_dates()
    issue_date_value = header_dates[0] if len(header_dates) >= 1 else first_line_before_label("Vystaveno:", "Datum vystavení:", "Datum vystaveni:")
    due_date_value = header_dates[1] if len(header_dates) >= 2 else first_line_before_label("Datum splatnosti:", "Den splatnosti:", "Datum splatnosti")

    title = f"Faktura {document_number}" if document_number else (supplier or "PDF doklad")
    if any("internet" in line.lower() or "tarif" in line.lower() or "linka" in line.lower() for line in lines):
        category = "Telefon a internet"
    elif any("software" in line.lower() or "licence" in line.lower() or "licence" in line.lower() for line in lines):
        category = "Software"
    else:
        category = "Ostatní"

    return {
        "title": title[:160],
        "supplier_name": supplier[:160],
        "document_number": document_number[:80],
        "expense_date": parse_any_date(issue_date_value) if issue_date_value else find_date(r"(?:Datum vystavení|Den vystavení|Vystaveno)\s*[: ]\s*(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})"),
        "due_date": parse_any_date(due_date_value) if due_date_value else find_date(r"(?:Datum splatnosti|Den splatnosti|Splatnost)\s*[: ]\s*(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})"),
        "amount": amount,
        "amount_with_vat": amount_with_vat if amount_with_vat != "0" else amount,
        "amount_without_vat": amount_without_vat if amount_without_vat != "0" else amount,
        "vat_rate": vat_rate,
        "currency": "CZK",
        "document_type": "Přijatá faktura",
        "status": "review",
        "category": category,
        "payment_method": "Převodem",
        "project_code": "",
        "variable_symbol": variable_symbol,
        "note": "Načteno z PDF - zkontroluj údaje před uložením.",
    }


@app.get("/api/ares")
def ares_lookup_api():
    ico = request.args.get("ico", "").strip()
    try:
        payload = fetch_ares_subject(ico)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    return jsonify({"ok": True, "data": payload})


@app.route("/", methods=["GET", "POST"])
def dashboard() -> str:
    conn = open_db()
    if request.method == "POST":
        set_app_setting(conn, "dashboard_notes", request.form.get("dashboard_notes", "").strip())
        flash("Poznámky na dashboardu byly uloženy.", "info")
        return redirect(url_for("dashboard"))

    all_invoices = list_invoices(conn)
    invoices = all_invoices[:8]
    customers = list_customers(conn)
    issuers = list_issuers(conn)
    expenses = list_expenses(conn)
    today = date.today()
    month_start = today.replace(day=1)
    total_all = 0.0
    total_paid = 0.0
    month_revenue = 0.0
    unpaid_amount = 0.0
    unpaid_count = 0
    overdue_amount = 0.0
    overdue_count = 0
    upcoming_due_count = 0
    draft_count = 0
    paid_cashflow = 0.0
    expected_cashflow = 0.0
    total_expenses = sum(float(item["amount"] or 0) for item in expenses)
    month_expenses = monthly_expense_total(conn, today.year, today.month)
    customer_totals: dict[str, dict[str, float]] = defaultdict(lambda: {"amount": 0.0, "count": 0.0})
    chart_monthly: list[dict[str, Any]] = []
    month_rows = year_month_overview(conn, today.year)
    month_map = {int(row["month"]): row for row in month_rows}
    for month in range(1, 13):
        row = month_map.get(month)
        chart_monthly.append(
            {
                "label": f"{month:02d}",
                "revenue": float(row["grand_total"] or 0) if row else 0.0,
                "count": int(row["invoice_count"] or 0) if row else 0,
            }
        )
    historical_invoice_years: set[int] = set()
    historical_month_counts: dict[int, int] = {month: 0 for month in range(1, 13)}
    historical_month_revenue: dict[int, float] = {month: 0.0 for month in range(1, 13)}

    invoice_rows = []
    alerts_overdue = []
    alerts_upcoming = []
    alerts_draft = []
    for row in all_invoices:
        totals = compute_totals(conn, int(row["id"]))
        grand_total = float(totals.grand_total)
        issue_date = parse_date(str(row["issue_date"]), default=today)
        due_date = parse_date(str(row["due_date"]), default=issue_date)
        status = str(row["status"])
        if issue_date.year != today.year:
            historical_invoice_years.add(issue_date.year)
            historical_month_counts[issue_date.month] = historical_month_counts.get(issue_date.month, 0) + 1
            historical_month_revenue[issue_date.month] = historical_month_revenue.get(issue_date.month, 0.0) + grand_total
        total_all += grand_total
        if status == "paid":
            total_paid += grand_total
            paid_cashflow += grand_total
        else:
            unpaid_amount += grand_total
            unpaid_count += 1
            expected_cashflow += grand_total
        if issue_date >= month_start:
            month_revenue += grand_total
        if status != "paid" and due_date < today:
            overdue_amount += grand_total
            overdue_count += 1
            if len(alerts_overdue) < 5:
                alerts_overdue.append({"id": row["id"], "number": row["invoice_number"], "customer": row["customer_name"], "due": format_date_cz(due_date.isoformat()), "total": f"{grand_total:.2f} {row['currency']}"})
        elif status != "paid" and due_date <= today + timedelta(days=3):
            upcoming_due_count += 1
            if len(alerts_upcoming) < 5:
                alerts_upcoming.append({"id": row["id"], "number": row["invoice_number"], "customer": row["customer_name"], "due": format_date_cz(due_date.isoformat()), "total": f"{grand_total:.2f} {row['currency']}"})
        if status == "draft":
            draft_count += 1
            if len(alerts_draft) < 5:
                alerts_draft.append({"id": row["id"], "number": row["invoice_number"], "customer": row["customer_name"], "issue": format_date_cz(issue_date.isoformat())})
        bucket = customer_totals[str(row["customer_name"])]
        bucket["amount"] += grand_total
        bucket["count"] += 1

    for row in invoices:
        totals = compute_totals(conn, int(row["id"]))
        computed_status = str(row["status"])
        due_date = parse_date(str(row["due_date"]), default=today)
        if computed_status != "paid" and due_date < today:
            computed_status = "overdue"
        invoice_rows.append({"id": row["id"], "number": row["invoice_number"], "customer": row["customer_name"], "issuer": row["issuer_name"] or "-", "status": status_pill(computed_status), "total": f"{totals.grand_total:.2f} {row['currency']}", "issue": format_date_cz(row["issue_date"])})

    total_unpaid = max(total_all - total_paid, 0.0)
    current_year_income_total = 0.0
    current_year_unpaid_total = 0.0
    for row in all_invoices:
        issue_date = parse_date(str(row["issue_date"]), default=today)
        if issue_date.year != today.year:
            continue
        totals = compute_totals(conn, int(row["id"]))
        grand_total = float(totals.grand_total)
        current_year_income_total += grand_total
        if str(row["status"]) != "paid":
            current_year_unpaid_total += grand_total
    historical_year_count = len(historical_invoice_years)
    chart_revenue_avg_raw = []
    for month in range(1, 13):
        avg_revenue = (historical_month_revenue.get(month, 0.0) / historical_year_count) if historical_year_count else 0.0
        chart_revenue_avg_raw.append({"label": f"{month:02d}", "value": avg_revenue})
    revenue_peak = max(
        [float(item["revenue"]) for item in chart_monthly]
        + [float(item["value"]) for item in chart_revenue_avg_raw]
        + [1.0]
    )
    chart_counts_avg_raw = []
    for month in range(1, 13):
        avg_count = (historical_month_counts.get(month, 0) / historical_year_count) if historical_year_count else 0.0
        chart_counts_avg_raw.append({"label": f"{month:02d}", "value": avg_count})
    count_peak = max(
        [float(item["count"]) for item in chart_monthly]
        + [float(item["value"]) for item in chart_counts_avg_raw]
        + [1.0]
    )
    chart_revenue = [{"label": item["label"], "height": max(8, int((item["revenue"] / revenue_peak) * 120)) if item["revenue"] else 8, "value": f"{item['revenue']:.0f} Kč"} for item in chart_monthly]
    chart_counts = [{"label": item["label"], "height": max(8, int((item["count"] / count_peak) * 120)) if item["count"] else 8, "value": item["count"], "value_text": str(item["count"])} for item in chart_monthly]
    line_width = 356.0
    line_height = 120.0
    step_x = line_width / 11 if len(chart_counts_avg_raw) > 1 else 0.0
    chart_revenue_avg_points = []
    for index, item in enumerate(chart_revenue_avg_raw):
        avg_value = float(item["value"])
        y = line_height - ((avg_value / revenue_peak) * line_height if revenue_peak else 0.0)
        chart_revenue_avg_points.append(f"{index * step_x:.2f},{y:.2f}")
    chart_counts_avg_points = []
    for index, item in enumerate(chart_counts_avg_raw):
        avg_value = float(item["value"])
        y = line_height - ((avg_value / count_peak) * line_height if count_peak else 0.0)
        chart_counts_avg_points.append(f"{index * step_x:.2f},{y:.2f}")
    chart_revenue_avg = [
        {
            "label": item["label"],
            "value": item["value"],
            "formatted": f"{item['value']:.0f} Kč",
        }
        for item in chart_revenue_avg_raw
    ]
    chart_counts_avg = [
        {
            "label": item["label"],
            "value": item["value"],
            "formatted": f"{item['value']:.1f}".replace(".", ","),
        }
        for item in chart_counts_avg_raw
    ]
    paid_ratio = 0 if total_all <= 0 else int((total_paid / total_all) * 100)
    unpaid_ratio = max(0, 100 - paid_ratio)
    top_customers = sorted(
        [{"name": name, "amount": data["amount"], "count": int(data["count"])} for name, data in customer_totals.items()],
        key=lambda item: item["amount"],
        reverse=True,
    )[:5]
    notes = get_app_setting(conn, "dashboard_notes", "")
    annual_overview = yearly_overview(conn)
    body = """
    <style>
      .dash-grid{display:grid;gap:18px}
      .dash-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px}
      .dash-kpi{padding:22px;background:linear-gradient(180deg,rgba(20,40,36,.98),rgba(13,26,23,.98));border:1px solid var(--line);border-radius:22px;box-shadow:var(--shadow)}
      .dash-kpi .icon{font-size:18px;opacity:.9}
      .dash-kpi .label{margin-top:10px;color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.08em;font-weight:700}
      .dash-kpi .value{margin-top:8px;font-size:30px;font-weight:900;letter-spacing:-.04em}
      .dash-kpi .sub{margin-top:8px;color:var(--muted);font-size:14px}
      .quick-actions{display:flex;flex-wrap:wrap;gap:12px}
      .quick-actions a{min-width:180px}
      .chart-grid{display:grid;grid-template-columns:1.4fr 1fr 1fr;gap:18px}
      .chart-card{padding:20px;background:rgba(15,30,27,.92);border:1px solid var(--line);border-radius:24px;box-shadow:var(--shadow)}
      .bars{display:flex;align-items:end;gap:10px;height:160px;margin-top:16px}
      .bars-overlay{position:relative;height:160px;margin-top:16px}
      .bars-overlay .bars{position:relative;z-index:2;margin-top:0}
      .line-svg{position:absolute;left:0;right:0;top:0;height:120px;width:100%;z-index:3;pointer-events:none;overflow:visible}
      .line-path{fill:none;stroke:#c084fc;stroke-width:3;stroke-linecap:round;stroke-linejoin:round;filter:drop-shadow(0 0 6px rgba(192,132,252,.28))}
      .line-dot{fill:#f5d0fe;stroke:#7c3aed;stroke-width:2}
      .bar-wrap{display:grid;gap:8px;justify-items:center;flex:1}
      .bar{position:relative;width:100%;max-width:28px;border-radius:10px 10px 4px 4px;background:linear-gradient(180deg,#2dd4bf,#0f766e);display:flex;align-items:center;justify-content:center;overflow:visible}
      .bar.alt{background:linear-gradient(180deg,#7dd3fc,#2563eb)}
      .bar-value{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%) rotate(-90deg);transform-origin:center;white-space:nowrap;font-size:10px;font-weight:800;letter-spacing:.02em;color:rgba(255,255,255,.92);text-shadow:0 1px 2px rgba(0,0,0,.35);pointer-events:none}
      .bar-label{font-size:11px;color:var(--muted)}
      .chart-legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:12px;color:var(--muted);font-size:12px}
      .chart-legend span{display:inline-flex;align-items:center;gap:8px}
      .chart-legend i{display:inline-block}
      .legend-bar{width:14px;height:14px;border-radius:4px;background:linear-gradient(180deg,#7dd3fc,#2563eb)}
      .legend-line{width:16px;height:0;border-top:3px solid #c084fc;border-radius:999px}
      .alert-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:18px}
      .alert-card{padding:20px;border-radius:22px;border:1px solid var(--line);box-shadow:var(--shadow)}
      .alert-card.red{background:rgba(58,16,22,.92);border-color:#7f1d1d}
      .alert-card.orange{background:rgba(61,33,10,.92);border-color:#92400e}
      .alert-card.blue{background:rgba(16,35,59,.92);border-color:#1d4ed8}
      .alert-card h3{margin:0 0 6px;font-size:18px}
      .alert-list{display:grid;gap:8px;margin-top:14px}
      .alert-item{padding:10px 12px;border-radius:14px;background:rgba(255,255,255,.04)}
      .two-col{display:grid;grid-template-columns:1.2fr .8fr;gap:18px}
      .top-list,.timeline{display:grid;gap:10px}
      .top-row,.timeline-row{display:grid;grid-template-columns:1fr auto auto;gap:12px;padding:10px 0;border-bottom:1px solid var(--line)}
      .top-row:last-child,.timeline-row:last-child{border-bottom:none}
      .mini-stats{display:grid;gap:8px;margin-top:12px}
      .mini-row{display:grid;grid-template-columns:1fr auto;gap:12px;align-items:center;padding:8px 0;border-bottom:1px solid var(--line)}
      .mini-row:last-child{border-bottom:none}
      .notes-box textarea{min-height:150px}
      .donut{margin-top:18px;height:14px;border-radius:999px;overflow:hidden;background:var(--panel-2);display:flex}
      .donut .paid{background:linear-gradient(90deg,#34d399,#15803d)}
      .donut .unpaid{background:linear-gradient(90deg,#f59e0b,#b45309)}
      @media (max-width:1100px){.dash-kpis,.chart-grid,.alert-grid,.two-col{grid-template-columns:1fr}}
    </style>
    <section class="dash-grid">
      <section class="dash-kpis">
        <div class="dash-kpi"><div class="icon">💰</div><div class="label">Tržby za měsíc</div><div class="value">{{ month_revenue }}</div><div class="sub">Aktuální měsíc od {{ month_start }}</div></div>
        <div class="dash-kpi"><div class="icon">📄</div><div class="label">Počet vystavených faktur</div><div class="value">{{ invoice_count }}</div><div class="sub">{{ customer_count }} odběratelů, {{ issuer_count }} firem</div></div>
        <div class="dash-kpi"><div class="icon">⏳</div><div class="label">Nezaplacené faktury</div><div class="value">{{ unpaid_amount }}</div><div class="sub">{{ unpaid_count }} faktur čeká na úhradu</div></div>
        <div class="dash-kpi"><div class="icon">⚠️</div><div class="label">Po splatnosti</div><div class="value">{{ overdue_amount }}</div><div class="sub">{{ overdue_count }} faktur po splatnosti</div></div>
      </section>

      <section class="dash-kpis" style="grid-template-columns:repeat(3,minmax(0,1fr));">
        <div class="dash-kpi"><div class="icon">💸</div><div class="label">Výdaje za měsíc</div><div class="value">{{ month_expenses }}</div><div class="sub">Výdaje v aktuálním měsíci</div></div>
        <div class="dash-kpi"><div class="icon">🧮</div><div class="label">Čistý výsledek za měsíc</div><div class="value">{{ month_profit }}</div><div class="sub">Tržby mínus výdaje</div></div>
        <div class="dash-kpi"><div class="icon">📉</div><div class="label">Výdaje celkem</div><div class="value">{{ total_expenses }}</div><div class="sub">{{ expense_count }} evidovaných výdajů</div></div>
      </section>

      <section class="chart-grid">
        <div class="chart-card">
          <div class="section-head"><div><h2>Tržby v čase</h2><div class="section-copy">Měsíční vývoj příjmů v aktuálním roce. Křivka ukazuje průměr minulých let bez aktuálního roku.</div></div></div>
          <div class="bars-overlay">
            <svg class="line-svg" viewBox="0 0 356 120" preserveAspectRatio="none" aria-hidden="true">
              <polyline class="line-path" points="{{ chart_revenue_avg_points }}"></polyline>
              {% for point in chart_revenue_avg_plot %}
              <circle class="line-dot" cx="{{ point.x }}" cy="{{ point.y }}" r="4">
                <title>{{ point.label }} • {{ point.formatted }}</title>
              </circle>
              {% endfor %}
            </svg>
            <div class="bars">{% for item in chart_revenue %}<div class="bar-wrap"><div class="bar" style="height:{{ item.height }}px" title="{{ item.value }}"><span class="bar-value">{{ item.value }}</span></div><div class="bar-label">{{ item.label }}</div></div>{% endfor %}</div>
          </div>
          <div class="chart-legend">
            <span><i class="legend-bar" style="background:linear-gradient(180deg,#2dd4bf,#0f766e)"></i> Aktuální rok</span>
            <span><i class="legend-line"></i> Průměr minulých let</span>
          </div>
        </div>
        <div class="chart-card">
          <div class="section-head"><div><h2>Zaplaceno vs nezaplaceno</h2><div class="section-copy">Rychlý poměr cashflow.</div></div></div>
          <div class="metric-value" style="font-size:26px">{{ total_paid }}</div>
          <div class="mini-stats">
            <div class="mini-row"><strong>Zaplaceno / nezaplaceno</strong><span>{{ total_paid }} / {{ total_unpaid }}</span></div>
            <div class="mini-row"><strong>Příjem za rok / nezaplaceno</strong><span>{{ current_year_income_total }} / {{ current_year_unpaid_total }}</span></div>
          </div>
          <div class="donut"><div class="paid" style="width:{{ paid_ratio }}%"></div><div class="unpaid" style="width:{{ unpaid_ratio }}%"></div></div>
        </div>
        <div class="chart-card">
          <div class="section-head"><div><h2>Počet faktur v čase</h2><div class="section-copy">Kolik faktur vzniklo v jednotlivých měsících. Křivka ukazuje průměr minulých let bez aktuálního roku.</div></div></div>
          <div class="bars-overlay">
            <svg class="line-svg" viewBox="0 0 356 120" preserveAspectRatio="none" aria-hidden="true">
              <polyline class="line-path" points="{{ chart_counts_avg_points }}"></polyline>
              {% for point in chart_counts_avg_plot %}
              <circle class="line-dot" cx="{{ point.x }}" cy="{{ point.y }}" r="4">
                <title>{{ point.label }} • {{ point.formatted }}</title>
              </circle>
              {% endfor %}
            </svg>
            <div class="bars">{% for item in chart_counts %}<div class="bar-wrap"><div class="bar alt" style="height:{{ item.height }}px" title="{{ item.value }}"><span class="bar-value">{{ item.value_text }}</span></div><div class="bar-label">{{ item.label }}</div></div>{% endfor %}</div>
          </div>
          <div class="chart-legend">
            <span><i class="legend-bar"></i> Aktuální rok</span>
            <span><i class="legend-line"></i> Průměr minulých let</span>
          </div>
        </div>
      </section>

      <section class="alert-grid">
        <div class="alert-card red"><h3>❗ Faktury po splatnosti</h3><div class="section-copy">{{ overdue_count }} faktur • {{ overdue_amount }}</div><div class="alert-list">{% for item in alerts_overdue %}<div class="alert-item"><a href="{{ url_for('invoice_detail_page', invoice_id=item.id) }}"><strong>{{ item.number }}</strong></a><br>{{ item.customer }} • {{ item.total }} • {{ item.due }}</div>{% endfor %}{% if not alerts_overdue %}<div class="alert-item">Žádné faktury po splatnosti.</div>{% endif %}</div></div>
        <div class="alert-card orange"><h3>⏰ Před splatností do 3 dnů</h3><div class="section-copy">{{ upcoming_due_count }} faktur vyžaduje pozornost</div><div class="alert-list">{% for item in alerts_upcoming %}<div class="alert-item"><a href="{{ url_for('invoice_detail_page', invoice_id=item.id) }}"><strong>{{ item.number }}</strong></a><br>{{ item.customer }} • {{ item.total }} • {{ item.due }}</div>{% endfor %}{% if not alerts_upcoming %}<div class="alert-item">V nejbližších 3 dnech nic nesplatné.</div>{% endif %}</div></div>
        <div class="alert-card blue"><h3>🔔 Neodeslané faktury</h3><div class="section-copy">{{ draft_count }} faktur ve stavu rozpracováno</div><div class="alert-list">{% for item in alerts_draft %}<div class="alert-item"><a href="{{ url_for('invoice_detail_page', invoice_id=item.id) }}"><strong>{{ item.number }}</strong></a><br>{{ item.customer }} • vystavení {{ item.issue }}</div>{% endfor %}{% if not alerts_draft %}<div class="alert-item">Žádné rozpracované faktury.</div>{% endif %}</div></div>
      </section>

      <section class="two-col">
        <div class="table-card"><div class="section-head"><div><h2>Poslední faktury</h2><div class="section-copy">Rychlý seznam posledních dokladů.</div></div><a class="button-link" href="{{ url_for('invoices_page') }}">Všechny faktury</a></div><table><thead><tr><th>Faktura</th><th>Klient</th><th>Částka</th><th>Stav</th></tr></thead><tbody>{% for row in invoice_rows %}<tr><td><a href="{{ url_for('invoice_detail_page', invoice_id=row.id) }}"><strong>{{ row.number }}</strong></a></td><td>{{ row.customer }}</td><td>{{ row.total }}</td><td>{{ row.status|safe }}</td></tr>{% endfor %}</tbody></table></div>
        <div class="content-card"><div class="section-head"><div><h2>Top zákazníci</h2><div class="section-copy">Kdo přináší nejvíc tržeb.</div></div></div><div class="top-list">{% for row in top_customers %}<div class="top-row"><div><strong>{{ row.name }}</strong></div><div>{{ row.count }} faktur</div><div>{{ '%.2f'|format(row.amount) }} Kč</div></div>{% endfor %}{% if not top_customers %}<div class="muted">Zatím nejsou data.</div>{% endif %}</div></div>
      </section>

      <section class="two-col">
        <div class="content-card"><div class="section-head"><div><h2>Cashflow</h2><div class="section-copy">Kolik už přišlo a kolik se očekává.</div></div></div><div class="top-list"><div class="top-row"><div><strong>Přišlo</strong></div><div></div><div>{{ paid_cashflow }}</div></div><div class="top-row"><div><strong>Očekávané příjmy</strong></div><div></div><div>{{ expected_cashflow }}</div></div><div class="top-row"><div><strong>Příjmy tento měsíc</strong></div><div></div><div>{{ month_revenue }}</div></div></div></div>
        <div class="content-card"><div class="section-head"><div><h2>Rychlé poznámky / TODO</h2><div class="section-copy">Interní úkoly a připomínky na jednom místě.</div></div></div><form method="post" class="notes-box stack"><div class="field"><label>Poznámky</label><textarea name="dashboard_notes" placeholder="Např. zavolat klientovi, poslat upomínku, připravit podklady">{{ notes }}</textarea></div><div><button class="button" type="submit">Uložit poznámky</button></div></form></div>
      </section>

      <details class="table-card">
        <summary style="list-style:none;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:16px;">
          <div>
            <h2 style="margin:0;">Roční přehled</h2>
            <div class="section-copy">Souhrn faktur podle roku. Po otevření dashboardu je blok sbalený.</div>
          </div>
          <div class="toolbar">
            <span class="muted">{{ annual_overview|length }} roků</span>
            <span class="button-link">Otevřít</span>
          </div>
        </summary>
        <div style="margin-top:14px;">
          <div class="toolbar" style="margin-bottom:14px;">
            <a class="button-link" href="{{ url_for('year_report_page') }}">Detail ročního přehledu</a>
          </div>
          <table>
            <thead><tr><th>Rok</th><th>Počet faktur</th><th>Celkem</th><th>Zaplaceno</th><th>Nezaplaceno</th></tr></thead>
            <tbody>
              {% for row in annual_overview %}
              <tr>
                <td>{{ row.year }}</td>
                <td>{{ row.invoice_count }}</td>
                <td>{{ '%.2f'|format(row.grand_total|float) }}</td>
                <td>{{ '%.2f'|format(row.paid_total|float) }}</td>
                <td>{{ '%.2f'|format((row.grand_total|float) - (row.paid_total|float)) }}</td>
              </tr>
              {% endfor %}
              {% if not annual_overview %}
              <tr><td colspan="5" class="muted">Zatím nejsou dostupná žádná roční data.</td></tr>
              {% endif %}
            </tbody>
          </table>
        </div>
      </details>
    </section>
    """
    return render_page(
        "Dashboard",
        "Provozní přehled fakturace, cashflow a upozornění.",
        body,
        "dashboard",
        invoice_count=len(all_invoices),
        customer_count=len(customers),
        issuer_count=len(issuers),
        total_all=f"{total_all:.2f} Kč",
        total_paid=f"{total_paid:.2f} Kč",
        total_unpaid=f"{total_unpaid:.2f} Kč",
        current_year_income_total=f"{current_year_income_total:.2f} Kč",
        current_year_unpaid_total=f"{current_year_unpaid_total:.2f} Kč",
        month_revenue=f"{month_revenue:.2f} Kč",
        unpaid_amount=f"{unpaid_amount:.2f} Kč",
        overdue_amount=f"{overdue_amount:.2f} Kč",
        month_expenses=f"{month_expenses:.2f} Kč",
        month_profit=f"{(month_revenue - month_expenses):.2f} Kč",
        total_expenses=f"{total_expenses:.2f} Kč",
        expense_count=len(expenses),
        unpaid_count=unpaid_count,
        overdue_count=overdue_count,
        upcoming_due_count=upcoming_due_count,
        draft_count=draft_count,
        invoice_rows=invoice_rows,
        chart_revenue=chart_revenue,
        chart_revenue_avg_points=" ".join(chart_revenue_avg_points),
        chart_revenue_avg_plot=[
            {
                "x": f"{index * step_x:.2f}",
                "y": f"{line_height - ((float(item['value']) / revenue_peak) * line_height if revenue_peak else 0.0):.2f}",
                "label": item["label"],
                "formatted": item["formatted"],
            }
            for index, item in enumerate(chart_revenue_avg)
        ],
        chart_counts=chart_counts,
        chart_counts_avg_points=" ".join(chart_counts_avg_points),
        chart_counts_avg_plot=[
            {
                "x": f"{index * step_x:.2f}",
                "y": f"{line_height - ((float(item['value']) / count_peak) * line_height if count_peak else 0.0):.2f}",
                "label": item["label"],
                "formatted": item["formatted"],
            }
            for index, item in enumerate(chart_counts_avg)
        ],
        paid_ratio=paid_ratio,
        unpaid_ratio=unpaid_ratio,
        alerts_overdue=alerts_overdue,
        alerts_upcoming=alerts_upcoming,
        alerts_draft=alerts_draft,
        top_customers=top_customers,
        paid_cashflow=f"{paid_cashflow:.2f} Kč",
        expected_cashflow=f"{expected_cashflow:.2f} Kč",
        notes=notes,
        annual_overview=annual_overview,
        month_start=month_start.isoformat(),
    )

@app.route("/customers")
def customers_page() -> str:
    conn = open_db()
    customer_rows = conn.execute(
        """
        SELECT
            c.id,
            c.name,
            c.email,
            c.phone,
            c.ico,
            c.dic,
            c.address,
            COALESCE(stats.invoice_count, 0) AS invoice_count,
            COALESCE(stats.invoice_total, 0) AS invoice_total
        FROM customers c
        LEFT JOIN (
            SELECT
                i.customer_id AS customer_id,
                COUNT(DISTINCT i.id) AS invoice_count,
                COALESCE(SUM(ii.quantity * ii.unit_price * (1 + (ii.vat_rate / 100.0))), 0) AS invoice_total
            FROM invoices i
            LEFT JOIN invoice_items ii ON ii.invoice_id = i.id
            GROUP BY i.customer_id
        ) stats ON stats.customer_id = c.id
        ORDER BY lower(c.name) ASC, c.id ASC
        """
    ).fetchall()
    body = """
    <section class="table-card">
      <style>
        .customers-table{table-layout:fixed}
        .customers-table th,.customers-table td{padding:6px 8px;line-height:1.18}
        .customers-table th{font-size:12px}
        .customers-table .name-cell strong{display:block;font-size:16px;line-height:1.1;margin-bottom:2px}
        .customers-table .name-meta{display:block;font-size:13px;line-height:1.18;color:var(--muted)}
        .customers-table .compact-cell{font-size:13px}
        .customers-table .invoice-summary{display:grid;gap:4px}
        .customers-table .invoice-count-link,
        .customers-table .invoice-count-empty{display:inline-flex;min-width:34px;justify-content:center;align-items:center;padding:3px 7px;border-radius:999px;font-weight:800;justify-self:start;font-size:12px}
        .customers-table .invoice-count-link{background:var(--accent-soft);color:var(--accent);border:1px solid var(--line)}
        .customers-table .invoice-count-empty{background:var(--panel-2);color:var(--muted);border:1px solid var(--line)}
        .customers-table .invoice-total{font-size:11px;line-height:1.15;color:var(--muted);white-space:nowrap}
        .customers-table .presence-icons{display:flex;gap:4px;align-items:center;flex-wrap:wrap}
        .customers-table .presence-dot{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:999px;border:1px solid var(--line);font-size:12px;font-weight:800}
        .customers-table .presence-dot.is-on{background:#123327;color:#86efac}
        .customers-table .presence-dot.is-off{background:var(--panel-2);color:var(--muted)}
        .customers-table .toolbar{gap:5px;flex-wrap:nowrap}
        .customers-table .toolbar form{margin:0}
        .customers-table .button-link,.customers-table .button-danger{padding:6px 8px;border-radius:12px;font-size:12px;white-space:nowrap}
      </style>
      <div class="section-head">
        <div><h2>Seznam odběratelů</h2><div class="section-copy">Přehled kontaktů pro vystavování faktur.</div></div>
        <a class="button" href="{{ url_for('new_customer_page') }}">Nový odběratel</a>
      </div>
      <table class="customers-table">
        <thead>
          <tr>
            <th style="width:44%;">Odběratel</th>
            <th style="width:12%;">Vyplnění</th>
            <th style="width:12%;">Faktury</th>
            <th style="width:24%;">Akce</th>
          </tr>
        </thead>
        <tbody>
          {% for row in rows %}
          <tr>
            <td class="name-cell">
              <strong>{{ row.name }}</strong>
              <span class="name-meta">{{ row.ico or 'Bez IČO' }}</span>
            </td>
            <td class="compact-cell">
              <div class="presence-icons">
                <span class="presence-dot {{ 'is-on' if row.address else 'is-off' }}" title="Adresa {{ 'vyplněná' if row.address else 'chybí' }}">A</span>
                <span class="presence-dot {{ 'is-on' if row.phone else 'is-off' }}" title="Telefon {{ 'vyplněný' if row.phone else 'chybí' }}">T</span>
                <span class="presence-dot {{ 'is-on' if row.email else 'is-off' }}" title="E-mail {{ 'vyplněný' if row.email else 'chybí' }}">E</span>
              </div>
            </td>
            <td class="compact-cell">
              <div class="invoice-summary">
                {% if row.invoice_count %}
                <a class="invoice-count-link" href="{{ url_for('customer_invoices_page', customer_id=row.id) }}" title="Zobrazit všechny faktury odběratele">{{ row.invoice_count }}</a>
                {% else %}
                <span class="invoice-count-empty">0</span>
                {% endif %}
                <span class="invoice-total">{{ '%.2f'|format(row.invoice_total) }} Kč</span>
              </div>
            </td>
            <td>
              <div class="toolbar">
                <a class="button-link" href="{{ url_for('new_invoice_page', customer_id=row.id) }}">Fakturovat</a>
                <a class="button-link" href="{{ url_for('edit_customer_page', customer_id=row.id) }}">Upravit</a>
                <form method="post" action="{{ url_for('delete_customer_page', customer_id=row.id) }}" onsubmit="return confirm('Opravdu smazat odběratele?');">
                  <button class="button-danger" type="submit">Smazat</button>
                </form>
              </div>
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </section>
    """
    return render_page("Odběratelé", "Správa kontaktů a fakturačních údajů.", body, "customers", rows=customer_rows, ares_script=ARES_LOOKUP_SCRIPT)


@app.route("/customers/<int:customer_id>/invoices")
def customer_invoices_page(customer_id: int) -> str:
    conn = open_db()
    customer = next((item for item in list_customers(conn) if int(item["id"]) == customer_id), None)
    if customer is None:
        flash("Odběratel nebyl nalezen.", "error")
        return redirect(url_for("customers_page"))

    selected_year = request.args.get("year", "").strip()
    invoice_rows_raw = conn.execute(
        """
        SELECT id, invoice_number, issue_date, due_date, status, currency
        FROM invoices
        WHERE customer_id = ?
        ORDER BY issue_date DESC, invoice_number DESC, id DESC
        """,
        (customer_id,),
    ).fetchall()
    available_years = sorted(
        {
            parse_date(str(row["issue_date"]), default=date.today()).year
            for row in invoice_rows_raw
            if str(row["issue_date"] or "").strip()
        },
        reverse=True,
    )
    selected_year_int = 0
    if selected_year:
        try:
            selected_year_int = int(selected_year)
        except ValueError:
            selected_year_int = 0
    if selected_year_int and selected_year_int not in available_years:
        selected_year_int = 0

    invoice_rows = []
    total_amount = 0.0
    for row in invoice_rows_raw:
        issue_date = parse_date(str(row["issue_date"]), default=date.today())
        if selected_year_int and issue_date.year != selected_year_int:
            continue
        totals = compute_totals(conn, int(row["id"]))
        grand_total = float(totals.grand_total or 0)
        total_amount += grand_total
        status = str(row["status"] or "").strip().lower()
        due_date = parse_date(str(row["due_date"]), default=date.today())
        if status != "paid" and due_date < date.today():
            status = "overdue"
        invoice_rows.append(
            {
                "id": int(row["id"]),
                "invoice_number": row["invoice_number"] or f"Faktura {row['id']}",
                "issue_date": format_date_cz(issue_date.isoformat()),
                "due_date": format_date_cz(row["due_date"]),
                "status": status_pill(status),
                "total": f"{grand_total:.2f} {row['currency'] or 'CZK'}",
            }
        )

    body = """
    <section class="table-card">
      <div class="section-head">
        <div>
          <h2>Faktury odběratele</h2>
          <div class="section-copy">{{ customer.name }} • {{ invoice_rows|length }} faktur{% if selected_year_label %} • rok {{ selected_year_label }}{% endif %} • {{ total_amount }}</div>
        </div>
        <a class="button-link" href="{{ url_for('customers_page') }}">Zpět na odběratele</a>
      </div>
      <form method="get" class="toolbar" style="margin-bottom:14px;">
        <div class="field" style="min-width:220px;margin:0;">
          <label>Rok</label>
          <select name="year" onchange="this.form.submit()">
            <option value="">Všechny roky</option>
            {% for year in available_years %}
            <option value="{{ year }}" {% if year == selected_year_label %}selected{% endif %}>{{ year }}</option>
            {% endfor %}
          </select>
        </div>
      </form>
      <table>
        <thead>
          <tr><th>Faktura</th><th>Vystavení</th><th>Splatnost</th><th>Částka</th><th>Stav</th></tr>
        </thead>
        <tbody>
          {% for row in invoice_rows %}
          <tr>
            <td><a href="{{ url_for('invoice_detail_page', invoice_id=row.id) }}"><strong>{{ row.invoice_number }}</strong></a></td>
            <td>{{ row.issue_date }}</td>
            <td>{{ row.due_date }}</td>
            <td>{{ row.total }}</td>
            <td>{{ row.status|safe }}</td>
          </tr>
          {% endfor %}
          {% if not invoice_rows %}
          <tr><td colspan="5" class="muted">K tomuto odběrateli zatím není vystavená žádná faktura.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </section>
    """
    return render_page(
        "Faktury odběratele",
        "Přehled všech faktur vystavených pro vybraného odběratele.",
        body,
        "customers",
        customer=customer,
        invoice_rows=invoice_rows,
        available_years=available_years,
        selected_year_label=selected_year_int or None,
        total_amount=f"{total_amount:.2f} Kč",
    )


@app.route("/customers/new", methods=["GET", "POST"])
def new_customer_page() -> str:
    conn = open_db()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("Název odběratele je povinný.", "error")
            return redirect(url_for("new_customer_page"))
        add_customer(conn, name, request.form.get("email", ""), request.form.get("address", ""), request.form.get("phone", ""), request.form.get("ico", ""), request.form.get("dic", ""))
        flash("Odběratel byl uložen.", "info")
        return redirect(url_for("customers_page"))

    body = """
    <section class="content-card"><div class="section-head"><div><h2>Nový odběratel</h2><div class="section-copy">Přidej kontaktní a fakturační údaje.</div></div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název</label><input id="customer-name" name="name" required></div><div class="field"><label>E-mail</label><input name="email"></div><div class="field"><label>Telefon</label><input name="phone"></div><div class="field"><label>IČO</label><div class="lookup-row"><input id="customer-ico" name="ico"><button class="button-secondary" type="button" onclick="lookupAres('customer')">Načíst z ARES</button></div><div class="lookup-status" id="customer-ares-status"></div></div><div class="field"><label>DIČ</label><input id="customer-dic" name="dic"></div><div class="field"><label>Adresa</label><textarea id="customer-address" name="address" rows="3"></textarea></div></div><div class="toolbar"><button class="button" type="submit">Uložit odběratele</button><a class="button-link" href="{{ url_for('customers_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Nový odběratel", "Samostatná stránka pro založení odběratele.", body, "customers", ares_script=ARES_LOOKUP_SCRIPT)


@app.route("/customers/<int:customer_id>", methods=["GET", "POST"])
def edit_customer_page(customer_id: int) -> str:
    conn = open_db()
    row = next((item for item in list_customers(conn) if int(item["id"]) == customer_id), None)
    if row is None:
        flash("Odběratel nebyl nalezen.", "error")
        return redirect(url_for("customers_page"))
    if request.method == "POST":
        update_customer(conn, customer_id, request.form.get("name", ""), request.form.get("email", ""), request.form.get("address", ""), request.form.get("phone", ""), request.form.get("ico", ""), request.form.get("dic", ""))
        flash("Odběratel byl aktualizován.", "info")
        return redirect(url_for("customers_page"))
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Upravit odběratele</h2><div class="section-copy">Změň údaje a ulož je zpět do databáze.</div></div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název</label><input id="customer-name" name="name" value="{{ row.name }}" required></div><div class="field"><label>E-mail</label><input name="email" value="{{ row.email or '' }}"></div><div class="field"><label>Telefon</label><input name="phone" value="{{ row.phone or '' }}"></div><div class="field"><label>IČO</label><div class="lookup-row"><input id="customer-ico" name="ico" value="{{ row.ico or '' }}"><button class="button-secondary" type="button" onclick="lookupAres('customer')">Načíst z ARES</button></div><div class="lookup-status" id="customer-ares-status"></div></div><div class="field"><label>DIČ</label><input id="customer-dic" name="dic" value="{{ row.dic or '' }}"></div><div class="field"><label>Adresa</label><textarea id="customer-address" name="address" rows="3">{{ row.address or '' }}</textarea></div></div><div class="toolbar"><button class="button" type="submit">Uložit změny</button><a class="button-link" href="{{ url_for('customers_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Upravit odběratele", "Editace kontaktu v novém webovém rozhraní.", body, "customers", row=row, ares_script=ARES_LOOKUP_SCRIPT)


@app.post("/customers/<int:customer_id>/delete")
def delete_customer_page(customer_id: int):
    conn = open_db()
    deleted = delete_customer_if_unused(conn, customer_id)
    if deleted:
        flash("Odběratel byl smazán.", "info")
    else:
        flash("Odběratele nelze smazat, protože má navázanou fakturu.", "error")
    return redirect(url_for("customers_page"))


@app.route("/issuers")
def issuers_page() -> str:
    conn = open_db()
    body = """
    <section class="table-card"><div class="section-head"><div><h2>Seznam firem</h2><div class="section-copy">Dodavatelé dostupní při vystavení faktury.</div></div><a class="button" href="{{ url_for('new_issuer_page') }}">Nová firma</a></div><table><thead><tr><th>Firma</th><th>IČO / DIČ</th><th>DPH</th><th>Akce</th></tr></thead><tbody>{% for row in rows %}<tr><td><strong>{{ row.company_name }}</strong><br><span class="muted">{{ (row.address or '-')|replace('\n', '<br>')|safe }}</span></td><td>{{ row.ico or '-' }} / {{ row.dic or '-' }}</td><td>{{ 'Plátce' if row.vat_payer else 'Neplátce' }}{% if row.is_default %} / Výchozí{% endif %}</td><td><div class="toolbar"><a class="button-link" href="{{ url_for('edit_issuer_page', issuer_id=row.id) }}">Upravit</a><form method="post" action="{{ url_for('delete_issuer_page', issuer_id=row.id) }}" onsubmit="return confirm('Opravdu smazat firmu?');"><button class="button-danger" type="submit">Smazat</button></form></div></td></tr>{% endfor %}</tbody></table></section>
    """
    return render_page("Firmy", "Správa fakturujících firem v novém webovém rozhraní.", body, "issuers", rows=list_issuers(conn), ares_script=ARES_LOOKUP_SCRIPT)


@app.route("/issuers/new", methods=["GET", "POST"])
def new_issuer_page() -> str:
    conn = open_db()
    if request.method == "POST":
        add_issuer(conn, request.form.get("company_name", ""), request.form.get("ico", ""), request.form.get("dic", ""), request.form.get("vat_payer") == "on", request.form.get("address", ""), request.form.get("email", ""), request.form.get("phone", ""), request.form.get("bank_account", ""), request.form.get("bank_code", ""), request.form.get("iban", ""), request.form.get("swift", ""), request.form.get("is_default") == "on")
        flash("Firma byla uložena.", "info")
        return redirect(url_for("issuers_page"))

    body = """
    <section class="content-card"><div class="section-head"><div><h2>Nová fakturující firma</h2><div class="section-copy">Dodavatel, bankovní údaje a režim DPH.</div></div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název firmy</label><input id="issuer-name" name="company_name" required></div><div class="field"><label>E-mail</label><input name="email"></div><div class="field"><label>IČO</label><div class="lookup-row"><input id="issuer-ico" name="ico"><button class="button-secondary" type="button" onclick="lookupAres('issuer')">Načíst z ARES</button></div><div class="lookup-status" id="issuer-ares-status"></div></div><div class="field"><label>DIČ</label><input id="issuer-dic" name="dic"></div><div class="field"><label>Telefon</label><input name="phone"></div><div class="field"><label>Adresa</label><textarea id="issuer-address" name="address" rows="3"></textarea></div><div class="field"><label>Číslo účtu</label><input name="bank_account"></div><div class="field"><label>Kód banky</label><input name="bank_code"></div><div class="field"><label>IBAN</label><input name="iban"></div><div class="field"><label>SWIFT</label><input name="swift"></div></div><div class="toolbar"><label><input id="issuer-vat-payer" type="checkbox" name="vat_payer" checked> Plátce DPH</label><label><input type="checkbox" name="is_default"> Výchozí firma</label></div><div class="toolbar"><button class="button" type="submit">Uložit firmu</button><a class="button-link" href="{{ url_for('issuers_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Nová firma", "Samostatná stránka pro založení fakturující firmy.", body, "issuers", ares_script=ARES_LOOKUP_SCRIPT)


@app.route("/issuers/<int:issuer_id>", methods=["GET", "POST"])
def edit_issuer_page(issuer_id: int) -> str:
    conn = open_db()
    row = next((item for item in list_issuers(conn) if int(item["id"]) == issuer_id), None)
    if row is None:
        flash("Firma nebyla nalezena.", "error")
        return redirect(url_for("issuers_page"))
    if request.method == "POST":
        update_issuer(conn, issuer_id, request.form.get("company_name", ""), request.form.get("ico", ""), request.form.get("dic", ""), request.form.get("vat_payer") == "on", request.form.get("address", ""), request.form.get("email", ""), request.form.get("phone", ""), request.form.get("bank_account", ""), request.form.get("bank_code", ""), request.form.get("iban", ""), request.form.get("swift", ""), request.form.get("is_default") == "on")
        flash("Firma byla aktualizována.", "info")
        return redirect(url_for("issuers_page"))
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Upravit firmu</h2><div class="section-copy">Editace dodavatele, DPH a bankovních údajů.</div></div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název firmy</label><input id="issuer-name" name="company_name" value="{{ row.company_name }}" required></div><div class="field"><label>E-mail</label><input name="email" value="{{ row.email or '' }}"></div><div class="field"><label>IČO</label><div class="lookup-row"><input id="issuer-ico" name="ico" value="{{ row.ico or '' }}"><button class="button-secondary" type="button" onclick="lookupAres('issuer')">Načíst z ARES</button></div><div class="lookup-status" id="issuer-ares-status"></div></div><div class="field"><label>DIČ</label><input id="issuer-dic" name="dic" value="{{ row.dic or '' }}"></div><div class="field"><label>Telefon</label><input name="phone" value="{{ row.phone or '' }}"></div><div class="field"><label>Adresa</label><textarea id="issuer-address" name="address" rows="3">{{ row.address or '' }}</textarea></div><div class="field"><label>Číslo účtu</label><input name="bank_account" value="{{ row.bank_account or '' }}"></div><div class="field"><label>Kód banky</label><input name="bank_code" value="{{ row.bank_code or '' }}"></div><div class="field"><label>IBAN</label><input name="iban" value="{{ row.iban or '' }}"></div><div class="field"><label>SWIFT</label><input name="swift" value="{{ row.swift or '' }}"></div></div><div class="toolbar"><label><input id="issuer-vat-payer" type="checkbox" name="vat_payer" {% if row.vat_payer %}checked{% endif %}> Plátce DPH</label><label><input type="checkbox" name="is_default" {% if row.is_default %}checked{% endif %}> Výchozí firma</label></div><div class="toolbar"><button class="button" type="submit">Uložit změny</button><a class="button-link" href="{{ url_for('issuers_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Upravit firmu", "Editace fakturující firmy ve webovém rozhraní.", body, "issuers", row=row, ares_script=ARES_LOOKUP_SCRIPT)


@app.post("/issuers/<int:issuer_id>/delete")
def delete_issuer_page(issuer_id: int):
    conn = open_db()
    deleted, message = delete_issuer_if_unused(conn, issuer_id)
    if deleted:
        flash("Firma byla smazána.", "info")
    else:
        flash(message or "Firmu nelze smazat.", "error")
    return redirect(url_for("issuers_page"))

@app.route("/invoices")
def invoices_page() -> str:
    conn = open_db()
    rows = []
    for row in list_invoices(conn):
        totals = compute_totals(conn, int(row["id"]))
        issue_date = parse_date(str(row["issue_date"]), default=date.today())
        rows.append(
            {
                "id": row["id"],
                "number": row["invoice_number"],
                "issuer": row["issuer_name"] or "-",
                "customer": row["customer_name"],
                "issue": format_date_cz(row["issue_date"]),
                "issue_sort": str(row["issue_date"]),
                "issue_year": issue_date.year,
                "due": format_date_cz(row["due_date"]),
                "status": invoice_status_cell(str(row["status"]), int(row["id"])),
                "total": f"{totals.grand_total:.2f} {row['currency']}",
                "note": row["note"] or "",
            }
        )
    year_groups: dict[int, list[dict[str, Any]]] = {}
    for row in rows:
        year_groups.setdefault(int(row["issue_year"]), []).append(row)
    for year, year_rows in year_groups.items():
        year_rows.sort(
            key=lambda item: (
                parse_date(str(item["issue_sort"]), default=date.today()),
                str(item["number"] or ""),
                int(item["id"]),
            ),
            reverse=True,
        )
    grouped_rows = [
        {
            "year": year,
            "rows": year_groups[year],
            "open": year == date.today().year,
            "total_amount": sum(parse_amount_text(item["total"]) for item in year_groups[year]),
        }
        for year in sorted(year_groups.keys(), reverse=True)
    ]
    body = """
    <section class="table-card">
      <div class="section-head">
        <div><h2>Databáze faktur</h2><div class="section-copy">Všechny vystavené faktury na jednom místě.</div></div>
        <div class="toolbar"><a class="button-link" href="{{ url_for('customers_page') }}">Odběratelé</a><a class="button-link" href="{{ url_for('data_tools_page') }}">Import / Export</a><a class="button-secondary" href="{{ url_for('import_invoice_pdf_page') }}">Import z PDF</a><a class="button" href="{{ url_for('new_invoice_page') }}">Nová faktura</a></div>
      </div>
      <div class="stack">
        {% for group in grouped_rows %}
        <details class="content-card" {% if group.open %}open{% endif %} style="padding:0;background:var(--panel-2);">
          <summary style="list-style:none;cursor:pointer;padding:18px 20px;display:flex;justify-content:space-between;align-items:center;font-weight:700;">
            <span>Rok {{ group.year }}</span>
            <div style="display:flex;gap:18px;align-items:center;">
              <span class="muted">{{ '%.2f'|format(group.total_amount) }} Kč</span>
              <span class="muted">{{ group.rows|length }} faktur</span>
            </div>
          </summary>
          <div style="padding:0 14px 14px 14px;">
            <table>
              <thead><tr><th>Faktura</th><th>Firma</th><th>Odběratel</th><th>Vystavení</th><th>Splatnost</th><th>Stav</th><th>Celkem</th></tr></thead>
              <tbody>
                {% for row in group.rows %}
                <tr>
                  <td><a href="{{ url_for('invoice_detail_page', invoice_id=row.id) }}"><strong>{{ row.number }}</strong></a><br><span class="muted">{{ row.note[:80] }}</span></td>
                  <td>{{ row.issuer }}</td>
                  <td>{{ row.customer }}</td>
                  <td>{{ row.issue }}</td>
                  <td>{{ row.due }}</td>
                  <td>{{ row.status|safe }}</td>
                  <td>{{ row.total }}</td>
                </tr>
                {% endfor %}
              </tbody>
            </table>
          </div>
        </details>
        {% endfor %}
      </div>
    </section>
    <script>
      document.addEventListener('DOMContentLoaded', function () {
        document.querySelectorAll('.invoice-paid-form').forEach(function(form) {
          form.addEventListener('submit', async function(event) {
            event.preventDefault();
            const button = form.querySelector('button');
            const confirmed = window.confirm('Označit fakturu jako zaplacenou?');
            if (!confirmed) return;
            const originalHtml = form.innerHTML;
            if (button) {
              button.disabled = true;
              button.textContent = 'Ukládám...';
            }
            try {
              const response = await fetch(form.action, {
                method: 'POST',
                headers: {
                  'X-Requested-With': 'fetch',
                  'Accept': 'application/json'
                }
              });
              if (!response.ok) throw new Error('request_failed');
              const data = await response.json();
              form.outerHTML = data.status_html;
            } catch (error) {
              form.innerHTML = originalHtml;
              alert('Označení faktury jako zaplacené se nepodařilo.');
            }
          });
        });
      });
    </script>
    """
    return render_page("Příjmy / Faktury", "Přehled vystavených faktur.", body, "invoices", grouped_rows=grouped_rows)


@app.route("/expenses")
def expenses_page() -> str:
    filters = {
        "date_from": request.args.get("date_from", "").strip(),
        "date_to": request.args.get("date_to", "").strip(),
        "category": request.args.get("category", "").strip(),
        "supplier_name": request.args.get("supplier_name", "").strip(),
        "status": request.args.get("status", "").strip(),
        "project_code": request.args.get("project_code", "").strip(),
        "query": request.args.get("query", "").strip(),
    }
    try:
        conn = open_db()
        try_ensure_recurring_expenses(conn)
        expenses = list_expenses(conn, **filters)
        rows = []
        for row in expenses:
            meta = expense_review_meta(row)
            rows.append(
                {
                    **dict(row),
                    "expense_date_raw": str(row["expense_date"] or ""),
                    "expense_date": format_date_cz(row["expense_date"]),
                    "status_html": status_pill(normalize_expense_status(row["status"], bool(row["attachment_verified"]), bool(row["price_confirmed"]))),
                    "amount_text": f"{float(row['amount'] or 0):.2f} {row['currency'] or 'CZK'}",
                    "recurring_text": meta["period_label"] if meta["recurring"] else "-",
                    "review_text": "Doklad i cena ověřeny" if meta["review_complete"] else ("Chybí PDF doklad" if not meta["attachment_verified"] else "Čeká na potvrzení ceny"),
                    "attachment_label": meta["attachment_label"],
                }
            )
        yearly = yearly_expense_overview(conn)
        categories = expense_category_options(conn)
        total = sum(float(row["amount"] or 0) for row in expenses)
    except sqlite3.OperationalError as exc:
        if "locked" not in str(exc).lower():
            raise
        flash("Databáze výdajů je právě používaná jiným procesem. Spusť aplikaci znovu přes spouštěcí soubor.", "error")
        rows = []
        yearly = []
        categories = ["Ostatní"]
        total = 0.0
    grouped_expenses: list[dict[str, Any]] = []
    year_bucket: dict[int, dict[str, Any]] = {}
    current_year = date.today().year
    for row in rows:
        raw_date = str(row.get("expense_date_raw", "")).strip()
        try:
            expense_year = parse_date(raw_date, default=date.today()).year
        except Exception:
            expense_year = current_year
        category_name = str(row.get("category") or "Bez kategorie").strip() or "Bez kategorie"
        if expense_year not in year_bucket:
            year_bucket[expense_year] = {
                "year": expense_year,
                "open": expense_year == current_year,
                "categories": {},
                "total_amount": 0.0,
                "count": 0,
            }
        bucket = year_bucket[expense_year]
        if category_name not in bucket["categories"]:
            bucket["categories"][category_name] = {
                "name": category_name,
                "rows": [],
                "total_amount": 0.0,
                "count": 0,
                "open": expense_year == current_year,
            }
        cat_bucket = bucket["categories"][category_name]
        amount_value = parse_amount_text(row.get("amount_text", "0"))
        cat_bucket["rows"].append(row)
        cat_bucket["total_amount"] += amount_value
        cat_bucket["count"] += 1
        bucket["total_amount"] += amount_value
        bucket["count"] += 1
    grouped_expenses = [
        {
            "year": year,
            "open": year_bucket[year]["open"],
            "total_amount": year_bucket[year]["total_amount"],
            "count": year_bucket[year]["count"],
            "categories": sorted(year_bucket[year]["categories"].values(), key=lambda item: str(item["name"]).lower()),
        }
        for year in sorted(year_bucket.keys(), reverse=True)
    ]
    document_types = ["Přijatá faktura", "Paragon", "Účtenka", "Interní výdaj"]
    payment_methods = ["Hotově", "Kartou", "Převodem"]
    body = """
    <style>
    .expense-strip-list{display:grid;gap:10px}
    .expense-strip{display:grid;grid-template-columns:116px minmax(0,1.7fr) 128px 180px 244px 176px;gap:16px;align-items:center;min-height:84px;padding:12px 14px;border:1px solid var(--line);border-radius:16px;background:#0b1917}
    .expense-strip strong{display:block}
    .expense-date,.expense-main,.expense-status,.expense-symbol{min-width:0}
    .expense-date .muted,.expense-main .muted,.expense-status .muted,.expense-symbol .muted{display:block;margin-top:4px;font-size:12px;line-height:1.32}
    .expense-date strong,.expense-main strong,.expense-symbol strong,.expense-amount{font-size:15px}
    .expense-date strong,.expense-main strong,.expense-main .muted,.expense-symbol strong{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .expense-status{display:grid;gap:6px;align-content:center}
    .expense-status .status-pill,.expense-status .badge,.expense-status .pill{width:100%;justify-content:center;text-align:center}
    .expense-status > *:first-child{justify-self:stretch}
    .expense-status .muted{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;min-height:32px}
    .expense-symbol{display:grid;gap:4px;align-content:center}.expense-symbol strong{font-size:14px}.expense-symbol .muted{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}.expense-amount{font-weight:800;white-space:nowrap;text-align:right}
    .expense-actions{justify-content:flex-end;flex-wrap:nowrap;gap:10px}
    .expense-actions a,.expense-actions button{min-width:88px;text-align:center}
    .expense-actions form{margin:0}
    @media (max-width: 1100px){
      .expense-strip{grid-template-columns:110px 1fr 132px;grid-template-areas:"date main amount" "date status symbol" "date actions actions"}
      .expense-date{grid-area:date}
      .expense-main{grid-area:main}
      .expense-status{grid-area:status}
      .expense-symbol{grid-area:symbol}
      .expense-amount{grid-area:amount;justify-self:end}
      .expense-actions{grid-area:actions;justify-content:flex-start}
    }
    @media (max-width: 760px){
      .expense-strip{grid-template-columns:1fr;grid-template-areas:none}
      .expense-date,.expense-main,.expense-status,.expense-symbol,.expense-amount,.expense-actions{grid-area:auto}
      .expense-amount{justify-self:start}
      .expense-actions{justify-content:flex-start;flex-wrap:wrap}
      .expense-main strong,.expense-amount,.expense-date strong{white-space:normal;overflow:visible;text-overflow:clip}
    }
    </style>
    <section class="content-card"><div class="toolbar"><a class="button" href="{{ url_for('new_expense_page') }}">Nový výdaj</a><a class="button-secondary" href="{{ url_for('import_expense_pdf_page') }}">Import z PDF</a><a class="button-link" href="{{ url_for('expense_suppliers_page') }}">Dodavatelé</a><a class="button-link" href="{{ url_for('expense_categories_page') }}">Kategorie</a></div></section>
    <details class="content-card"><summary style="list-style:none;cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-weight:700;"><span>Filtry a hledání</span><span class="muted">Rozbalit</span></summary><div style="padding-top:16px;"><div class="section-copy" style="margin-bottom:14px;">Najdi výdaje podle data, kategorie, dodavatele, projektu nebo stavu.</div><form method="get" class="form-grid-3"><div class="field"><label>Od data</label><input type="date" name="date_from" value="{{ filters.date_from }}"></div><div class="field"><label>Do data</label><input type="date" name="date_to" value="{{ filters.date_to }}"></div><div class="field"><label>Kategorie</label><select name="category"><option value="">Všechny kategorie</option>{% for item in categories %}<option value="{{ item }}" {% if filters.category == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Dodavatel</label><input name="supplier_name" value="{{ filters.supplier_name }}"></div><div class="field"><label>Stav</label><input name="status" value="{{ filters.status }}"></div><div class="field"><label>Projekt</label><input name="project_code" value="{{ filters.project_code }}"></div><div class="field" style="grid-column:1 / -1;"><label>Hledat</label><input name="query" value="{{ filters.query }}" placeholder="Název, poznámka, číslo dokladu, dodavatel"></div><div class="toolbar"><button class="button" type="submit">Filtrovat</button><a class="button-link" href="{{ url_for('expenses_page') }}">Reset</a></div></form></div></details>
    <section class="table-card"><div class="section-head"><div><h2>Seznam výdajů</h2><div class="section-copy">Výdaje jsou rozdělené podle roku a uvnitř ještě podle kategorií.</div></div></div><div class="stack">{% for year_group in grouped_expenses %}<details class="content-card" {% if year_group.open %}open{% endif %} style="padding:0;background:var(--panel-2);"><summary style="list-style:none;cursor:pointer;padding:18px 20px;display:flex;justify-content:space-between;align-items:center;font-weight:700;"><span>Rok {{ year_group.year }}</span><div style="display:flex;gap:18px;align-items:center;"><span class="muted">{{ '%.2f'|format(year_group.total_amount) }} Kč</span><span class="muted">{{ year_group.count }} výdajů</span></div></summary><div style="padding:0 14px 14px 14px;" class="stack">{% for category_group in year_group.categories %}<details class="content-card" {% if category_group.open %}open{% endif %} style="padding:0;background:#10211e;"><summary style="list-style:none;cursor:pointer;padding:14px 16px;display:flex;justify-content:space-between;align-items:center;font-weight:700;"><span>{{ category_group.name }}</span><div style="display:flex;gap:16px;align-items:center;"><span class="muted">{{ '%.2f'|format(category_group.total_amount) }} Kč</span><span class="muted">{{ category_group.count }} položek</span></div></summary><div style="padding:0 10px 10px 10px;"><div class="expense-strip-list">{% for row in category_group.rows %}<article class="expense-strip"><div class="expense-date"><strong>{{ row.expense_date }}</strong><span class="muted">{{ row.category or '-' }}</span></div><div class="expense-main"><strong>{{ row.title }}</strong><span class="muted">{{ row.supplier_name or '-' }}{% if row.recurring_text and row.recurring_text != 'Neopakovat' %} · {{ row.recurring_text }}{% endif %}</span></div><div class="expense-amount">{{ row.amount_text }}</div><div class="expense-symbol"><strong>{% if row.document_number %}{{ row.document_number }}{% else %}-{% endif %}</strong><span class="muted">{% if row.variable_symbol %}Var. symbol: {{ row.variable_symbol }}{% else %}Var. symbol: -{% endif %}</span></div><div class="expense-status">{{ row.status_html|safe }}<span class="muted">{{ row.review_text }} · {{ row.attachment_label }}</span></div><div class="toolbar expense-actions"><a class="button-link" href="{{ url_for('edit_expense_page', expense_id=row.id) }}">Upravit</a><form method="post" action="{{ url_for('delete_expense_page', expense_id=row.id) }}" onsubmit="return confirm('Opravdu smazat výdaj?');"><button class="button-danger" type="submit">Smazat</button></form></div></article>{% endfor %}</div></div></details>{% endfor %}</div></details>{% endfor %}{% if not grouped_expenses %}<div class="muted">Zatím nejsou evidované žádné výdaje.</div>{% endif %}</div></section>
    """
    return render_page("Výdaje", "Zjednodušená evidence výdajů s načtením z PDF a opakováním.", body, "expenses", expenses=rows, grouped_expenses=grouped_expenses, total=f"{total:.2f} Kč", yearly=yearly, today=date.today().isoformat(), filters=filters, categories=categories, document_types=document_types, payment_methods=payment_methods, expense_statuses=expense_status_options(), recurring_periods=recurring_period_options())


@app.route("/expense-suppliers")
def expense_suppliers_page() -> str:
    conn = open_db()
    supplier_rows = conn.execute(
        """
        SELECT
            trim(coalesce(supplier_name, '')) AS supplier_name,
            MAX(nullif(trim(coalesce(supplier_ico, '')), '')) AS supplier_ico,
            MAX(nullif(trim(coalesce(supplier_dic, '')), '')) AS supplier_dic,
            COUNT(*) AS expense_count,
            COALESCE(SUM(amount), 0) AS total_amount,
            SUM(CASE WHEN trim(coalesce(document_number, '')) <> '' THEN 1 ELSE 0 END) AS document_count,
            SUM(CASE WHEN trim(coalesce(attachment_path, '')) <> '' THEN 1 ELSE 0 END) AS attachment_count
        FROM expenses
        WHERE trim(coalesce(supplier_name, '')) <> ''
        GROUP BY trim(coalesce(supplier_name, ''))
        ORDER BY lower(trim(coalesce(supplier_name, ''))) ASC
        """
    ).fetchall()
    body = """
    <section class="table-card">
      <style>
        .suppliers-table{table-layout:fixed}
        .suppliers-table th,.suppliers-table td{padding:6px 8px;line-height:1.18}
        .suppliers-table th{font-size:12px}
        .suppliers-table .name-cell strong{display:block;font-size:16px;line-height:1.1;margin-bottom:2px}
        .suppliers-table .name-meta{display:block;font-size:13px;line-height:1.18;color:var(--muted)}
        .suppliers-table .compact-cell{font-size:13px}
        .suppliers-table .summary{display:grid;gap:4px}
        .suppliers-table .count-link,
        .suppliers-table .count-empty{display:inline-flex;min-width:34px;justify-content:center;align-items:center;padding:3px 7px;border-radius:999px;font-weight:800;justify-self:start;font-size:12px}
        .suppliers-table .count-link{background:var(--accent-soft);color:var(--accent);border:1px solid var(--line)}
        .suppliers-table .count-empty{background:var(--panel-2);color:var(--muted);border:1px solid var(--line)}
        .suppliers-table .summary-total{font-size:11px;line-height:1.15;color:var(--muted);white-space:nowrap}
        .suppliers-table .presence-icons{display:flex;gap:4px;align-items:center;flex-wrap:wrap}
        .suppliers-table .presence-dot{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:999px;border:1px solid var(--line);font-size:12px;font-weight:800}
        .suppliers-table .presence-dot.is-on{background:#123327;color:#86efac}
        .suppliers-table .presence-dot.is-off{background:var(--panel-2);color:var(--muted)}
        .suppliers-table .toolbar{gap:5px;flex-wrap:nowrap}
        .suppliers-table .toolbar form{margin:0}
        .suppliers-table .button-link{padding:6px 8px;border-radius:12px;font-size:12px;white-space:nowrap}
      </style>
      <div class="section-head">
        <div><h2>Dodavatelé</h2><div class="section-copy">Přehled dodavatelů převzatých z výdajů.</div></div>
        <a class="button-link" href="{{ url_for('expenses_page') }}">Zpět na výdaje</a>
      </div>
      <table class="suppliers-table">
        <thead>
          <tr>
            <th style="width:44%;">Dodavatel</th>
            <th style="width:12%;">Vyplnění</th>
            <th style="width:12%;">Doklady</th>
            <th style="width:20%;">Akce</th>
          </tr>
        </thead>
        <tbody>
          {% for row in rows %}
          <tr>
            <td class="name-cell">
              <strong>{{ row.supplier_name }}</strong>
              <span class="name-meta">{{ row.supplier_ico or 'Bez IČO' }}</span>
            </td>
            <td class="compact-cell">
              <div class="presence-icons">
                <span class="presence-dot {{ 'is-on' if row.supplier_ico else 'is-off' }}" title="IČO {{ 'vyplněné' if row.supplier_ico else 'chybí' }}">I</span>
                <span class="presence-dot {{ 'is-on' if row.supplier_dic else 'is-off' }}" title="DIČ {{ 'vyplněné' if row.supplier_dic else 'chybí' }}">D</span>
                <span class="presence-dot {{ 'is-on' if row.attachment_count else 'is-off' }}" title="PDF doklad {{ 'existuje' if row.attachment_count else 'chybí' }}">P</span>
              </div>
            </td>
            <td class="compact-cell">
              <div class="summary">
                <a class="count-link" href="{{ url_for('expenses_page', supplier_name=row.supplier_name) }}" title="Zobrazit výdaje dodavatele">{{ row.expense_count }}</a>
                <span class="summary-total">{{ '%.2f'|format(row.total_amount) }} Kč</span>
              </div>
            </td>
            <td>
              <div class="toolbar">
                <a class="button-link" href="{{ url_for('edit_expense_supplier_page') }}?supplier_name={{ row.supplier_name|urlencode }}">Upravit</a>
                <a class="button-link" href="{{ url_for('expenses_page', supplier_name=row.supplier_name) }}">Výdaje</a>
              </div>
            </td>
          </tr>
          {% endfor %}
          {% if not rows %}
          <tr><td colspan="4" class="muted">Zatím nejsou evidovaní žádní dodavatelé u výdajů.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </section>
    """
    return render_page("Dodavatelé", "Seznam dodavatelů z evidovaných výdajů.", body, "expenses", rows=supplier_rows)


@app.route("/expense-suppliers/edit", methods=["GET", "POST"])
def edit_expense_supplier_page() -> str:
    conn = open_db()
    original_supplier_name = request.values.get("supplier_name", "").strip()
    if not original_supplier_name:
        flash("Dodavatel nebyl vybrán.", "error")
        return redirect(url_for("expense_suppliers_page"))

    supplier_row = conn.execute(
        """
        SELECT
            trim(coalesce(supplier_name, '')) AS supplier_name,
            MAX(nullif(trim(coalesce(supplier_ico, '')), '')) AS supplier_ico,
            MAX(nullif(trim(coalesce(supplier_dic, '')), '')) AS supplier_dic,
            COUNT(*) AS expense_count
        FROM expenses
        WHERE trim(coalesce(supplier_name, '')) = ?
        GROUP BY trim(coalesce(supplier_name, ''))
        """,
        (original_supplier_name,),
    ).fetchone()
    if supplier_row is None:
        flash("Dodavatel nebyl nalezen.", "error")
        return redirect(url_for("expense_suppliers_page"))

    if request.method == "POST":
        new_name = request.form.get("name", "").strip()
        new_ico = request.form.get("ico", "").strip()
        new_dic = request.form.get("dic", "").strip()
        if not new_name:
            flash("Název dodavatele je povinný.", "error")
            return redirect(url_for("edit_expense_supplier_page", supplier_name=original_supplier_name))
        conn.execute(
            """
            UPDATE expenses
            SET supplier_name = ?, supplier_ico = ?, supplier_dic = ?
            WHERE trim(coalesce(supplier_name, '')) = ?
            """,
            (new_name, new_ico, new_dic, original_supplier_name),
        )
        conn.commit()
        flash("Profil dodavatele byl upraven ve všech navázaných výdajích.", "info")
        return redirect(url_for("expense_suppliers_page"))

    body = """
    <section class="content-card">
      <div class="section-head">
        <div>
          <h2>Upravit dodavatele</h2>
          <div class="section-copy">Změna se propíše do všech výdajů tohoto dodavatele.</div>
        </div>
      </div>
      <form method="post" class="stack">
        <input type="hidden" name="supplier_name" value="{{ supplier.supplier_name }}">
        <div class="form-grid">
          <div class="field"><label>Název dodavatele</label><input id="customer-name" name="name" value="{{ supplier.supplier_name }}" required></div>
          <div class="field"><label>IČO</label><div class="lookup-row"><input id="customer-ico" name="ico" value="{{ supplier.supplier_ico or '' }}"><button class="button-secondary" type="button" onclick="lookupAres('customer')">Načíst z ARES</button></div><div class="lookup-status" id="customer-ares-status"></div></div>
          <div class="field"><label>DIČ</label><input id="customer-dic" name="dic" value="{{ supplier.supplier_dic or '' }}"></div>
          <div class="field"><label>Počet výdajů</label><input value="{{ supplier.expense_count }}" disabled></div>
          <div class="field" style="grid-column:1 / -1;"><label>Adresa</label><textarea id="customer-address" rows="3" disabled placeholder="Dodavatel výdajů nemá samostatné pole adresy, ale ARES může pomoct s dohledáním údajů."></textarea></div>
        </div>
        <div class="toolbar">
          <button class="button" type="submit">Uložit změny</button>
          <a class="button-link" href="{{ url_for('expense_suppliers_page') }}">Zpět</a>
        </div>
      </form>
    </section>
    """
    return render_page("Upravit dodavatele", "Hromadná úprava profilu dodavatele ve výdajích.", body, "expenses", supplier=supplier_row, ares_script=ARES_LOOKUP_SCRIPT)


@app.route("/expense-categories", methods=["GET", "POST"])
def expense_categories_page() -> str:
    conn = open_db()
    if request.method == "POST":
        try:
            action = request.form.get("action", "").strip()
            if action == "create":
                add_expense_category(conn, request.form.get("name", ""))
                flash("Kategorie byla přidána.", "info")
            elif action == "update":
                category_id = parse_int_field(request.form.get("category_id", "0"), 0)
                update_expense_category(conn, category_id, request.form.get("name", ""))
                flash("Kategorie byla upravena.", "info")
            elif action == "delete":
                category_id = parse_int_field(request.form.get("category_id", "0"), 0)
                deleted, message = delete_expense_category_if_unused(conn, category_id)
                if deleted:
                    flash("Kategorie byla smazána.", "info")
                else:
                    flash(message or "Kategorii nelze smazat.", "error")
            return redirect(url_for("expense_categories_page"))
        except Exception as exc:
            flash(str(exc), "error")
            return redirect(url_for("expense_categories_page"))

    usage_rows = conn.execute(
        """
        SELECT trim(coalesce(category, '')) AS category_name, COUNT(*) AS usage_count
        FROM expenses
        WHERE trim(coalesce(category, '')) <> ''
        GROUP BY trim(coalesce(category, ''))
        """
    ).fetchall()
    usage_map = {str(row["category_name"]): int(row["usage_count"] or 0) for row in usage_rows}
    rows = []
    for row in list_expense_categories(conn):
        item = dict(row)
        usage_count = usage_map.get(str(row["name"] or "").strip(), 0)
        item["usage_count"] = usage_count
        item["is_used"] = usage_count > 0
        rows.append(item)
    body = """
    <style>
      .category-state{display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:800;border:1px solid var(--line)}
      .category-state.used{background:#3b3014;color:#fde68a}
      .category-state.unused{background:#123327;color:#86efac}
      .button-danger[disabled]{opacity:.45;cursor:not-allowed}
    </style>
    <section class="content-card"><div class="section-head"><div><h2>Nová kategorie</h2><div class="section-copy">Přidej kategorii pro výdaje a filtry.</div></div></div><form method="post" class="toolbar"><input type="hidden" name="action" value="create"><div class="field" style="min-width:280px;flex:1;"><label>Název kategorie</label><input name="name" required></div><div style="align-self:end;"><button class="button" type="submit">Přidat kategorii</button></div></form></section>
    <section class="table-card"><div class="section-head"><div><h2>Seznam kategorií</h2><div class="section-copy">Mazání je povolené jen u nevyužitých kategorií. Použité a nepoužité jsou viditelně odlišené.</div></div></div><table><thead><tr><th>Název</th><th>Stav</th><th>Akce</th></tr></thead><tbody>{% for row in rows %}<tr><td><form method="post" id="category-update-{{ row.id }}"><input type="hidden" name="action" value="update"><input type="hidden" name="category_id" value="{{ row.id }}"><div class="field" style="min-width:280px;"><label>Název</label><input name="name" value="{{ row.name }}" required></div></form></td><td>{% if row.is_used %}<span class="category-state used">Využitá · {{ row.usage_count }}×</span>{% else %}<span class="category-state unused">Nevyužitá</span>{% endif %}</td><td><div class="toolbar" style="flex-wrap:nowrap;"> <button class="button-link" type="submit" form="category-update-{{ row.id }}">Uložit</button>{% if row.is_used %}<button class="button-danger" type="button" disabled title="Použitou kategorii nelze smazat.">Smazat</button>{% else %}<form method="post" onsubmit="return confirm('Opravdu smazat kategorii?');"><input type="hidden" name="action" value="delete"><input type="hidden" name="category_id" value="{{ row.id }}"><button class="button-danger" type="submit">Smazat</button></form>{% endif %}</div></td></tr>{% endfor %}{% if not rows %}<tr><td colspan="3" class="muted">Zatím nejsou založené žádné kategorie.</td></tr>{% endif %}</tbody></table></section>
    """
    return render_page("Kategorie výdajů", "Správa kategorií použitých u výdajů a ve filtrech.", body, "expense-categories", rows=rows)


def collect_year_report_data(conn, selected_year_raw: str = "") -> dict[str, Any]:
    issuers = list_issuers(conn)
    issuer_row = next((row for row in issuers if int(row["is_default"] or 0) == 1), issuers[0] if issuers else None)
    invoice_year_rows = yearly_overview(conn)
    expense_year_rows = yearly_expense_overview(conn)
    available_years = sorted(
        {int(row["year"]) for row in invoice_year_rows if row["year"] is not None}
        | {int(row["year"]) for row in expense_year_rows if row["year"] is not None},
        reverse=True,
    )
    current_year = date.today().year
    if not available_years:
        available_years = [current_year]

    try:
        selected_year = int(selected_year_raw) if str(selected_year_raw).strip() else available_years[0]
    except ValueError:
        selected_year = available_years[0]
    if selected_year not in available_years:
        selected_year = available_years[0]

    income_rows: list[dict[str, Any]] = []
    income_total = 0.0
    for row in list_invoices(conn):
        issue_date = parse_date(str(row["issue_date"]), default=date.today())
        if issue_date.year != selected_year:
            continue
        totals = compute_totals(conn, int(row["id"]))
        total = float(totals.grand_total)
        income_total += total
        income_rows.append(
            {
                "id": int(row["id"]),
                "number": str(row["invoice_number"] or ""),
                "customer": str(row["customer_name"] or "-"),
                "issue": format_date_cz(row["issue_date"]),
                "status": status_pill(str(row["status"] or "")),
                "total": total,
                "currency": str(row["currency"] or "CZK"),
                "month": issue_date.month,
            }
        )

    expense_rows: list[dict[str, Any]] = []
    expense_total = 0.0
    for row in list_expenses(conn):
        expense_date = parse_date(str(row["expense_date"]), default=date.today())
        if expense_date.year != selected_year:
            continue
        amount = float(row["amount"] or 0.0)
        expense_total += amount
        expense_rows.append(
            {
                "id": int(row["id"]),
                "title": str(row["title"] or "-"),
                "supplier_name": str(row["supplier_name"] or "-"),
                "expense_date": format_date_cz(row["expense_date"]),
                "category": str(row["category"] or "-"),
                "status": status_pill(normalize_expense_status(str(row["status"] or ""), bool(row["attachment_verified"]), bool(row["price_confirmed"]))),
                "amount": amount,
                "currency": str(row["currency"] or "CZK"),
                "attachment_path": str(row["attachment_path"] or ""),
                "month": expense_date.month,
            }
        )

    month_labels = ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"]
    income_month_totals = {month: 0.0 for month in range(1, 13)}
    expense_month_totals = {month: 0.0 for month in range(1, 13)}
    for row in income_rows:
        income_month_totals[int(row["month"])] += float(row["total"])
    for row in expense_rows:
        expense_month_totals[int(row["month"])] += float(row["amount"])
    month_rows = [
        {
            "label": month_labels[month - 1],
            "income": income_month_totals[month],
            "expense": expense_month_totals[month],
            "profit": income_month_totals[month] - expense_month_totals[month],
        }
        for month in range(1, 13)
    ]

    return {
        "issuer": {
            "company_name": str(issuer_row["company_name"] or "-") if issuer_row else "-",
            "address": str(issuer_row["address"] or "-") if issuer_row else "-",
            "ico": str(issuer_row["ico"] or "-") if issuer_row else "-",
            "dic": str(issuer_row["dic"] or "-") if issuer_row else "-",
        },
        "available_years": available_years,
        "selected_year": selected_year,
        "income_rows": income_rows,
        "expense_rows": expense_rows,
        "income_total": income_total,
        "expense_total": expense_total,
        "profit_total": income_total - expense_total,
        "month_rows": month_rows,
    }


@app.get("/year-report")
def year_report_page() -> str:
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    body = """
    <section class="content-card">
      <div class="section-head">
        <div>
          <h2>Roční přehled</h2>
          <div class="section-copy">Vyber rok a porovnej na jedné stránce příjmy, výdaje a čistý zisk.</div>
        </div>
        <div class="toolbar">
          <a class="button-secondary" target="_blank" href="{{ url_for('year_report_print_page', year=selected_year) }}">Tisk A4</a>
          <a class="button-link" href="{{ url_for('year_report_print_invoices_page', year=selected_year) }}">Výpis příjmy</a>
          <a class="button-link" href="{{ url_for('year_report_print_expense_docs_page', year=selected_year) }}">Výpis výdaje</a>
        </div>
      </div>
      <form method="get" class="toolbar">
        <div class="field" style="min-width:220px;">
          <label>Rok</label>
          <select name="year" onchange="this.form.submit()">
            {% for year in available_years %}
            <option value="{{ year }}" {% if year == selected_year %}selected{% endif %}>{{ year }}</option>
            {% endfor %}
          </select>
        </div>
        <noscript><button class="button" type="submit">Zobrazit</button></noscript>
      </form>
    </section>

    <section class="grid-3">
      <div class="card">
        <div class="metric-label">Příjmy za rok {{ selected_year }}</div>
        <div class="metric-value">{{ '%.2f'|format(income_total) }} Kč</div>
        <div class="metric-sub">{{ income_rows|length }} faktur</div>
      </div>
      <div class="card">
        <div class="metric-label">Výdaje za rok {{ selected_year }}</div>
        <div class="metric-value">{{ '%.2f'|format(expense_total) }} Kč</div>
        <div class="metric-sub">{{ expense_rows|length }} výdajů</div>
      </div>
      <div class="card">
        <div class="metric-label">Zisk za rok {{ selected_year }}</div>
        <div class="metric-value">{{ '%.2f'|format(profit_total) }} Kč</div>
        <div class="metric-sub">Zisk = příjmy - výdaje</div>
      </div>
    </section>

    <section class="stack">
      <details class="table-card">
        <summary style="list-style:none;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:16px;">
          <div>
            <h2 style="margin:0;">Výdaje</h2>
            <div class="section-copy">Přehled všech výdajů za vybraný rok.</div>
          </div>
          <div class="metric-sub">{{ '%.2f'|format(expense_total) }} Kč</div>
        </summary>
        <div style="margin-top:14px;">
          <table>
            <thead><tr><th>Doklad</th><th>Dodavatel</th><th>Datum</th><th>Kategorie</th><th>Stav</th><th>Částka</th></tr></thead>
            <tbody>
              {% for row in expense_rows %}
              <tr>
                <td><a href="{{ url_for('edit_expense_page', expense_id=row.id) }}"><strong>{{ row.title }}</strong></a></td>
                <td>{{ row.supplier_name }}</td>
                <td>{{ row.expense_date }}</td>
                <td>{{ row.category }}</td>
                <td>{{ row.status|safe }}</td>
                <td>{{ '%.2f'|format(row.amount) }} {{ row.currency }}</td>
              </tr>
              {% endfor %}
              {% if not expense_rows %}
              <tr><td colspan="6" class="muted">V tomto roce nejsou evidované žádné výdaje.</td></tr>
              {% endif %}
            </tbody>
          </table>
        </div>
      </details>

      <details class="table-card">
        <summary style="list-style:none;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:16px;">
          <div>
            <h2 style="margin:0;">Příjmy</h2>
            <div class="section-copy">Přehled všech příjmů za vybraný rok.</div>
          </div>
          <div class="metric-sub">{{ '%.2f'|format(income_total) }} Kč</div>
        </summary>
        <div style="margin-top:14px;">
          <table>
            <thead><tr><th>Faktura</th><th>Odběratel</th><th>Vystavení</th><th>Stav</th><th>Částka</th></tr></thead>
            <tbody>
              {% for row in income_rows %}
              <tr>
                <td><a href="{{ url_for('invoice_detail_page', invoice_id=row.id) }}"><strong>{{ row.number }}</strong></a></td>
                <td>{{ row.customer }}</td>
                <td>{{ row.issue }}</td>
                <td>{{ row.status|safe }}</td>
                <td>{{ '%.2f'|format(row.total) }} {{ row.currency }}</td>
              </tr>
              {% endfor %}
              {% if not income_rows %}
              <tr><td colspan="5" class="muted">V tomto roce nejsou evidované žádné příjmy.</td></tr>
              {% endif %}
            </tbody>
          </table>
        </div>
      </details>
    </section>
    """
    return render_page(
        "Roční přehled",
        "Souhrn příjmů, výdajů a zisku s výběrem roku.",
        body,
        "year-report",
        **data,
    )


@app.get("/year-report/print-invoices")
def year_report_print_invoices_page():
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    selected_year = int(data["selected_year"])
    income_rows = list(data["income_rows"])
    if not income_rows:
        flash("V tomto roce nejsou žádné faktury k tisku.", "error")
        return redirect(url_for("year_report_page", year=selected_year))
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        flash("Chybí knihovna pro sloučení PDF: pip install pypdf", "error")
        return redirect(url_for("year_report_page", year=selected_year))

    writer = PdfWriter()
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_base = Path(temp_dir)
        for row in income_rows:
            invoice_id = int(row["id"])
            invoice, items, _ = get_invoice_detail(conn, invoice_id)
            invoice_number = str(invoice["invoice_number"] or f"invoice_{invoice_id}")
            service_name = str(items[0]["description"]) if items else "sluzba"
            forbidden = '<>:"/\\|?*'
            service_safe = "".join("_" if ch in forbidden else ch for ch in service_name).strip()
            service_safe = "_".join(service_safe.split()) or "sluzba"
            invoice_path = temp_base / f"{invoice_number}_{invoice_id}_{service_safe}.pdf"
            export_invoice_pdf(conn, invoice_id, invoice_path)
            reader = PdfReader(str(invoice_path))
            for page in reader.pages:
                writer.add_page(page)

        with tempfile.NamedTemporaryFile(delete=False, suffix=f"_faktury_{selected_year}.pdf") as temp_file:
            output_path = Path(temp_file.name)
        with output_path.open("wb") as merged_file:
            writer.write(merged_file)

    return send_file(output_path, as_attachment=True, download_name=f"faktury_{selected_year}.pdf")


@app.get("/year-report/print-expenses")
def year_report_print_expense_docs_page():
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    selected_year = int(data["selected_year"])
    expense_rows = list(data["expense_rows"])
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        flash("Chybí knihovna pro sloučení PDF: pip install pypdf", "error")
        return redirect(url_for("year_report_page", year=selected_year))

    writer = PdfWriter()
    added_count = 0
    for row in expense_rows:
        attachment_path = str(row.get("attachment_path") or "").strip()
        if not attachment_path:
            continue
        source_path = Path(attachment_path)
        if not source_path.exists() or source_path.suffix.lower() != ".pdf":
            continue
        try:
            reader = PdfReader(str(source_path))
            for page in reader.pages:
                writer.add_page(page)
            added_count += 1
        except Exception:
            continue

    if added_count == 0:
        flash("V tomto roce nejsou žádné PDF doklady výdajů k tisku.", "error")
        return redirect(url_for("year_report_page", year=selected_year))

    with tempfile.NamedTemporaryFile(delete=False, suffix=f"_doklady_vydaju_{selected_year}.pdf") as temp_file:
        output_path = Path(temp_file.name)
    with output_path.open("wb") as merged_file:
        writer.write(merged_file)
    return send_file(output_path, as_attachment=True, download_name=f"doklady_vydaju_{selected_year}.pdf")


@app.get("/year-report/income")
def year_report_income_page() -> str:
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    body = """
    <section class="content-card">
      <div class="section-head">
        <div>
          <h2>Soupis příjmů</h2>
          <div class="section-copy">Samostatný seznam všech příjmů za vybraný rok.</div>
        </div>
        <div class="toolbar">
          <a class="button-link" href="{{ url_for('year_report_page', year=selected_year) }}">Zpět na roční přehled</a>
        </div>
      </div>
      <form method="get" class="toolbar">
        <div class="field" style="min-width:220px;">
          <label>Rok</label>
          <select name="year" onchange="this.form.submit()">
            {% for year in available_years %}
            <option value="{{ year }}" {% if year == selected_year %}selected{% endif %}>{{ year }}</option>
            {% endfor %}
          </select>
        </div>
      </form>
    </section>

    <section class="grid-3">
      <div class="card">
        <div class="metric-label">Příjmy za rok {{ selected_year }}</div>
        <div class="metric-value">{{ '%.2f'|format(income_total) }} Kč</div>
        <div class="metric-sub">{{ income_rows|length }} faktur</div>
      </div>
    </section>

    <section class="table-card">
      <div class="section-head">
        <div>
          <h2>Příjmy</h2>
          <div class="section-copy">Detailní seznam všech faktur za rok {{ selected_year }}.</div>
        </div>
      </div>
      <table>
        <thead><tr><th>Faktura</th><th>Odběratel</th><th>Vystavení</th><th>Stav</th><th>Částka</th></tr></thead>
        <tbody>
          {% for row in income_rows %}
          <tr>
            <td><a href="{{ url_for('invoice_detail_page', invoice_id=row.id) }}"><strong>{{ row.number }}</strong></a></td>
            <td>{{ row.customer }}</td>
            <td>{{ row.issue }}</td>
            <td>{{ row.status|safe }}</td>
            <td>{{ '%.2f'|format(row.total) }} {{ row.currency }}</td>
          </tr>
          {% endfor %}
          {% if not income_rows %}
          <tr><td colspan="5" class="muted">V tomto roce nejsou evidované žádné příjmy.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </section>
    """
    return render_page(
        "Soupis příjmů",
        "Přehled všech příjmů za vybraný rok.",
        body,
        "year-report",
        **data,
    )


@app.get("/year-report/expenses")
def year_report_expenses_page() -> str:
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    body = """
    <section class="content-card">
      <div class="section-head">
        <div>
          <h2>Soupis výdajů</h2>
          <div class="section-copy">Samostatný seznam všech výdajů za vybraný rok.</div>
        </div>
        <div class="toolbar">
          <a class="button-link" href="{{ url_for('year_report_page', year=selected_year) }}">Zpět na roční přehled</a>
        </div>
      </div>
      <form method="get" class="toolbar">
        <div class="field" style="min-width:220px;">
          <label>Rok</label>
          <select name="year" onchange="this.form.submit()">
            {% for year in available_years %}
            <option value="{{ year }}" {% if year == selected_year %}selected{% endif %}>{{ year }}</option>
            {% endfor %}
          </select>
        </div>
      </form>
    </section>

    <section class="grid-3">
      <div class="card">
        <div class="metric-label">Výdaje za rok {{ selected_year }}</div>
        <div class="metric-value">{{ '%.2f'|format(expense_total) }} Kč</div>
        <div class="metric-sub">{{ expense_rows|length }} výdajů</div>
      </div>
    </section>

    <section class="table-card">
      <div class="section-head">
        <div>
          <h2>Výdaje</h2>
          <div class="section-copy">Detailní seznam všech výdajů za rok {{ selected_year }}.</div>
        </div>
      </div>
      <table>
        <thead><tr><th>Doklad</th><th>Dodavatel</th><th>Datum</th><th>Kategorie</th><th>Stav</th><th>Částka</th></tr></thead>
        <tbody>
          {% for row in expense_rows %}
          <tr>
            <td><a href="{{ url_for('edit_expense_page', expense_id=row.id) }}"><strong>{{ row.title }}</strong></a></td>
            <td>{{ row.supplier_name }}</td>
            <td>{{ row.expense_date }}</td>
            <td>{{ row.category }}</td>
            <td>{{ row.status|safe }}</td>
            <td>{{ '%.2f'|format(row.amount) }} {{ row.currency }}</td>
          </tr>
          {% endfor %}
          {% if not expense_rows %}
          <tr><td colspan="6" class="muted">V tomto roce nejsou evidované žádné výdaje.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </section>
    """
    return render_page(
        "Soupis výdajů",
        "Přehled všech výdajů za vybraný rok.",
        body,
        "year-report",
        **data,
    )


@app.get("/year-report/print")
def year_report_print_page() -> str:
    conn = open_db()
    data = collect_year_report_data(conn, request.args.get("year", "").strip())
    body = """
<!doctype html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Roční přehled {{ selected_year }}</title>
  <style>
    @page { size: A4; margin: 10mm; }
    body{margin:0;background:#e9eef8;font-family:"Segoe UI",Tahoma,sans-serif;color:#111827}
    .sheet{width:190mm;min-height:277mm;margin:0 auto 8mm auto;background:#fff;box-shadow:0 10px 30px rgba(15,23,42,.16);position:relative;overflow:hidden}
    .topbar{height:18mm;background:#1d2433}
    .hero{margin:-12mm 10mm 0 10mm;background:#eef0ff;border-radius:10px;padding:8mm 9mm 6mm 9mm;display:flex;justify-content:space-between;align-items:flex-start}
    .hero h1{margin:0;font-size:20pt;line-height:1.05}
    .hero .sub{margin-top:2mm;color:#5b6475;font-size:9pt}
    .year-pill{background:#5b21b6;color:#fff;border-radius:8px;padding:4mm 7mm;font-weight:700;font-size:12pt;min-width:34mm;text-align:center}
    .meta{display:grid;grid-template-columns:1fr 1fr;gap:8mm;padding:7mm 10mm 0 10mm}
    .meta h3{margin:0 0 2mm 0;font-size:8pt;letter-spacing:.12em;text-transform:uppercase;color:#5b21b6}
    .meta .box{min-height:18mm}
    .issuer-card{padding:0 10mm;margin-top:5mm}
    .issuer-panel{background:#fff;border-left:4px solid #5b21b6;padding:0 0 0 4mm}
    .issuer-panel h3{margin:0 0 2mm 0;font-size:8pt;letter-spacing:.12em;text-transform:uppercase;color:#5b21b6}
    .issuer-name{font-weight:800;font-size:11pt;margin-bottom:1mm}
    .issuer-line{font-size:9pt;line-height:1.45}
    .summary{margin:6mm 10mm 0 10mm;background:#4420b8;color:#fff;border-radius:0;padding:5mm 6mm;display:grid;grid-template-columns:1fr 1fr 1fr;gap:6mm}
    .summary .label{font-size:8pt;opacity:.82;text-transform:uppercase;letter-spacing:.08em}
    .summary .value{margin-top:2mm;font-size:16pt;font-weight:800}
    .content{padding:6mm 10mm 8mm 10mm}
    .section-title{margin:0 0 3mm 0;font-size:8pt;letter-spacing:.12em;text-transform:uppercase;color:#5b21b6}
    table{width:100%;border-collapse:collapse;font-size:8.7pt}
    th,td{padding:2.2mm 2mm;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top}
    th{background:#eef0ff;color:#24304d;font-size:8pt}
    .grid{display:grid;grid-template-columns:1.08fr .92fr;gap:8mm}
    .compact td,.compact th{padding-top:1.8mm;padding-bottom:1.8mm}
    .totals{margin-top:5mm;display:grid;grid-template-columns:1fr 70mm;gap:8mm;align-items:end}
    .note{font-size:8pt;color:#6b7280}
    .grand{background:#4420b8;color:#fff;padding:4mm 5mm;border-radius:0;display:flex;justify-content:space-between;align-items:center;font-weight:800}
    .grand .label{font-size:10pt}
    .grand .value{font-size:14pt}
    .muted{color:#6b7280}
    .toolbar{position:sticky;top:0;padding:8px;text-align:right;background:#e9eef8}
    .toolbar button{border:none;border-radius:10px;padding:10px 14px;background:#4420b8;color:#fff;font-weight:700;cursor:pointer}
    .page-break{page-break-before:always;break-before:page}
    .list-sheet .hero{padding-bottom:5mm}
    .list-sheet .content{padding-top:5mm}
    .list-sheet table{font-size:8.4pt}
    .list-sheet th,.list-sheet td{padding:2mm 1.8mm}
    .list-sheet .content-card-lite{margin:0 10mm 6mm 10mm;background:#eef0ff;border-radius:10px;padding:5mm 6mm}
    .list-sheet .content-card-lite .summary-line{display:flex;justify-content:space-between;gap:8mm;font-size:10pt;font-weight:700}
    @media print {.toolbar{display:none} body{background:#fff} .sheet{box-shadow:none;margin:0 auto} .page-break{page-break-before:always;break-before:page}}
  </style>
</head>
<body>
  <div class="toolbar"><button onclick="window.print()">Tisk / Uložit do PDF</button></div>
  <div class="sheet">
    <div class="topbar"></div>
    <div class="hero">
      <div>
        <h1>Roční přehled</h1>
        <div class="sub">Příjmy, výdaje a zisk na jedné stránce A4.</div>
      </div>
      <div class="year-pill">{{ selected_year }}</div>
    </div>
    <div class="issuer-card">
      <div class="issuer-panel">
        <h3>Dodavatel</h3>
        <div class="issuer-name">{{ issuer.company_name }}</div>
        {% for line in issuer.address.splitlines() %}
        <div class="issuer-line">{{ line }}</div>
        {% endfor %}
        <div class="issuer-line">IČ: {{ issuer.ico }}{% if issuer.dic and issuer.dic != '-' %} &nbsp;&nbsp; DIČ: {{ issuer.dic }}{% endif %}</div>
      </div>
    </div>
    <div class="meta">
      <div class="box">
        <h3>Příjmy</h3>
        <div><strong>{{ income_rows|length }}</strong> faktur</div>
        <div class="muted">Součet všech příjmů za vybraný rok.</div>
      </div>
      <div class="box">
        <h3>Výdaje</h3>
        <div><strong>{{ expense_rows|length }}</strong> výdajů</div>
        <div class="muted">Součet všech výdajů za vybraný rok.</div>
      </div>
    </div>
    <div class="summary">
      <div><div class="label">Příjmy</div><div class="value">{{ '%.2f'|format(income_total) }} Kč</div></div>
      <div><div class="label">Výdaje</div><div class="value">{{ '%.2f'|format(expense_total) }} Kč</div></div>
      <div><div class="label">Zisk</div><div class="value">{{ '%.2f'|format(profit_total) }} Kč</div></div>
    </div>
    <div class="content">
      <div class="grid">
        <div>
          <h3 class="section-title">Měsíční přehled</h3>
          <table class="compact">
            <thead><tr><th>Měsíc</th><th>Příjmy</th><th>Výdaje</th><th>Zisk</th></tr></thead>
            <tbody>
              {% for row in month_rows %}
              <tr>
                <td>{{ row.label }}</td>
                <td>{{ '%.0f'|format(row.income) }} Kč</td>
                <td>{{ '%.0f'|format(row.expense) }} Kč</td>
                <td><strong>{{ '%.0f'|format(row.profit) }} Kč</strong></td>
              </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
        <div>
          <h3 class="section-title">Souhrn</h3>
          <table class="compact">
            <tbody>
              <tr><th>Rok</th><td>{{ selected_year }}</td></tr>
              <tr><th>Počet příjmů</th><td>{{ income_rows|length }}</td></tr>
              <tr><th>Počet výdajů</th><td>{{ expense_rows|length }}</td></tr>
              <tr><th>Příjmy celkem</th><td>{{ '%.2f'|format(income_total) }} Kč</td></tr>
              <tr><th>Výdaje celkem</th><td>{{ '%.2f'|format(expense_total) }} Kč</td></tr>
              <tr><th>Zisk</th><td><strong>{{ '%.2f'|format(profit_total) }} Kč</strong></td></tr>
            </tbody>
          </table>
          <div style="margin-top:6mm">
            <h3 class="section-title">Vzorec</h3>
            <div class="note">Zisk = příjmy - výdaje</div>
          </div>
        </div>
      </div>
      <div class="totals">
        <div class="note">Přehled je připravený pro tisk na jednu A4. Využívá stejný čistý styl a barevný směr jako faktura.</div>
        <div class="grand"><span class="label">Zisk za rok {{ selected_year }}</span><span class="value">{{ '%.2f'|format(profit_total) }} Kč</span></div>
      </div>
    </div>
  </div>

  <div class="sheet list-sheet page-break">
    <div class="topbar"></div>
    <div class="hero">
      <div>
        <h1>Soupis příjmů</h1>
        <div class="sub">Detailní seznam všech příjmů za rok {{ selected_year }}.</div>
      </div>
      <div class="year-pill">{{ selected_year }}</div>
    </div>
    <div class="issuer-card">
      <div class="issuer-panel">
        <h3>Dodavatel</h3>
        <div class="issuer-name">{{ issuer.company_name }}</div>
        {% for line in issuer.address.splitlines() %}
        <div class="issuer-line">{{ line }}</div>
        {% endfor %}
        <div class="issuer-line">IČ: {{ issuer.ico }}{% if issuer.dic and issuer.dic != '-' %} &nbsp;&nbsp; DIČ: {{ issuer.dic }}{% endif %}</div>
      </div>
    </div>
    <div class="content-card-lite">
      <div class="summary-line"><span>Počet příjmů: {{ income_rows|length }}</span><span>Příjmy celkem: {{ '%.2f'|format(income_total) }} Kč</span></div>
    </div>
    <div class="content">
      <h3 class="section-title">Příjmy</h3>
      <table>
        <thead><tr><th>Faktura</th><th>Odběratel</th><th>Vystavení</th><th>Stav</th><th>Částka</th></tr></thead>
        <tbody>
          {% for row in income_rows %}
          <tr>
            <td><strong>{{ row.number }}</strong></td>
            <td>{{ row.customer }}</td>
            <td>{{ row.issue }}</td>
            <td>{{ row.status|striptags }}</td>
            <td>{{ '%.2f'|format(row.total) }} {{ row.currency }}</td>
          </tr>
          {% endfor %}
          {% if not income_rows %}
          <tr><td colspan="5" class="muted">V tomto roce nejsou evidované žádné příjmy.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>

  <div class="sheet list-sheet page-break">
    <div class="topbar"></div>
    <div class="hero">
      <div>
        <h1>Soupis výdajů</h1>
        <div class="sub">Detailní seznam všech výdajů za rok {{ selected_year }}.</div>
      </div>
      <div class="year-pill">{{ selected_year }}</div>
    </div>
    <div class="issuer-card">
      <div class="issuer-panel">
        <h3>Dodavatel</h3>
        <div class="issuer-name">{{ issuer.company_name }}</div>
        {% for line in issuer.address.splitlines() %}
        <div class="issuer-line">{{ line }}</div>
        {% endfor %}
        <div class="issuer-line">IČ: {{ issuer.ico }}{% if issuer.dic and issuer.dic != '-' %} &nbsp;&nbsp; DIČ: {{ issuer.dic }}{% endif %}</div>
      </div>
    </div>
    <div class="content-card-lite">
      <div class="summary-line"><span>Počet výdajů: {{ expense_rows|length }}</span><span>Výdaje celkem: {{ '%.2f'|format(expense_total) }} Kč</span></div>
    </div>
    <div class="content">
      <h3 class="section-title">Výdaje</h3>
      <table>
        <thead><tr><th>Doklad</th><th>Dodavatel</th><th>Datum</th><th>Kategorie</th><th>Stav</th><th>Částka</th></tr></thead>
        <tbody>
          {% for row in expense_rows %}
          <tr>
            <td><strong>{{ row.title }}</strong></td>
            <td>{{ row.supplier_name }}</td>
            <td>{{ row.expense_date }}</td>
            <td>{{ row.category }}</td>
            <td>{{ row.status|striptags }}</td>
            <td>{{ '%.2f'|format(row.amount) }} {{ row.currency }}</td>
          </tr>
          {% endfor %}
          {% if not expense_rows %}
          <tr><td colspan="6" class="muted">V tomto roce nejsou evidované žádné výdaje.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>
    """
    return render_template_string(body, **data)


@app.route("/expenses/new", methods=["GET", "POST"])
def new_expense_page() -> str:
    conn = open_db()
    try_ensure_recurring_expenses(conn)
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Název výdaje je povinný.", "error")
            return redirect(url_for("new_expense_page"))
        expense_date = parse_date(request.form.get("expense_date", ""), default=date.today())
        amount = float(str(request.form.get("amount", "0")).replace(",", "."))
        recurring = request.form.get("recurring") == "on"
        recurring_period = request.form.get("recurring_period", "")
        attachment_verified, price_confirmed, price_manual_override = review_flags_from_form(request.form)
        status = normalize_expense_status(
            request.form.get("status", "review"),
            attachment_verified=attachment_verified,
            price_confirmed=price_confirmed,
        )
        if recurring and not attachment_verified:
            status = "review"
        elif expense_is_review_complete({"attachment_path": request.form.get("attachment_path", ""), "price_confirmed": 1 if price_confirmed else 0, "attachment_verified": 1 if attachment_verified else 0}):
            status = "ok"
        recurring_series_id = f"exp-{datetime.now().strftime('%Y%m%d%H%M%S')}-{os.getpid()}"
        add_expense(
            conn,
            expense_date,
            title,
            amount,
            request.form.get("category", ""),
            request.form.get("note", ""),
            float(str(request.form.get("amount_without_vat", "0")).replace(",", ".")),
            float(str(request.form.get("vat_rate", "0")).replace(",", ".")),
            float(str(request.form.get("amount_with_vat", "0")).replace(",", ".")),
            request.form.get("currency", "CZK"),
            request.form.get("paid_date", ""),
            request.form.get("due_date", ""),
            request.form.get("supplier_name", ""),
            request.form.get("supplier_ico", ""),
            request.form.get("supplier_dic", ""),
            request.form.get("document_number", ""),
            request.form.get("variable_symbol", ""),
            request.form.get("document_type", ""),
            status,
            request.form.get("payment_method", ""),
            request.form.get("payment_account", ""),
            request.form.get("attachment_path", ""),
            request.form.get("external_link", ""),
            request.form.get("project_code", ""),
            request.form.get("cost_center", ""),
            request.form.get("expense_scope", ""),
            request.form.get("tax_deductible") == "on",
            recurring,
            recurring_period,
            recurring_series_id if recurring else "",
            None,
            False,
            price_confirmed,
            price_manual_override,
            attachment_verified,
        )
        if recurring:
            root = conn.execute("SELECT id FROM expenses WHERE recurring_series_id = ? ORDER BY id DESC LIMIT 1", (recurring_series_id,)).fetchone()
            if root is not None:
                sync_recurring_expense_series(conn, int(root["id"]), until_year=date.today().year)
        flash("Výdaj byl uložen.", "info")
        return redirect(url_for("expenses_page"))

    categories = expense_category_options(conn)
    document_types = ["Přijatá faktura", "Paragon", "Účtenka", "Zálohová faktura", "Interní výdaj"]
    payment_methods = ["Hotově", "Kartou", "Převodem"]
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Nový výdaj</h2><div class="section-copy">Zjednodušený zápis výdaje s možností doplnit PDF doklad a nastavit opakování.</div></div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název výdaje</label><input name="title" required></div><div class="field"><label>Dodavatel</label><input name="supplier_name"></div><div class="field"><label>Částka celkem</label><input name="amount" value="0" required></div><div class="field"><label>Kategorie</label><select name="category">{% for item in categories %}<option value="{{ item }}">{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Typ dokladu</label><select name="document_type">{% for item in document_types %}<option value="{{ item }}">{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Datum dokladu</label><input type="date" name="expense_date" value="{{ today }}" required></div><div class="field"><label>Datum úhrady</label><input type="date" name="paid_date"></div><div class="field"><label>Datum splatnosti</label><input type="date" name="due_date"></div><div class="field"><label>Stav</label><select name="status">{% for value, label in expense_statuses %}<option value="{{ value }}" {% if value == 'review' %}selected{% endif %}>{{ label }}</option>{% endfor %}</select></div><div class="field"><label>Číslo dokladu</label><input name="document_number"></div><div class="field"><label>Projekt / zakázka</label><input name="project_code"></div><div class="field"><label>Způsob úhrady</label><select name="payment_method">{% for item in payment_methods %}<option value="{{ item }}">{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Opakování</label><select name="recurring_period"><option value="">Neopakovat</option>{% for value, label in recurring_periods %}<option value="{{ value }}">{{ label }}</option>{% endfor %}</select></div><div class="field" style="align-self:end"><label><input type="checkbox" name="recurring"> Vytvořit jako opakovaný výdaj</label><div class="muted">Budoucí výskyty se založí do konce roku a budou ve stavu Dolož fakturu.</div></div><div class="field" style="grid-column:1 / -1;"><label>Poznámka</label><input name="note" placeholder="Doplňující informace k výdaji"></div></div><input type="hidden" name="amount_without_vat" value="0"><input type="hidden" name="vat_rate" value="0"><input type="hidden" name="amount_with_vat" value="0"><input type="hidden" name="currency" value="CZK"><input type="hidden" name="supplier_ico" value=""><input type="hidden" name="supplier_dic" value=""><input type="hidden" name="variable_symbol" value=""><input type="hidden" name="payment_account" value=""><input type="hidden" name="attachment_path" value=""><input type="hidden" name="external_link" value=""><input type="hidden" name="cost_center" value=""><input type="hidden" name="expense_scope" value="Podnikání"><div class="toolbar"><button class="button" type="submit">Uložit výdaj</button><a class="button-link" href="{{ url_for('expenses_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Nový výdaj", "Samostatná stránka pro založení výdaje.", body, "expenses", today=date.today().isoformat(), categories=categories, document_types=document_types, payment_methods=payment_methods, expense_statuses=expense_status_options(), recurring_periods=recurring_period_options())


@app.route("/expenses/<int:expense_id>/edit", methods=["GET", "POST"])
def edit_expense_page(expense_id: int) -> str:
    conn = open_db()
    try_ensure_recurring_expenses(conn)
    row = get_expense(conn, expense_id)
    if row is None:
        flash("Výdaj nebyl nalezen.", "error")
        return redirect(url_for("expenses_page"))
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        if not title:
            flash("Název výdaje je povinný.", "error")
            return redirect(url_for("edit_expense_page", expense_id=expense_id))
        attachment_verified, price_confirmed, price_manual_override = review_flags_from_form(
            request.form,
            fallback_attachment=str(row["attachment_path"] or ""),
        )
        current_status = normalize_expense_status(
            request.form.get("status", "review"),
            attachment_verified=attachment_verified,
            price_confirmed=price_confirmed,
        )
        recurring = request.form.get("recurring") == "on" if not int(row["recurring_generated"] or 0) else bool(row["recurring"])
        recurring_period = request.form.get("recurring_period", "") if not int(row["recurring_generated"] or 0) else str(row["recurring_period"] or "")
        if recurring and not attachment_verified:
            current_status = "review"
        elif attachment_verified and price_confirmed:
            current_status = "ok"
        old_amount = float(row["amount"] or 0)
        update_expense(
            conn,
            expense_id,
            parse_date(request.form.get("expense_date", ""), default=date.today()),
            title,
            float(str(request.form.get("amount", "0")).replace(",", ".")),
            request.form.get("category", ""),
            request.form.get("note", ""),
            float(str(request.form.get("amount_without_vat", "0")).replace(",", ".")),
            float(str(request.form.get("vat_rate", "0")).replace(",", ".")),
            float(str(request.form.get("amount_with_vat", "0")).replace(",", ".")),
            request.form.get("currency", "CZK"),
            request.form.get("paid_date", ""),
            request.form.get("due_date", ""),
            request.form.get("supplier_name", ""),
            request.form.get("supplier_ico", ""),
            request.form.get("supplier_dic", ""),
            request.form.get("document_number", ""),
            request.form.get("variable_symbol", ""),
            request.form.get("document_type", ""),
            current_status,
            request.form.get("payment_method", ""),
            request.form.get("payment_account", ""),
            request.form.get("attachment_path", ""),
            request.form.get("external_link", ""),
            request.form.get("project_code", ""),
            request.form.get("cost_center", ""),
            request.form.get("expense_scope", ""),
            request.form.get("tax_deductible") == "on",
            recurring,
            recurring_period,
            request.form.get("recurring_series_id", str(row["recurring_series_id"] or "")),
            int(request.form.get("recurring_source_id", row["recurring_source_id"] or 0) or 0) or None,
            coerce_checkbox(request.form.get("recurring_generated")) or bool(row["recurring_generated"]),
            price_confirmed,
            price_manual_override or abs(float(str(request.form.get("amount", "0")).replace(",", ".")) - old_amount) > 0.0001,
            attachment_verified,
        )
        refreshed = get_expense(conn, expense_id)
        if refreshed is not None and int(refreshed["recurring"] or 0):
            if int(refreshed["recurring_generated"] or 0):
                if abs(float(refreshed["amount"] or 0) - old_amount) > 0.0001:
                    propagate_recurring_amounts(conn, expense_id)
            else:
                rebuild_recurring_future_expenses(conn, expense_id)
                if abs(float(refreshed["amount"] or 0) - old_amount) > 0.0001:
                    propagate_recurring_amounts(conn, expense_id)
        flash("Výdaj byl aktualizován.", "info")
        return redirect(url_for("expenses_page"))

    categories = expense_category_options(conn)
    document_types = ["Přijatá faktura", "Paragon", "Účtenka", "Zálohová faktura", "Interní výdaj"]
    payment_methods = ["Hotově", "Kartou", "Převodem"]
    review_meta = expense_review_meta(row)
    current_expense_status = normalize_expense_status(str(row["status"] or ""), bool(review_meta["attachment_verified"]), bool(review_meta["price_confirmed"]))
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Upravit výdaj</h2><div class="section-copy">Úprava hlavních údajů, opakování a kontroly dokladu.</div></div><div>{{ status_pill(current_expense_status)|safe }}</div></div><form method="post" class="stack"><div class="form-grid"><div class="field"><label>Název výdaje</label><input name="title" value="{{ row.title }}" required></div><div class="field"><label>Dodavatel</label><input name="supplier_name" value="{{ row.supplier_name or '' }}"></div><div class="field"><label>Částka celkem</label><input name="amount" value="{{ row.amount or 0 }}" required></div><div class="field"><label>Kategorie</label><select name="category">{% for item in categories %}<option value="{{ item }}" {% if row.category == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Typ dokladu</label><select name="document_type">{% for item in document_types %}<option value="{{ item }}" {% if row.document_type == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Datum dokladu</label><input type="date" name="expense_date" value="{{ row.expense_date }}" required></div><div class="field"><label>Datum úhrady</label><input type="date" name="paid_date" value="{{ row.paid_date or '' }}"></div><div class="field"><label>Datum splatnosti</label><input type="date" name="due_date" value="{{ row.due_date or '' }}"></div><div class="field"><label>Stav</label><select name="status">{% for value, label in expense_statuses %}<option value="{{ value }}" {% if current_expense_status == value %}selected{% endif %}>{{ label }}</option>{% endfor %}</select></div><div class="field"><label>Číslo dokladu</label><input name="document_number" value="{{ row.document_number or '' }}"></div><div class="field"><label>Projekt / zakázka</label><input name="project_code" value="{{ row.project_code or '' }}"></div><div class="field"><label>Způsob úhrady</label><select name="payment_method">{% for item in payment_methods %}<option value="{{ item }}" {% if row.payment_method == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Perioda opakování</label>{% if row.recurring_generated %}<input value="{{ review_meta.period_label }}" disabled><input type="hidden" name="recurring_period" value="{{ row.recurring_period or '' }}">{% else %}<select name="recurring_period"><option value="">Neopakovat</option>{% for value, label in recurring_periods %}<option value="{{ value }}" {% if row.recurring_period == value %}selected{% endif %}>{{ label }}</option>{% endfor %}</select>{% endif %}</div><div class="field"><label>PDF doklad</label><div class="muted">{{ review_meta.attachment_label }}</div><div class="toolbar"><a class="button-link" href="{{ url_for('import_expense_pdf_for_existing_page', expense_id=row.id) }}">Načíst / nahradit PDF</a>{% if review_meta.attachment_path %}<a class="button-link" href="{{ url_for('expense_attachment_page', expense_id=row.id) }}" target="_blank">Otevřít PDF</a>{% endif %}</div></div><div class="field" style="grid-column:1 / -1;"><label>Poznámka</label><input name="note" value="{{ row.note or '' }}"></div></div><input type="hidden" name="amount_without_vat" value="{{ row.amount_without_vat or 0 }}"><input type="hidden" name="vat_rate" value="{{ row.vat_rate or 0 }}"><input type="hidden" name="amount_with_vat" value="{{ row.amount_with_vat or row.amount or 0 }}"><input type="hidden" name="currency" value="{{ row.currency or 'CZK' }}"><input type="hidden" name="supplier_ico" value="{{ row.supplier_ico or '' }}"><input type="hidden" name="supplier_dic" value="{{ row.supplier_dic or '' }}"><input type="hidden" name="variable_symbol" value="{{ row.variable_symbol or '' }}"><input type="hidden" name="payment_account" value="{{ row.payment_account or '' }}"><input type="hidden" name="attachment_path" value="{{ row.attachment_path or '' }}"><input type="hidden" name="external_link" value="{{ row.external_link or '' }}"><input type="hidden" name="cost_center" value="{{ row.cost_center or '' }}"><input type="hidden" name="expense_scope" value="{{ row.expense_scope or '' }}"><input type="hidden" name="recurring_series_id" value="{{ row.recurring_series_id or '' }}"><input type="hidden" name="recurring_source_id" value="{{ row.recurring_source_id or '' }}"><input type="hidden" name="recurring_generated" value="{{ 1 if row.recurring_generated else 0 }}"><input type="hidden" name="price_confirmed" value="{{ 1 if review_meta.price_confirmed else 0 }}"><div class="content-card" style="padding:18px;margin-top:8px"><div class="section-head"><div><h2>Kontrola opakovaného výdaje</h2><div class="section-copy">Každý opakovaný výdaj musí mít PDF a potvrzenou cenu. Změna ceny se propíše do budoucích výskytů stejné série.</div></div></div><div class="grid-2"><div><strong>Opakování:</strong> {{ review_meta.period_label if row.recurring else 'Ne' }}</div><div><strong>Doklad:</strong> {{ 'Doložený' if review_meta.attachment_verified else 'Chybí PDF' }}</div><div><strong>Cena:</strong> {{ 'Potvrzená' if review_meta.price_confirmed else 'Nepotvrzená' }}</div><div><strong>Kontrola:</strong> {{ 'OK' if review_meta.review_complete else 'Dolož fakturu' }}</div></div></div><div class="toolbar"><label><input type="checkbox" name="tax_deductible" {% if row.tax_deductible %}checked{% endif %}> Daňově uznatelné</label><label><input type="checkbox" name="recurring" {% if row.recurring %}checked{% endif %} {% if row.recurring_generated %}disabled{% endif %}> Pravidelný výdaj</label></div><div class="toolbar"><button class="button" type="submit">Uložit změny</button><a class="button-link" href="{{ url_for('expenses_page') }}">Zpět</a></div></form></section>
    """
    return render_page("Upravit výdaj", "Jednoduchá editace výdaje.", body, "expenses", row=row, categories=categories, document_types=document_types, payment_methods=payment_methods, expense_statuses=expense_status_options(), recurring_periods=recurring_period_options(), review_meta=review_meta, status_pill=status_pill, current_expense_status=current_expense_status)


@app.post("/expenses/<int:expense_id>/delete")
def delete_expense_page(expense_id: int):
    try:
        conn = open_db()
        delete_expense(conn, expense_id)
        flash("Výdaj byl smazán.", "info")
    except sqlite3.OperationalError as exc:
        if "locked" in str(exc).lower():
            flash("Výdaj teď nejde smazat, protože databázi používá jiný proces. Zavři stará okna aplikace a zkus to znovu.", "error")
        else:
            raise
    except Exception as exc:
        flash(f"Výdaj se nepodařilo smazat: {exc}", "error")
    return redirect(url_for("expenses_page"))


@app.route("/expenses/<int:expense_id>/pdf", methods=["GET", "POST"])
def import_expense_pdf_for_existing_page(expense_id: int):
    conn = open_db()
    row = get_expense(conn, expense_id)
    if row is None:
        flash("Výdaj nebyl nalezen.", "error")
        return redirect(url_for("expenses_page"))

    if request.method == "GET":
        body = """
        <section class="content-card"><div class="section-head"><div><h2>Nahrát PDF k výdaji</h2><div class="section-copy">Nahraj PDF k existujícímu výdaji. Po schválení se doplní doklad, cena a kontrolní stav.</div></div></div><form method="post" enctype="multipart/form-data" class="stack"><div class="field"><label>Výdaj</label><input value="{{ row.title }}" disabled></div><div class="field"><label>PDF doklad</label><input type="file" name="expense_pdf" accept=".pdf,application/pdf" required></div><div class="toolbar"><button class="button" type="submit">Načíst PDF</button><a class="button-link" href="{{ url_for('edit_expense_page', expense_id=row.id) }}">Zpět</a></div></form></section>
        """
        return render_page("PDF k výdaji", "Doložení dokladu k existujícímu výdaji.", body, "expenses", row=row)

    upload = request.files.get("expense_pdf")
    if upload is None or not upload.filename:
        flash("Vyber PDF doklad pro načtení.", "error")
        return redirect(url_for("import_expense_pdf_for_existing_page", expense_id=expense_id))
    if Path(upload.filename).suffix.lower() != ".pdf":
        flash("Podporován je pouze PDF soubor.", "error")
        return redirect(url_for("import_expense_pdf_for_existing_page", expense_id=expense_id))

    try:
        stored_path = _store_uploaded_pdf(upload)
        extracted_text = _extract_pdf_text(stored_path)
        guessed = _guess_expense_from_pdf_text(extracted_text)
    except Exception as exc:
        flash(f"Načtení PDF se nepodařilo: {exc}", "error")
        return redirect(url_for("edit_expense_page", expense_id=expense_id))

    categories = expense_category_options(conn)
    document_types = ["Přijatá faktura", "Paragon", "Účtenka", "Interní výdaj"]
    payment_methods = ["Hotově", "Kartou", "Převodem"]
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Schválení PDF pro existující výdaj</h2><div class="section-copy">Zkontroluj údaje z PDF. Po potvrzení se výdaj aktualizuje a označí jako OK.</div></div></div><form method="post" action="{{ url_for('save_expense_pdf_page') }}" class="stack"><div class="form-grid"><div class="field"><label>Název výdaje</label><input name="title" value="{{ guessed.title or row.title }}" required></div><div class="field"><label>Dodavatel</label><input name="supplier_name" value="{{ guessed.supplier_name or row.supplier_name }}"></div><div class="field"><label>Částka celkem</label><input name="amount" value="{{ guessed.amount or row.amount }}" required></div><div class="field"><label>Kategorie</label><select name="category">{% for item in categories %}<option value="{{ item }}" {% if (guessed.category or row.category) == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Typ dokladu</label><select name="document_type">{% for item in document_types %}<option value="{{ item }}" {% if (guessed.document_type or row.document_type) == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Datum dokladu</label><input type="date" name="expense_date" value="{{ guessed.expense_date or row.expense_date }}" required></div><div class="field"><label>Datum splatnosti</label><input type="date" name="due_date" value="{{ guessed.due_date or row.due_date }}"></div><div class="field"><label>Stav</label><select name="status">{% for value, label in expense_statuses %}<option value="{{ value }}" {% if value == 'ok' %}selected{% endif %}>{{ label }}</option>{% endfor %}</select></div><div class="field"><label>Číslo dokladu</label><input name="document_number" value="{{ guessed.document_number or row.document_number }}"></div><div class="field"><label>Projekt / zakázka</label><input name="project_code" value="{{ row.project_code }}"></div><div class="field"><label>Způsob úhrady</label><select name="payment_method">{% for item in payment_methods %}<option value="{{ item }}" {% if (guessed.payment_method or row.payment_method) == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field" style="grid-column:1 / -1;"><label>Poznámka</label><input name="note" value="{{ row.note }}"></div></div><input type="hidden" name="expense_id" value="{{ row.id }}"><input type="hidden" name="attachment_path" value="{{ attachment_path }}"><input type="hidden" name="amount_without_vat" value="{{ guessed.amount_without_vat or row.amount_without_vat }}"><input type="hidden" name="vat_rate" value="{{ guessed.vat_rate or row.vat_rate }}"><input type="hidden" name="amount_with_vat" value="{{ guessed.amount_with_vat or row.amount_with_vat or row.amount }}"><input type="hidden" name="currency" value="{{ guessed.currency or row.currency }}"><input type="hidden" name="variable_symbol" value="{{ guessed.variable_symbol or row.variable_symbol or '' }}"><div class="toolbar"><button class="button" type="submit">Schválit a aktualizovat výdaj</button><a class="button-link" href="{{ url_for('edit_expense_page', expense_id=row.id) }}">Zrušit</a></div></form></section>
    <section class="content-card"><div class="section-head"><div><h2>Náhled extrahovaného textu</h2><div class="section-copy">Pomáhá při kontrole správnosti informací z PDF.</div></div></div><pre style="white-space:pre-wrap;color:var(--muted);font:12px/1.5 Consolas,monospace;max-height:320px;overflow:auto">{{ extracted_text }}</pre></section>
    """
    return render_page("Schválení PDF výdaje", "Potvrzení údajů načtených z PDF.", body, "expenses", guessed=guessed, attachment_path=str(stored_path), extracted_text=extracted_text, categories=categories, document_types=document_types, payment_methods=payment_methods, expense_statuses=expense_status_options(), row=row)


@app.get("/expenses/<int:expense_id>/attachment")
def expense_attachment_page(expense_id: int):
    conn = open_db()
    row = get_expense(conn, expense_id)
    if row is None or not str(row["attachment_path"] or "").strip():
        flash("PDF příloha nebyla nalezena.", "error")
        return redirect(url_for("edit_expense_page", expense_id=expense_id))
    path = Path(str(row["attachment_path"]))
    if not path.exists():
        flash("Soubor PDF už v úložišti není.", "error")
        return redirect(url_for("edit_expense_page", expense_id=expense_id))
    return send_file(path, mimetype="application/pdf")


@app.route("/expenses/import-pdf", methods=["GET", "POST"])
def import_expense_pdf_page():
    conn = open_db()
    if request.method == "GET":
        body = """
        <section class="content-card"><div class="section-head"><div><h2>Import výdaje z PDF</h2><div class="section-copy">Nahraj PDF doklad, aplikace vytáhne data a před uložením je dá ke schválení.</div></div></div><form method="post" enctype="multipart/form-data" class="stack"><div class="field"><label>PDF doklad</label><input type="file" name="expense_pdf" accept=".pdf,application/pdf" required></div><div class="muted">JPG a PNG se nenačítají. Import je určený jen pro PDF.</div><div class="toolbar"><button class="button" type="submit">Načíst z PDF</button><a class="button-link" href="{{ url_for('expenses_page') }}">Zpět na výdaje</a></div></form></section>
        """
        return render_page("Import výdaje z PDF", "Samostatná stránka pro načtení výdaje z PDF dokladu.", body, "expenses")

    upload = request.files.get("expense_pdf")
    if upload is None or not upload.filename:
        flash("Vyber PDF doklad pro načtení.", "error")
        return redirect(url_for("import_expense_pdf_page"))
    if Path(upload.filename).suffix.lower() != ".pdf":
        flash("Podporován je pouze PDF soubor. JPG a PNG se nenačítají.", "error")
        return redirect(url_for("import_expense_pdf_page"))

    try:
        stored_path = _store_uploaded_pdf(upload)
        extracted_text = _extract_pdf_text(stored_path)
        guessed = _guess_expense_from_pdf_text(extracted_text)
    except Exception as exc:
        flash(f"Načtení PDF se nepodařilo: {exc}", "error")
        return redirect(url_for("import_expense_pdf_page"))

    categories = expense_category_options(conn)
    document_types = ["Přijatá faktura", "Paragon", "Účtenka", "Interní výdaj"]
    payment_methods = ["Hotově", "Kartou", "Převodem"]
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Schválení údajů z PDF</h2><div class="section-copy">Aplikace předvyplnila údaje z PDF. Zkontroluj je a potvrď uložení.</div></div></div><form method="post" action="{{ url_for('save_expense_pdf_page') }}" class="stack"><div class="form-grid"><div class="field"><label>Název výdaje</label><input name="title" value="{{ guessed.title }}" required></div><div class="field"><label>Dodavatel</label><input name="supplier_name" value="{{ guessed.supplier_name }}"></div><div class="field"><label>Částka celkem</label><input name="amount" value="{{ guessed.amount }}" required></div><div class="field"><label>Kategorie</label><select name="category">{% for item in categories %}<option value="{{ item }}" {% if guessed.category == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Typ dokladu</label><select name="document_type">{% for item in document_types %}<option value="{{ item }}" {% if guessed.document_type == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field"><label>Datum dokladu</label><input type="date" name="expense_date" value="{{ guessed.expense_date }}" required></div><div class="field"><label>Datum splatnosti</label><input type="date" name="due_date" value="{{ guessed.due_date }}"></div><div class="field"><label>Stav</label><select name="status">{% for value, label in expense_statuses %}<option value="{{ value }}" {% if guessed.status == value or value == 'ok' %}selected{% endif %}>{{ label }}</option>{% endfor %}</select></div><div class="field"><label>Číslo dokladu</label><input name="document_number" value="{{ guessed.document_number }}"></div><div class="field"><label>Projekt / zakázka</label><input name="project_code" value="{{ guessed.project_code }}"></div><div class="field"><label>Způsob úhrady</label><select name="payment_method">{% for item in payment_methods %}<option value="{{ item }}" {% if guessed.payment_method == item %}selected{% endif %}>{{ item }}</option>{% endfor %}</select></div><div class="field" style="grid-column:1 / -1;"><label>Poznámka</label><input name="note" value="{{ guessed.note }}"></div></div><input type="hidden" name="attachment_path" value="{{ attachment_path }}"><input type="hidden" name="amount_without_vat" value="{{ guessed.amount_without_vat }}"><input type="hidden" name="vat_rate" value="{{ guessed.vat_rate }}"><input type="hidden" name="amount_with_vat" value="{{ guessed.amount_with_vat }}"><input type="hidden" name="currency" value="{{ guessed.currency }}"><input type="hidden" name="variable_symbol" value="{{ guessed.variable_symbol or '' }}"><div class="toolbar"><button class="button" type="submit">Schválit a uložit výdaj</button><a class="button-link" href="{{ url_for('expenses_page') }}">Zrušit</a></div></form></section>
    <section class="content-card"><div class="section-head"><div><h2>Náhled extrahovaného textu</h2><div class="section-copy">Pomáhá při kontrole správnosti informací z PDF.</div></div></div><pre style="white-space:pre-wrap;color:var(--muted);font:12px/1.5 Consolas,monospace;max-height:320px;overflow:auto">{{ extracted_text }}</pre></section>
    """
    return render_page("Schválení PDF výdaje", "Potvrzení údajů načtených z PDF.", body, "expenses", guessed=guessed, attachment_path=str(stored_path), extracted_text=extracted_text, categories=categories, document_types=document_types, payment_methods=payment_methods, expense_statuses=expense_status_options())


@app.post("/expenses/save-from-pdf")
def save_expense_pdf_page():
    conn = open_db()
    try:
        title = request.form.get("title", "").strip()
        if not title:
            flash("Název výdaje je povinný.", "error")
            return redirect(url_for("import_expense_pdf_page"))
        expense_id_raw = request.form.get("expense_id", "").strip()
        expense_date = parse_date(request.form.get("expense_date", ""), default=date.today())
        attachment_path = request.form.get("attachment_path", "").strip()
        if attachment_path:
            attachment_path = str(_relocate_expense_pdf(attachment_path, expense_date))
        amount = parse_amount_text(request.form.get("amount", "0"))
        amount_without_vat = parse_amount_text(request.form.get("amount_without_vat", "0"))
        amount_with_vat = parse_amount_text(request.form.get("amount_with_vat", "0"))
        vat_rate = parse_float_field(request.form.get("vat_rate", "0"), 0.0)
        currency = request.form.get("currency", "CZK").strip() or "CZK"
        category = request.form.get("category", "").strip()
        note = request.form.get("note", "").strip()
        due_date = request.form.get("due_date", "").strip()
        supplier_name = request.form.get("supplier_name", "").strip()
        document_number = request.form.get("document_number", "").strip()
        variable_symbol = request.form.get("variable_symbol", "").strip()
        document_type = request.form.get("document_type", "").strip()
        payment_method = request.form.get("payment_method", "").strip()
        project_code = request.form.get("project_code", "").strip()
        normalized_status = normalize_expense_status(request.form.get("status", "ok"), attachment_verified=True, price_confirmed=True)

        if expense_id_raw:
            expense_id = parse_int_field(expense_id_raw, 0)
            row = get_expense(conn, expense_id)
            if row is None or expense_id <= 0:
                flash("Výdaj nebyl nalezen.", "error")
                return redirect(url_for("expenses_page"))
            update_expense(
                conn,
                expense_id,
                expense_date,
                title,
                amount,
                category,
                note,
                amount_without_vat,
                vat_rate,
                amount_with_vat,
                currency,
                request.form.get("paid_date", row["paid_date"] or ""),
                due_date,
                supplier_name,
                request.form.get("supplier_ico", row["supplier_ico"] or ""),
                request.form.get("supplier_dic", row["supplier_dic"] or ""),
                document_number,
                variable_symbol or str(row["variable_symbol"] or ""),
                document_type,
                normalized_status,
                payment_method,
                request.form.get("payment_account", row["payment_account"] or ""),
                attachment_path,
                request.form.get("external_link", row["external_link"] or ""),
                project_code,
                request.form.get("cost_center", row["cost_center"] or ""),
                request.form.get("expense_scope", row["expense_scope"] or "Podnikání"),
                bool(row["tax_deductible"]),
                bool(row["recurring"]),
                str(row["recurring_period"] or ""),
                str(row["recurring_series_id"] or ""),
                int(row["recurring_source_id"] or 0) or None,
                bool(row["recurring_generated"]),
                True,
                True,
                True,
            )
            if int(row["recurring"] or 0):
                if int(row["recurring_generated"] or 0):
                    propagate_recurring_amounts(conn, expense_id)
                else:
                    rebuild_recurring_future_expenses(conn, expense_id)
                    propagate_recurring_amounts(conn, expense_id)
            flash("PDF bylo přiřazeno k výdaji a údaje byly potvrzeny.", "info")
            return redirect(url_for("edit_expense_page", expense_id=expense_id))

        add_expense(
            conn,
            expense_date,
            title,
            amount,
            category,
            note,
            amount_without_vat,
            vat_rate,
            amount_with_vat,
            currency,
            "",
            due_date,
            supplier_name,
            "",
            "",
            document_number,
            variable_symbol,
            document_type,
            normalized_status,
            payment_method,
            "",
            attachment_path,
            "",
            project_code,
            "",
            "Podnikání",
            True,
            False,
            "",
            "",
            None,
            False,
            True,
            True,
            True,
        )
        flash("Výdaj z PDF byl uložen po schválení.", "info")
        return redirect(url_for("expenses_page"))
    except Exception as exc:
        log_app_error("save_expense_pdf_page", exc)
        flash(f"Uložení PDF výdaje se nepovedlo: {exc}", "error")
        return redirect(url_for("import_expense_pdf_page"))


@app.post("/invoices/import-excel")
def import_invoices_excel_page():
    conn = open_db()
    upload = request.files.get("excel_file")
    if upload is None or not upload.filename:
        flash("Vyber Excel soubor pro import.", "error")
        return redirect(url_for("invoices_page"))

    suffix = Path(upload.filename).suffix.lower() or ".xlsx"
    issuer_id_raw = request.form.get("issuer_id", "").strip()
    issuer_id = int(issuer_id_raw) if issuer_id_raw else None

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            upload.save(temp_file.name)
            temp_path = Path(temp_file.name)
        imported_invoices, imported_items, added_customers = import_invoices_from_excel(conn, temp_path, issuer_id)
    except Exception as exc:
        flash(f"Import z Excelu se nepodařil: {exc}", "error")
        return redirect(url_for("invoices_page"))
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink(missing_ok=True)

    flash(f"Import hotov: faktury {imported_invoices}, položky {imported_items}, noví odběratelé {added_customers}.", "info")
    return redirect(url_for("data_tools_page"))


@app.route("/data-tools")
def data_tools_page() -> str:
    conn = open_db()
    body = """
    <section class="grid-2">
      <div class="content-card"><div class="section-head"><div><h2>Import faktur z Excelu</h2><div class="section-copy">Jeden soubor = jeden rok, každý list = jedna faktura.</div></div></div><form method="post" action="{{ url_for('import_invoices_excel_page') }}" enctype="multipart/form-data" class="stack"><div class="field"><label>Soubor Excel</label><input type="file" name="excel_file" accept=".xlsx,.xlsm" required></div><div class="field"><label>Fakturující firma pro import</label><select name="issuer_id">{% for row in issuers %}<option value="{{ row.id }}">{{ row.company_name }}</option>{% endfor %}</select></div><div><button class="button" type="submit">Importovat faktury</button></div></form></div>
      <div class="content-card"><div class="section-head"><div><h2>Záloha a obnova</h2><div class="section-copy">Zálohuj databázi i PDF doklady do jednoho ZIP balíčku nebo obnov celý stav ze zálohy.</div></div></div><div class="stack"><a class="button" href="{{ url_for('backup_data_page') }}" onclick="return startBackupDownload(this)"><span class="backup-button-idle">Stáhnout kompletní ZIP zálohu</span><span class="backup-button-busy" style="display:none;">Vytvářím ZIP zálohu...</span></a><div id="backupStatusMessage" class="muted" style="display:none;">Aplikace právě připravuje zálohu. Podle velikosti souborů to může chvíli trvat.</div><form method="post" action="{{ url_for('restore_data_page') }}" enctype="multipart/form-data" class="stack"><div class="field"><label>ZIP nebo JSON záloha</label><input type="file" name="backup_file" accept=".zip,.json" required></div><div><button class="button-secondary" type="submit">Obnovit ze zálohy</button></div></form></div></div>
    </section>
    <section class="content-card"><div class="section-head"><div><h2>Správa databáze</h2><div class="section-copy">Nebezpečné operace nad celou databází, doklady i složkou exports.</div></div></div><div class="toolbar"><form method="post" action="{{ url_for('wipe_data_page') }}" onsubmit="return confirm('Opravdu smazat všechny faktury, odběratele, firmy, uložené doklady i celou složku exports včetně záloh?');"><button class="button-danger" type="submit">Smazat vše z databáze i soubory</button></form></div></section>
    """
    return render_page("Import / Export", "Hromadný import, zálohy a správa databáze.", body, "data-tools", issuers=list_issuers(conn))


@app.get("/data-tools/backup")
def backup_data_page():
    conn = open_db()
    try:
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = EXPORT_DIR / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        _create_backup_bundle(conn, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_path.name)
    except Exception as exc:
        flash(f"Zálohu se nepodařilo vytvořit: {exc}", "error")
        return redirect(url_for("settings_page"))


@app.post("/data-tools/restore")
def restore_data_page():
    conn = open_db()
    upload = request.files.get("backup_file")
    if upload is None or not upload.filename:
        flash("Vyber JSON zálohu pro obnovu.", "error")
        return redirect(url_for("data_tools_page"))

    temp_path: Path | None = None
    try:
        suffix = Path(upload.filename).suffix.lower() or ".zip"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            upload.save(temp_file.name)
            temp_path = Path(temp_file.name)
        customers, issuers, invoices, items = _restore_backup_bundle(conn, temp_path)
    except Exception as exc:
        flash(f"Obnova ze zálohy se nepodařila: {exc}", "error")
        return redirect(url_for("data_tools_page"))
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink(missing_ok=True)

    flash(f"Obnova hotova: odběratelé {customers}, firmy {issuers}, faktury {invoices}, položky {items}.", "info")
    return redirect(url_for("data_tools_page"))


@app.post("/data-tools/wipe")
def wipe_data_page():
    conn = open_db()
    try:
        clear_all_business_data(conn)
        deleted = _clear_business_files()
        flash(
            "Všechna data z databáze byla smazána. "
            f"Soubory: příjmy {deleted['income_docs']}, výdaje {deleted['expense_docs']}, "
            f"dočasné uploady {deleted['temp_uploads']}, exporty {deleted['exports']}. "
            "Výchozí firma byla znovu vytvořena.",
            "info",
        )
    except Exception as exc:
        flash(f"Úplné smazání se nepodařilo dokončit: {exc}", "error")
    return redirect(url_for("settings_page"))


@app.get("/guides")
def guides_page() -> str:
    workspace_dir = str(Path.cwd())
    db_path = str(Path(DB_FILE).resolve())
    export_path = str(INCOME_DOCS_DIR.resolve())
    upload_path = str(EXPENSE_DOCS_DIR.resolve())
    body = """
    <style>
      .guide-stack{display:grid;gap:16px}
      .guide-box{padding:0;overflow:hidden}
      .guide-box summary{list-style:none;cursor:pointer;padding:18px 20px;display:flex;justify-content:space-between;align-items:center;gap:16px}
      .guide-box summary::-webkit-details-marker{display:none}
      .guide-box .guide-content{padding:0 20px 20px 20px;display:grid;gap:18px}
      .guide-box .guide-title{margin:0;font-size:20px;letter-spacing:-.02em}
      .guide-list{display:grid;gap:10px}
      .guide-list div{line-height:1.6}
      .guide-table{width:100%;border-collapse:collapse}
      .guide-table th,.guide-table td{padding:12px 10px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
      .guide-table th{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}
      .guide-table tr:last-child td{border-bottom:none}
    </style>
    <section class="guide-stack">
      <section class="content-card">
        <div class="section-head"><div><h2>Obsah</h2><div class="section-copy">Rychlé odkazy na jednotlivé části návodu.</div></div></div>
        <div class="toolbar">
          <a class="button-link" href="#guides-overview">Co aplikace umí</a>
          <a class="button-link" href="#guides-runtime">Jak aplikace běží</a>
          <a class="button-link" href="#guides-invoices">Faktury a příjmy</a>
          <a class="button-link" href="#guides-expenses">Výdaje</a>
          <a class="button-link" href="#guides-contacts">Odběratelé a firmy</a>
          <a class="button-link" href="#guides-year">Roční přehled a tisk</a>
          <a class="button-link" href="#guides-backup">Zálohy a databáze</a>
          <a class="button-link" href="#guides-data">Kde jsou data</a>
          <a class="button-link" href="#guides-notes">Důležité poznámky</a>
        </div>
      </section>

      <details class="content-card guide-box" id="guides-overview">
        <summary>
          <div>
            <h2 class="guide-title">Co Aplikace Umí</h2>
            <div class="section-copy">Rychlý přehled hlavních sekcí a k čemu slouží.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Dashboard:</strong> hlavní přehled tržeb, výdajů, nezaplacených faktur, grafů a upozornění. Je zde i sbalený box <strong>Roční přehled</strong>.</div>
          <div><strong>Příjmy / Faktury:</strong> seznam vystavených faktur podle roků, import z Excelu a PDF, detail faktury, export PDF a HTML.</div>
          <div><strong>Výdaje:</strong> evidence nákladů, import z PDF, filtry, kategorie a samostatná záložka <strong>Dodavatelé</strong>.</div>
          <div><strong>Odběratelé:</strong> kompaktní seznam klientů, počet faktur, součet částek a rychlá akce <strong>Fakturovat</strong>.</div>
          <div><strong>Roční přehled:</strong> samostatná stránka se součtem příjmů, výdajů, zisku a tiskovými výstupy.</div>
          <div><strong>Nastavení:</strong> Google Drive, správa firem, import faktur z Excelu, záloha, obnova a smazání databáze.</div>
          <div><strong>Návody:</strong> technický i uživatelský přehled aplikace na jednom místě.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-runtime">
        <summary>
          <div>
            <h2 class="guide-title">Jak Aplikace Běží</h2>
            <div class="section-copy">Technický provoz lokálně na tomto počítači.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Typ aplikace:</strong> Flask webová aplikace v Pythonu.</div>
          <div><strong>Spuštění:</strong> hlavní soubor je <code>invoice_manager_gui.py</code> nebo spouštěcí <code>spustit_fakturaci.bat</code>.</div>
          <div><strong>Adresa v prohlížeči:</strong> <code>http://127.0.0.1:8000</code>.</div>
          <div><strong>Běh serveru:</strong> lokálně na tomto počítači, nejde o veřejný web.</div>
          <div><strong>Pracovní složka projektu:</strong> <code>{{ workspace_dir }}</code>.</div>
          <div><strong>Databáze:</strong> SQLite soubor <code>{{ db_path }}</code>.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-invoices">
        <summary>
          <div>
            <h2 class="guide-title">Faktury A Příjmy</h2>
            <div class="section-copy">Jak funguje tvorba, import a správa příjmových dokladů.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Nová faktura:</strong> může obsahovat více služeb a poznámku / projekt / dotaci i na více řádků.</div>
          <div><strong>DPH:</strong> cena se zadává jako včetně DPH, ale pokud firma není plátce DPH, sloupec DPH se ve formuláři skryje.</div>
          <div><strong>Stavy faktur:</strong> rozpracovaná, odeslaná, zaplacená a po splatnosti.</div>
          <div><strong>Seznam faktur:</strong> faktury jsou seskupené po rocích a řazené podle data vystavení.</div>
          <div><strong>Import z Excelu:</strong> spouští se z <strong>Nastavení</strong>. Jeden soubor = jeden rok, každý list = jedna faktura.</div>
          <div><strong>Import z PDF:</strong> dostupný přímo v sekci <strong>Příjmy / Faktury</strong>. Nejdřív zobrazí schvalovací formulář, až pak fakturu uloží.</div>
          <div><strong>Export:</strong> detail faktury umí export PDF i HTML. Výstupy se ukládají do <code>{{ export_path }}</code> podle roku.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-expenses">
        <summary>
          <div>
            <h2 class="guide-title">Výdaje</h2>
            <div class="section-copy">Evidence nákladů, doklady a dodavatelé.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Nový výdaj:</strong> je na samostatné stránce.</div>
          <div><strong>Import z PDF:</strong> je na samostatné stránce. Doklad se načte, předvyplní údaje a čeká na schválení.</div>
          <div><strong>Stavy výdajů:</strong> používají se jen <strong>Dolož fakturu</strong> a <strong>OK</strong>.</div>
          <div><strong>Dodavatelé:</strong> samostatná stránka v sekci Výdaje, podobná odběratelům. Zobrazuje dodavatele, dostupnost IČO/DIČ/PDF a součet dokladů.</div>
          <div><strong>Kategorie:</strong> nejsou v bočním menu, ale přímo v horní části sekce Výdaje.</div>
          <div><strong>Opakované výdaje:</strong> generují se jen do konce příslušného kalendářního roku, nepřetékají do dalšího roku.</div>
          <div><strong>PDF doklady výdajů:</strong> po schválení se ukládají do <code>{{ upload_path }}</code> podle roku.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-contacts">
        <summary>
          <div>
            <h2 class="guide-title">Odběratelé A Firmy</h2>
            <div class="section-copy">Správa klientů a vlastních fakturačních firem.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Odběratelé:</strong> mají vlastní samostatnou stránku a formulář pro nový záznam je oddělený.</div>
          <div><strong>Seznam odběratelů:</strong> je kompaktní, řazený abecedně a ukazuje počet faktur i součet částek.</div>
          <div><strong>Fakturovat odběratele:</strong> v akcích u odběratele je tlačítko <strong>Fakturovat</strong>, které otevře novou fakturu s předvybraným odběratelem.</div>
          <div><strong>Firmy:</strong> už nejsou v levém menu. Otevřeš je z <strong>Nastavení</strong>.</div>
          <div><strong>ARES:</strong> u firem i odběratelů lze načíst název, adresu, IČO a DIČ podle zadaného IČO.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-year">
        <summary>
          <div>
            <h2 class="guide-title">Roční Přehled A Tisk</h2>
            <div class="section-copy">Souhrn roku a tiskové výstupy.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Roční přehled:</strong> počítá příjmy, výdaje a zisk pro vybraný rok.</div>
          <div><strong>Bloky v přehledu:</strong> příjmy a výdaje jsou pod sebou a výchozí stav je sbalený.</div>
          <div><strong>Tisk A4:</strong> vytvoří více stránkový dokument se souhrnem, soupisem příjmů a soupisem výdajů.</div>
          <div><strong>Výpis příjmy:</strong> vytiskne všechny faktury vybraného roku do jednoho PDF.</div>
          <div><strong>Výpis výdaje:</strong> vytiskne všechny dostupné PDF doklady výdajů vybraného roku do jednoho PDF.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-backup">
        <summary>
          <div>
            <h2 class="guide-title">Zálohy, Obnova A Databáze</h2>
            <div class="section-copy">Hromadné datové operace jsou nově v Nastavení.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Import faktur z Excelu:</strong> je v Nastavení.</div>
          <div><strong>Záloha všeho:</strong> exportuje firmy, odběratele, faktury, položky, výdaje i nastavení do JSON.</div>
          <div><strong>Obnova ze zálohy:</strong> načte dříve uložený JSON a vrátí stav databáze.</div>
          <div><strong>Smazat vše z databáze i soubory:</strong> odstraní obchodní data, PDF doklady, exportované faktury, dočasné nahrané soubory i celou složku <code>exports</code> včetně záloh. Výchozí firma se znovu vytvoří.</div>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-data">
        <summary>
          <div>
            <h2 class="guide-title">Kde Jsou Data</h2>
            <div class="section-copy">Důležité cesty a co v nich najdeš.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content">
          <table class="guide-table">
            <thead><tr><th>Položka</th><th>Umístění</th><th>Poznámka</th></tr></thead>
            <tbody>
              <tr><td>Hlavní databáze</td><td><code>{{ db_path }}</code></td><td>SQLite databáze celé aplikace.</td></tr>
              <tr><td>Exportované faktury</td><td><code>{{ export_path }}</code></td><td>PDF a HTML faktury v podsložkách podle roku.</td></tr>
              <tr><td>PDF doklady výdajů</td><td><code>{{ upload_path }}</code></td><td>Schválené PDF soubory výdajů v podsložkách podle roku.</td></tr>
              <tr><td>Kód aplikace</td><td><code>{{ workspace_dir }}</code></td><td>Python soubory aplikace a backendu.</td></tr>
            </tbody>
          </table>
        </div>
      </details>

      <details class="content-card guide-box" id="guides-notes">
        <summary>
          <div>
            <h2 class="guide-title">Důležité Poznámky</h2>
            <div class="section-copy">Praktická doporučení pro běžný provoz.</div>
          </div>
          <span class="button-link">Otevřít</span>
        </summary>
        <div class="guide-content guide-list">
          <div><strong>Aplikace běží lokálně:</strong> když zavřeš server, web přestane být dostupný, ale data v databázi zůstanou.</div>
          <div><strong>Staré exporty se nepřepíšou v otevřeném PDF náhledu:</strong> pro kontrolu vždy otevři nově vytvořený soubor.</div>
          <div><strong>Zálohuj pravidelně:</strong> hlavně před velkým importem nebo mazáním dat.</div>
          <div><strong>PDF importy vždy kontroluj:</strong> import pomáhá, ale není účetní autorita.</div>
          <div><strong>Jedna spuštěná instance:</strong> je lepší mít otevřenou jen jednu instanci aplikace kvůli SQLite databázi.</div>
        </div>
      </details>
    </section>
    """
    return render_page("Návody", "Kompletní přehled funkcí, provozu, dat a omezení aplikace.", body, "guides", workspace_dir=workspace_dir, db_path=db_path, export_path=export_path, upload_path=upload_path)


@app.route("/invoices/import-pdf", methods=["GET", "POST"])
def import_invoice_pdf_page():
    conn = open_db()
    issuers = list_issuers(conn)
    if request.method == "GET":
        body = """
        <section class="content-card"><div class="section-head"><div><h2>Import příjmu z PDF</h2><div class="section-copy">Nahraj PDF faktury, aplikace vytáhne údaje a před uložením je dá ke schválení.</div></div></div><form method="post" enctype="multipart/form-data" class="stack"><div class="field"><label>PDF faktura</label><input type="file" name="invoice_pdf" accept=".pdf,application/pdf" required></div><div class="muted">Import slouží pro příjmové faktury. Po načtení můžeš údaje ještě upravit.</div><div class="toolbar"><button class="button" type="submit">Načíst z PDF</button><a class="button-link" href="{{ url_for('invoices_page') }}">Zpět na příjmy</a></div></form></section>
        """
        return render_page("Import příjmu z PDF", "Samostatná stránka pro načtení faktury z PDF.", body, "invoices")

    upload = request.files.get("invoice_pdf")
    if upload is None or not upload.filename:
        flash("Vyber PDF fakturu pro načtení.", "error")
        return redirect(url_for("import_invoice_pdf_page"))
    if Path(upload.filename).suffix.lower() != ".pdf":
        flash("Podporován je pouze PDF soubor.", "error")
        return redirect(url_for("import_invoice_pdf_page"))

    try:
        stored_path = _store_uploaded_pdf(upload)
        extracted_text = _extract_pdf_text(stored_path)
        guessed = _guess_invoice_from_pdf_text(extracted_text)
        guessed_items = _guess_invoice_items_from_pdf_text(extracted_text, guessed)
    except Exception as exc:
        flash(f"Načtení PDF se nepodařilo: {exc}", "error")
        return redirect(url_for("import_invoice_pdf_page"))

    guessed_issuer_id = _find_matching_issuer_id(issuers, guessed)
    body = """
    <section class="content-card"><div class="section-head"><div><h2>Schválení údajů z PDF</h2><div class="section-copy">Aplikace předvyplnila údaje z příjmové faktury. Zkontroluj je a potvrď uložení.</div></div></div><form method="post" action="{{ url_for('save_invoice_pdf_page') }}" class="stack"><div class="form-grid-3"><div class="field"><label>Číslo faktury</label><input name="invoice_number" value="{{ guessed.invoice_number }}"></div><div class="field"><label>Fakturující firma</label><select name="issuer_id" required>{% for row in issuers %}<option value="{{ row.id }}" {% if row.id == guessed_issuer_id %}selected{% endif %}>{{ row.company_name }}</option>{% endfor %}</select></div><div class="field"><label>Datum vystavení</label><input type="date" name="issue_date" value="{{ guessed.issue_date }}" required></div><div class="field"><label>Datum splatnosti</label><input type="date" name="due_date" value="{{ guessed.due_date }}" required></div><div class="field"><label>Měna</label><input name="currency" value="{{ guessed.currency or 'CZK' }}"></div><div class="field"><label>Poznámka</label><textarea name="note" rows="4" data-autogrow>{{ guessed.note }}</textarea></div></div><div class="content-card" style="padding:18px;background:var(--panel-2);"><div class="section-head"><div><h2>Odběratel</h2><div class="section-copy">Odběratel se při uložení dohledá nebo založí automaticky.</div></div></div><div class="form-grid"><div class="field"><label>Název</label><input id="customer-name" name="customer_name" value="{{ guessed.customer_name }}" required></div><div class="field"><label>E-mail</label><input id="customer-email" name="customer_email" value="{{ guessed.customer_email }}"></div><div class="field"><label>Telefon</label><input id="customer-phone" name="customer_phone" value="{{ guessed.customer_phone }}"></div><div class="field"><label>IČO</label><div class="lookup-row"><input id="customer-ico" name="customer_ico" value="{{ guessed.customer_ico }}"><button class="button-secondary" type="button" onclick="lookupAres('customer')">Načíst z ARES</button></div><div class="lookup-status" id="customer-lookup-status"></div></div><div class="field"><label>DIČ</label><input id="customer-dic" name="customer_dic" value="{{ guessed.customer_dic }}"></div><div class="field" style="grid-column:1 / -1;"><label>Adresa</label><textarea id="customer-address" name="customer_address" rows="4" data-autogrow>{{ guessed.customer_address }}</textarea></div></div></div><div class="content-card" style="padding:18px;background:var(--panel-2);"><div class="section-head"><div><h2>Položky faktury</h2><div class="section-copy">Import se pokusil rozpoznat více položek. Před uložením je můžeš upravit.</div></div></div><div class="service-table" id="pdfInvoiceServiceRows">{% for item in guessed_items %}<div class="service-row"><div class="field"><label>Popis</label><input name="description[]" value="{{ item.description }}" required></div><div class="field"><label>Množství</label><input name="quantity[]" value="{{ item.quantity or '1' }}"></div><div class="field"><label>Cena vč. DPH</label><input name="price[]" value="{{ item.price or '0' }}"></div><div class="field"><label>DPH %</label><input name="vat[]" value="{{ item.vat_rate or guessed.vat_rate or '21' }}"></div><button class="button-danger" type="button" onclick="removePdfInvoiceServiceRow(this)">Smazat</button></div>{% endfor %}</div><div class="toolbar" style="margin-top:12px;"><button class="button-secondary" type="button" onclick="addPdfInvoiceServiceRow()">Přidat položku</button></div></div><input type="hidden" name="attachment_path" value="{{ attachment_path }}"><div class="toolbar"><button class="button" type="submit">Schválit a uložit fakturu</button><a class="button-link" href="{{ url_for('invoices_page') }}">Zrušit</a></div></form></section><section class="content-card"><div class="section-head"><div><h2>Náhled extrahovaného textu</h2><div class="section-copy">Pomáhá při kontrole správnosti informací z PDF.</div></div></div><pre style="white-space:pre-wrap;color:var(--muted);font:12px/1.5 Consolas,monospace;max-height:320px;overflow:auto">{{ extracted_text }}</pre></section><script>function addPdfInvoiceServiceRow(){const container=document.getElementById('pdfInvoiceServiceRows');const row=document.createElement('div');row.className='service-row';row.innerHTML='<div class=\"field\"><label>Popis</label><input name=\"description[]\" required></div><div class=\"field\"><label>Množství</label><input name=\"quantity[]\" value=\"1\"></div><div class=\"field\"><label>Cena vč. DPH</label><input name=\"price[]\" value=\"0\"></div><div class=\"field\"><label>DPH %</label><input name=\"vat[]\" value=\"21\"></div><button class=\"button-danger\" type=\"button\" onclick=\"removePdfInvoiceServiceRow(this)\">Smazat</button>';container.appendChild(row);}function removePdfInvoiceServiceRow(button){const container=document.getElementById('pdfInvoiceServiceRows');if(container.children.length===1)return;button.parentElement.remove();}</script>
    """
    return render_page("Schválení PDF faktury", "Potvrzení údajů načtených z PDF.", body, "invoices", guessed=guessed, guessed_items=guessed_items, guessed_issuer_id=guessed_issuer_id, issuers=issuers, attachment_path=str(stored_path), extracted_text=extracted_text, ares_script=ARES_LOOKUP_SCRIPT)


@app.post("/invoices/save-from-pdf")
def save_invoice_pdf_page():
    conn = open_db()
    try:
        customer_name = request.form.get("customer_name", "").strip()
        if not customer_name:
            flash("Odběratel je povinný.", "error")
            return redirect(url_for("import_invoice_pdf_page"))
        issuer_id = parse_int_field(request.form.get("issuer_id", "0"), 0)
        if issuer_id <= 0:
            flash("Vyber fakturující firmu.", "error")
            return redirect(url_for("import_invoice_pdf_page"))
        issue_date = parse_date(request.form.get("issue_date", ""), default=date.today())
        due_date = parse_date(request.form.get("due_date", ""), default=issue_date)
        due_days = max(0, (due_date - issue_date).days)
        currency = request.form.get("currency", "CZK").strip() or "CZK"
        note = request.form.get("note", "").strip()
        invoice_number = request.form.get("invoice_number", "").strip()
        customer_id = _find_or_create_customer_for_invoice_pdf(
            conn,
            name=customer_name,
            email=request.form.get("customer_email", "").strip(),
            phone=request.form.get("customer_phone", "").strip(),
            ico=request.form.get("customer_ico", "").strip(),
            dic=request.form.get("customer_dic", "").strip(),
            address=request.form.get("customer_address", "").strip(),
        )
        invoice_id = new_invoice(conn, customer_id=customer_id, issue_date=issue_date, due_days=due_days, currency=currency, note=note, issuer_id=issuer_id)
        if invoice_number:
            conn.execute("UPDATE invoices SET invoice_number = ? WHERE id = ?", (invoice_number, invoice_id))
            conn.commit()
        issuers = list_issuers(conn)
        issuer_row = next((row for row in issuers if int(row["id"]) == issuer_id), None)
        issuer_is_vat_payer = bool(issuer_row["vat_payer"]) if issuer_row is not None else True
        descriptions = request.form.getlist("description[]")
        quantities = request.form.getlist("quantity[]")
        prices = request.form.getlist("price[]")
        vats = request.form.getlist("vat[]")
        if len(vats) < len(descriptions):
            vats.extend(["0"] * (len(descriptions) - len(vats)))
        added = 0
        for desc, qty_raw, price_raw, vat_raw in zip(descriptions, quantities, prices, vats):
            description = str(desc or "").strip()
            if not description:
                continue
            quantity = parse_float_field(qty_raw, 1.0)
            gross_price = parse_amount_text(price_raw)
            vat_rate = parse_float_field(vat_raw, 0.0)
            if issuer_is_vat_payer:
                divisor = 1.0 + (vat_rate / 100.0)
                unit_price = gross_price / divisor if divisor > 0 else gross_price
                item_vat = vat_rate
            else:
                unit_price = gross_price
                item_vat = 0.0
            add_item(conn, invoice_id, description, quantity, unit_price, item_vat)
            added += 1
        if added == 0:
            flash("Faktura musí mít alespoň jednu položku.", "error")
            delete_invoice(conn, invoice_id)
            return redirect(url_for("import_invoice_pdf_page"))
        attachment_path = request.form.get("attachment_path", "").strip()
        if attachment_path:
            _relocate_income_import_pdf(attachment_path, issue_date, invoice_number or str(invoice_id))
        flash(f"Faktura {invoice_number or invoice_id} byla uložena z PDF s {added} položkami.", "info")
        return redirect(url_for("invoice_detail_page", invoice_id=invoice_id))
    except Exception as exc:
        log_app_error("save_invoice_pdf_page", exc)
        flash(f"Uložení faktury z PDF se nepovedlo: {exc}", "error")
        return redirect(url_for("import_invoice_pdf_page"))


@app.route("/invoices/new", methods=["GET", "POST"])
def new_invoice_page() -> str:
    conn = open_db()
    customers = sorted(list_customers(conn), key=lambda row: str(row["name"] or "").lower())
    issuers = list_issuers(conn)
    selected_customer_id = 0
    if request.method == "POST":
        try:
            customer_id = parse_int_field(request.form.get("customer_id", "0"), 0)
            issuer_id = parse_int_field(request.form.get("issuer_id", "0"), 0)
            issue_date = parse_date(request.form.get("issue_date", ""), default=date.today())
            due_days = parse_int_field(request.form.get("due_days", "14"), 14)
            currency = request.form.get("currency", "CZK").strip() or "CZK"
            note = request.form.get("note", "").strip()
            issuer_row = next((row for row in issuers if int(row["id"]) == issuer_id), None)
            issuer_is_vat_payer = bool(issuer_row["vat_payer"]) if issuer_row is not None else True
            if customer_id <= 0 or issuer_id <= 0:
                flash("Vyber odběratele a fakturující firmu.", "error")
                return redirect(url_for("new_invoice_page"))
            invoice_id = new_invoice(conn, customer_id=customer_id, issue_date=issue_date, due_days=due_days, currency=currency, note=note, issuer_id=issuer_id)
            descriptions = request.form.getlist("description[]")
            quantities = request.form.getlist("quantity[]")
            prices = request.form.getlist("price[]")
            vats = request.form.getlist("vat[]")
            if len(vats) < len(descriptions):
                vats.extend(["0"] * (len(descriptions) - len(vats)))
            added = 0
            for desc, qty_raw, price_raw, vat_raw in zip(descriptions, quantities, prices, vats):
                desc = desc.strip()
                if not desc:
                    continue
                qty = parse_float_field(qty_raw, 1.0)
                gross_price = parse_float_field(price_raw, 0.0)
                vat_value = str(vat_raw).replace(",", ".").strip()
                vat_rate = parse_float_field(vat_value, 0.0) if vat_value else 0.0
                if issuer_is_vat_payer:
                    divisor = 1.0 + (vat_rate / 100.0)
                    unit_price = gross_price / divisor if divisor > 0 else gross_price
                    item_vat = vat_rate
                else:
                    unit_price = gross_price
                    item_vat = 0.0
                add_item(conn, invoice_id, desc, qty, unit_price, item_vat)
                added += 1
            if added == 0:
                conn.execute("DELETE FROM invoice_items WHERE invoice_id = ?", (invoice_id,))
                conn.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
                conn.commit()
                flash("Fakturu nelze uložit bez alespoň jedné položky.", "error")
                return redirect(url_for("new_invoice_page"))
            flash(f"Faktura {invoice_id} byla vytvořena s {added} položkami.", "info")
            return redirect(url_for("invoice_detail_page", invoice_id=invoice_id))
        except Exception as exc:
            log_app_error("new_invoice_page", exc)
            flash(f"Uložení faktury se nepovedlo: {exc}", "error")
            return redirect(url_for("new_invoice_page"))

    body = """
    <section class="content-card"><div class="section-head"><div><h2>Nová faktura</h2><div class="section-copy">Webový formulář pro vytvoření faktury s více službami.</div></div></div><form method="post" class="stack"><div class="form-grid-3"><div class="field"><label>Odběratel</label><select name="customer_id" required><option value="" selected>Vyber odběratele</option>{% for row in customers %}<option value="{{ row.id }}">{{ row.name }}</option>{% endfor %}</select></div><div class="field"><label>Fakturující firma</label><select id="invoice-issuer-id" name="issuer_id" required>{% for row in issuers %}<option value="{{ row.id }}" data-vat-payer="{{ 1 if row.vat_payer else 0 }}">{{ row.company_name }}</option>{% endfor %}</select></div><div class="field"><label>Datum vystavení</label><input type="date" name="issue_date" value="{{ today }}" required></div><div class="field"><label>Splatnost za dní</label><input name="due_days" value="14"></div><div class="field"><label>Měna</label><input name="currency" value="CZK"></div><div class="field"><label>Poznámka / projekt / dotace</label><textarea name="note" rows="5" data-autogrow placeholder="Číslo projektu, dotace, smlouva nebo doplňující informace"></textarea></div></div><div class="content-card" style="padding:18px;background:var(--panel-2);"><div class="section-head"><div><h2>Fakturované služby</h2><div class="section-copy">Cena se zadává jako včetně DPH.</div></div><button class="button-secondary" type="button" onclick="addServiceRow()">Přidat službu</button></div><div id="serviceRows" class="service-table"><div class="service-row"><div class="field"><label>Popis</label><input name="description[]" required></div><div class="field"><label>Množství</label><input name="quantity[]" value="1"></div><div class="field"><label>Cena vč. DPH</label><input name="price[]" value="0"></div><div class="field" data-vat-column><label>DPH %</label><input data-vat-input name="vat[]" value="21" placeholder="Neplátce DPH"></div><button class="button-danger" type="button" onclick="removeServiceRow(this)">Smazat</button></div></div></div><div><button class="button" type="submit">Vytvořit fakturu</button></div></form></section>
    <script>
      function invoiceVatEnabled() {
        const issuer = document.getElementById('invoice-issuer-id');
        if (!issuer) return true;
        const selected = issuer.options[issuer.selectedIndex];
        return String(selected?.dataset?.vatPayer || '1') === '1';
      }
      function syncInvoiceVatInputs() {
        const enabled = invoiceVatEnabled();
        document.querySelectorAll('#serviceRows .service-row').forEach(function(row) {
          row.classList.toggle('no-vat', !enabled);
          const vatField = row.querySelector('[data-vat-column]');
          if (vatField) vatField.style.display = enabled ? '' : 'none';
        });
        document.querySelectorAll('[data-vat-input]').forEach(function(input) {
          if (enabled) {
            input.disabled = false;
            input.placeholder = '';
            if ((!String(input.value).trim() || String(input.value).trim() === '0') && String(input.dataset.userValue || '').trim()) {
              input.value = input.dataset.userValue;
            } else if (!String(input.value).trim()) {
              input.value = '21';
            }
          } else {
            input.dataset.userValue = input.value;
            input.value = '';
            input.disabled = true;
            input.placeholder = 'Neplátce DPH';
          }
        });
      }
      function addServiceRow() {
        const container = document.getElementById('serviceRows');
        const row = document.createElement('div');
        row.className = 'service-row';
        row.innerHTML = '<div class="field"><label>Popis</label><input name="description[]" required></div><div class="field"><label>Množství</label><input name="quantity[]" value="1"></div><div class="field"><label>Cena vč. DPH</label><input name="price[]" value="0"></div><div class="field" data-vat-column><label>DPH %</label><input data-vat-input name="vat[]" value="21" placeholder="Neplátce DPH"></div><button class="button-danger" type="button" onclick="removeServiceRow(this)">Smazat</button>';
        container.appendChild(row);
        syncInvoiceVatInputs();
      }
      function removeServiceRow(button) {
        const container = document.getElementById('serviceRows');
        if (container.children.length === 1) return;
        button.parentElement.remove();
      }
      function invoiceFormHasItems() {
        return Array.from(document.querySelectorAll('input[name=\"description[]\"]')).some(function(input) {
          return String(input.value || '').trim() !== '';
        });
      }
      document.addEventListener('DOMContentLoaded', function() {
        const form = document.querySelector('form');
        const issuer = document.getElementById('invoice-issuer-id');
        if (issuer) issuer.addEventListener('change', syncInvoiceVatInputs);
        if (form) {
          form.addEventListener('submit', function(event) {
            if (!invoiceFormHasItems()) {
              event.preventDefault();
              alert('Fakturu nelze uložit bez alespoň jedné položky.');
            }
          });
        }
        syncInvoiceVatInputs();
      });
    </script>
    """
    return render_page("Nová faktura", "Kompletně předělané webové vystavení faktury se službami.", body, "invoice-new", customers=customers, issuers=issuers, today=date.today().isoformat())


@app.route("/invoices/<int:invoice_id>/edit", methods=["GET", "POST"])
def edit_invoice_page(invoice_id: int) -> str:
    conn = open_db()
    invoice, items, _ = get_invoice_detail(conn, invoice_id)
    customers = list_customers(conn)
    issuers = list_issuers(conn)
    if request.method == "POST":
        try:
            customer_id = parse_int_field(request.form.get("customer_id", "0"), 0)
            issuer_id = parse_int_field(request.form.get("issuer_id", "0"), 0)
            issue_date = parse_date(request.form.get("issue_date", ""), default=date.today())
            due_days = parse_int_field(request.form.get("due_days", "14"), 14)
            currency = request.form.get("currency", "CZK").strip() or "CZK"
            note = request.form.get("note", "").strip()
            issuer_row = next((row for row in issuers if int(row["id"]) == issuer_id), None)
            issuer_is_vat_payer = bool(issuer_row["vat_payer"]) if issuer_row is not None else True
            if customer_id <= 0 or issuer_id <= 0:
                flash("Vyber odběratele a fakturující firmu.", "error")
                return redirect(url_for("edit_invoice_page", invoice_id=invoice_id))

            update_invoice(conn, invoice_id, customer_id, issuer_id, issue_date, due_days, currency, note)

            descriptions = request.form.getlist("description[]")
            quantities = request.form.getlist("quantity[]")
            prices = request.form.getlist("price[]")
            vats = request.form.getlist("vat[]")
            if len(vats) < len(descriptions):
                vats.extend(["0"] * (len(descriptions) - len(vats)))
            new_items: list[tuple[str, float, float, float]] = []
            for desc, qty_raw, price_raw, vat_raw in zip(descriptions, quantities, prices, vats):
                desc = desc.strip()
                if not desc:
                    continue
                qty = parse_float_field(qty_raw, 1.0)
                gross_price = parse_float_field(price_raw, 0.0)
                vat_value = str(vat_raw).replace(",", ".").strip()
                vat_rate = parse_float_field(vat_value, 0.0) if vat_value else 0.0
                if issuer_is_vat_payer:
                    divisor = 1.0 + (vat_rate / 100.0)
                    unit_price = gross_price / divisor if divisor > 0 else gross_price
                    item_vat = vat_rate
                else:
                    unit_price = gross_price
                    item_vat = 0.0
                new_items.append((desc, qty, unit_price, item_vat))
            if not new_items:
                flash("Fakturu nelze uložit bez alespoň jedné položky.", "error")
                return redirect(url_for("edit_invoice_page", invoice_id=invoice_id))
            replace_invoice_items(conn, invoice_id, new_items)
            flash("Faktura byla aktualizována.", "info")
            return redirect(url_for("invoice_detail_page", invoice_id=invoice_id))
        except Exception as exc:
            log_app_error(f"edit_invoice_page:{invoice_id}", exc)
            flash(f"Uložení faktury se nepovedlo: {exc}", "error")
            return redirect(url_for("edit_invoice_page", invoice_id=invoice_id))

    issue = parse_date(str(invoice["issue_date"]), default=date.today())
    due = parse_date(str(invoice["due_date"]), default=issue)
    due_days = max((due - issue).days, 0)
    item_rows = []
    issuer_is_vat_payer = bool(invoice["issuer_vat_payer"]) if invoice["issuer_vat_payer"] is not None else True
    for item in items:
        qty = float(item["quantity"])
        unit_price = float(item["unit_price"])
        vat_rate = float(item["vat_rate"])
        if issuer_is_vat_payer:
            gross_price = unit_price * (1.0 + vat_rate / 100.0)
        else:
            gross_price = unit_price
        item_rows.append({"description": item["description"], "quantity": qty, "price": gross_price, "vat": vat_rate})
    if not item_rows:
        item_rows = [{"description": "", "quantity": 1, "price": 0, "vat": 21}]

    body = """
    <section class="content-card"><div class="section-head"><div><h2>Upravit fakturu</h2><div class="section-copy">Uprav odběratele, položky i poznámku projektu nebo dotace.</div></div></div><form method="post" class="stack"><div class="form-grid-3"><div class="field"><label>Odběratel</label><select name="customer_id" required>{% for row in customers %}<option value="{{ row.id }}" {% if row.id == invoice.customer_id %}selected{% endif %}>{{ row.name }}</option>{% endfor %}</select></div><div class="field"><label>Fakturující firma</label><select id="invoice-issuer-id" name="issuer_id" required>{% for row in issuers %}<option value="{{ row.id }}" data-vat-payer="{{ 1 if row.vat_payer else 0 }}" {% if row.id == invoice.issuer_id %}selected{% endif %}>{{ row.company_name }}</option>{% endfor %}</select></div><div class="field"><label>Datum vystavení</label><input type="date" name="issue_date" value="{{ invoice.issue_date }}" required></div><div class="field"><label>Splatnost za dní</label><input name="due_days" value="{{ due_days }}"></div><div class="field"><label>Měna</label><input name="currency" value="{{ invoice.currency }}"></div><div class="field"><label>Poznámka / projekt / dotace</label><textarea name="note" rows="5" data-autogrow placeholder="Číslo projektu, dotace, smlouva nebo doplňující informace">{{ invoice.note or '' }}</textarea></div></div><div class="content-card" style="padding:18px;background:var(--panel-2);"><div class="section-head"><div><h2>Fakturované služby</h2><div class="section-copy">Cena se zadává jako včetně DPH.</div></div><button class="button-secondary" type="button" onclick="addServiceRow()">Přidat službu</button></div><div id="serviceRows" class="service-table">{% for item in item_rows %}<div class="service-row{% if not issuer_is_vat_payer %} no-vat{% endif %}"><div class="field"><label>Popis</label><input name="description[]" value="{{ item.description }}" required></div><div class="field"><label>Množství</label><input name="quantity[]" value="{{ item.quantity }}"></div><div class="field"><label>Cena vč. DPH</label><input name="price[]" value="{{ '%.2f'|format(item.price) }}"></div><div class="field" data-vat-column {% if not issuer_is_vat_payer %}style="display:none"{% endif %}><label>DPH %</label><input data-vat-input name="vat[]" value="{{ '' if not issuer_is_vat_payer else '%.1f'|format(item.vat) }}" placeholder="Neplátce DPH"></div><button class="button-danger" type="button" onclick="removeServiceRow(this)">Smazat</button></div>{% endfor %}</div></div><div class="toolbar"><button class="button" type="submit">Uložit fakturu</button><a class="button-link" href="{{ url_for('invoice_detail_page', invoice_id=invoice.id) }}">Zpět</a></div></form></section>
    <script>
      function invoiceVatEnabled() {
        const issuer = document.getElementById('invoice-issuer-id');
        if (!issuer) return true;
        const selected = issuer.options[issuer.selectedIndex];
        return String(selected?.dataset?.vatPayer || '1') === '1';
      }
      function syncInvoiceVatInputs() {
        const enabled = invoiceVatEnabled();
        document.querySelectorAll('#serviceRows .service-row').forEach(function(row) {
          row.classList.toggle('no-vat', !enabled);
          const vatField = row.querySelector('[data-vat-column]');
          if (vatField) vatField.style.display = enabled ? '' : 'none';
        });
        document.querySelectorAll('[data-vat-input]').forEach(function(input) {
          if (enabled) {
            input.disabled = false;
            input.placeholder = '';
            if ((!String(input.value).trim() || String(input.value).trim() === '0') && String(input.dataset.userValue || '').trim()) {
              input.value = input.dataset.userValue;
            } else if (!String(input.value).trim()) {
              input.value = '21';
            }
          } else {
            input.dataset.userValue = input.value;
            input.value = '';
            input.disabled = true;
            input.placeholder = 'Neplátce DPH';
          }
        });
      }
      function addServiceRow() {
        const container = document.getElementById('serviceRows');
        const row = document.createElement('div');
        row.className = 'service-row';
        row.innerHTML = '<div class="field"><label>Popis</label><input name="description[]" required></div><div class="field"><label>Množství</label><input name="quantity[]" value="1"></div><div class="field"><label>Cena vč. DPH</label><input name="price[]" value="0"></div><div class="field" data-vat-column><label>DPH %</label><input data-vat-input name="vat[]" value="21" placeholder="Neplátce DPH"></div><button class="button-danger" type="button" onclick="removeServiceRow(this)">Smazat</button>';
        container.appendChild(row);
        syncInvoiceVatInputs();
      }
      function removeServiceRow(button) {
        const container = document.getElementById('serviceRows');
        if (container.children.length === 1) return;
        button.parentElement.remove();
      }
      function invoiceFormHasItems() {
        return Array.from(document.querySelectorAll('input[name=\"description[]\"]')).some(function(input) {
          return String(input.value || '').trim() !== '';
        });
      }
      document.addEventListener('DOMContentLoaded', function() {
        const form = document.querySelector('form');
        const issuer = document.getElementById('invoice-issuer-id');
        if (issuer) issuer.addEventListener('change', syncInvoiceVatInputs);
        if (form) {
          form.addEventListener('submit', function(event) {
            if (!invoiceFormHasItems()) {
              event.preventDefault();
              alert('Fakturu nelze uložit bez alespoň jedné položky.');
            }
          });
        }
        syncInvoiceVatInputs();
      });
    </script>
    """
    invoice_form = dict(invoice)
    invoice_form["issue_date"] = issue.isoformat()
    return render_page("Upravit fakturu", "Editace vystavené faktury.", body, "invoices", invoice=invoice_form, customers=customers, issuers=issuers, due_days=due_days, item_rows=item_rows)


@app.route("/invoices/<int:invoice_id>")
def invoice_detail_page(invoice_id: int) -> str:
    conn = open_db()
    invoice, items, totals = get_invoice_detail(conn, invoice_id)
    issuer_is_vat_payer = bool(invoice["issuer_vat_payer"]) if invoice["issuer_vat_payer"] is not None else True
    body = """
    <section class="invoice-top"><div class="content-card"><div class="section-head"><div><h2>Faktura {{ invoice.invoice_number }}</h2><div class="section-copy">Detail faktury, položek a exportu.</div></div><div>{{ status|safe }}</div></div><div class="stack"><div><strong>Firma:</strong> {{ invoice.issuer_name or '-' }}</div><div><strong>Odběratel:</strong> {{ invoice.customer_name }}</div><div><strong>Vystavení:</strong> {{ invoice.issue_date }} | <strong>Splatnost:</strong> {{ invoice.due_date }}</div><div><strong>Poznámka / projekt / dotace:</strong><br>{{ (invoice.note or '-')|replace('\n', '<br>')|safe }}</div></div></div><div class="content-card"><div class="metric-label">K úhradě</div><div class="detail-total">{{ grand_total }}</div><div class="metric-sub">{% if issuer_is_vat_payer %}Základ {{ subtotal }}, DPH {{ vat_total }}{% else %}Nejsem plátce DPH{% endif %}</div><div class="toolbar" style="margin-top:16px;"><a class="button-link" href="{{ url_for('edit_invoice_page', invoice_id=invoice.id) }}">Upravit fakturu</a><form method="post" action="{{ url_for('mark_invoice_paid_page', invoice_id=invoice.id) }}"><button class="button" type="submit">Označit jako zaplacenou</button></form><a class="button-link" href="{{ url_for('export_invoice_page', invoice_id=invoice.id, fmt='pdf') }}">Export PDF</a><a class="button-link" href="{{ url_for('export_invoice_page', invoice_id=invoice.id, fmt='html') }}">Export HTML</a><form method="post" action="{{ url_for('delete_invoice_page', invoice_id=invoice.id) }}" onsubmit="return confirm('Opravdu smazat fakturu?');"><button class="button-danger" type="submit">Smazat</button></form></div></div></section>
    <section class="table-card"><div class="section-head"><div><h2>Položky faktury</h2><div class="section-copy">Všechny služby a částky na dokladu.</div></div></div><table><thead><tr><th>#</th><th>Popis</th><th>Množství</th><th>{% if issuer_is_vat_payer %}Cena bez DPH{% else %}Cena{% endif %}</th>{% if issuer_is_vat_payer %}<th>DPH</th>{% endif %}<th>{% if issuer_is_vat_payer %}Mezisoučet{% else %}Celkem{% endif %}</th></tr></thead><tbody>{% for item in items %}<tr><td>{{ loop.index }}</td><td>{{ item.description_html|safe }}</td><td>{% if not item.is_note %}{{ '%.2f'|format(item.quantity|float) }}{% endif %}</td><td>{% if not item.is_note %}{{ '%.2f'|format(item.unit_price|float) }}{% endif %}</td>{% if issuer_is_vat_payer %}<td>{% if not item.is_note %}{{ '%.1f'|format(item.vat_rate|float) }} %{% endif %}</td>{% endif %}<td>{% if not item.is_note %}{{ '%.2f'|format((item.quantity|float) * (item.unit_price|float)) }}{% endif %}</td></tr>{% endfor %}</tbody></table></section>
    """
    invoice_display = dict(invoice)
    invoice_display["issue_date"] = format_date_cz(invoice["issue_date"])
    invoice_display["due_date"] = format_date_cz(invoice["due_date"])
    items_display = []
    for item in items:
        item_display = dict(item)
        item_display["description_html"] = str(item["description"] or "").replace("\r", "").replace("\n", " ").strip()
        item_display["is_note"] = invoice_item_is_note(item)
        items_display.append(item_display)
    return render_page(f"Faktura {invoice['invoice_number']}", "Detail jedné faktury se všemi dostupnými akcemi.", body, "invoices", invoice=invoice_display, items=items_display, status=status_pill(str(invoice["status"])), subtotal=f"{totals.subtotal:.2f} {invoice['currency']}", vat_total=f"{totals.vat_total:.2f} {invoice['currency']}", grand_total=f"{totals.grand_total:.2f} {invoice['currency']}", issuer_is_vat_payer=issuer_is_vat_payer)

@app.post("/invoices/<int:invoice_id>/paid")
def mark_invoice_paid_page(invoice_id: int):
    conn = open_db()
    mark_paid(conn, invoice_id, date.today())
    flash("Faktura byla označena jako zaplacená.", "info")
    return redirect(url_for("invoice_detail_page", invoice_id=invoice_id))


@app.post("/invoices/<int:invoice_id>/paid-from-list")
def mark_invoice_paid_from_list_page(invoice_id: int):
    conn = open_db()
    invoice_rows = [row for row in list_invoices(conn) if int(row["id"]) == invoice_id]
    if not invoice_rows:
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"ok": False, "message": "Faktura nebyla nalezena."}), 404
        flash("Faktura nebyla nalezena.", "error")
        return redirect(url_for("invoices_page"))

    invoice = invoice_rows[0]
    current_status = str(invoice["status"] or "").strip().lower()
    if current_status not in {"draft", "sent", "overdue"}:
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"ok": False, "message": "Tento stav nelze z přehledu označit jako zaplacený."}), 400
        flash("Tento stav nelze z přehledu označit jako zaplacený.", "error")
        return redirect(url_for("invoices_page"))

    mark_paid(conn, invoice_id, date.today())
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"ok": True, "status_html": status_pill("paid")})
    flash("Faktura byla z přehledu označena jako zaplacená.", "info")
    return redirect(url_for("invoices_page"))


@app.post("/invoices/<int:invoice_id>/delete")
def delete_invoice_page(invoice_id: int):
    conn = open_db()
    delete_invoice(conn, invoice_id)
    flash("Faktura byla smazána.", "info")
    return redirect(url_for("invoices_page"))


@app.route("/invoices/<int:invoice_id>/export/<fmt>")
def export_invoice_page(invoice_id: int, fmt: str):
    conn = open_db()
    invoice, items, _ = get_invoice_detail(conn, invoice_id)
    invoice_number = str(invoice["invoice_number"] or f"invoice_{invoice_id}")
    service_name = str(items[0]["description"]) if items else "sluzba"
    forbidden = '<>:"/\\|?*'
    service_safe = "".join("_" if ch in forbidden else ch for ch in service_name).strip()
    service_safe = "_".join(service_safe.split()) or "sluzba"
    if fmt == "pdf":
        output_path = _invoice_export_path(invoice_number, service_safe, invoice["issue_date"], "pdf")
        export_invoice_pdf(conn, invoice_id, output_path)
        maybe_copy_export(conn, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_path.name)
    if fmt == "html":
        output_path = _invoice_export_path(invoice_number, service_safe, invoice["issue_date"], "html")
        export_invoice_html(conn, invoice_id, output_path)
        maybe_copy_export(conn, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_path.name)
    flash("Neznámý formát exportu.", "error")
    return redirect(url_for("invoice_detail_page", invoice_id=invoice_id))


@app.route("/settings", methods=["GET", "POST"])
def settings_page() -> str:
    conn = open_db()
    if request.method == "POST":
        set_app_setting(conn, "google_drive_folder", request.form.get("google_drive_folder", "").strip())
        set_app_setting(conn, "google_drive_auto_export", "1" if request.form.get("google_drive_auto_export") == "on" else "0")
        flash("Nastavení bylo uloženo.", "info")
        return redirect(url_for("settings_page"))
    body = """
    <section class="grid-2">
      <div class="content-card">
        <div class="section-head"><div><h2>Google Drive</h2><div class="section-copy">Automatické kopírování exportu do synchronizované složky.</div></div></div>
        <form method="post" class="stack">
          <div class="field"><label>Synchronizovaná složka</label><input name="google_drive_folder" value="{{ drive_folder }}" placeholder="Například C:/Users/jmeno/Google Drive/Faktury"></div>
          <div class="toolbar" style="justify-content:space-between;align-items:center;gap:16px;">
            <label style="display:inline-flex;align-items:center;gap:10px;padding:12px 14px;border-radius:16px;background:var(--panel-2);border:1px solid var(--line);font-weight:700;">
              <input type="checkbox" name="google_drive_auto_export" {% if auto_export %}checked{% endif %} style="width:18px;height:18px;accent-color:var(--accent);margin:0;">
              <span>Po exportu kopírovat soubor i do Google Drive</span>
            </label>
            <button class="button" type="submit">Uložit nastavení</button>
          </div>
        </form>
      </div>
      <div class="content-card">
        <div class="section-head"><div><h2>Správa firem</h2><div class="section-copy">Fakturační firmy jsou nově dostupné přímo z nastavení.</div></div></div>
        <div class="stack">
          <div class="muted">Otevři seznam firem, přidej novou fakturující firmu nebo uprav existující údaje.</div>
          <div class="toolbar">
            <a class="button" href="{{ url_for('issuers_page') }}">Firmy</a>
            <a class="button-link" href="{{ url_for('new_issuer_page') }}">Nová firma</a>
          </div>
        </div>
      </div>
    </section>

    <section class="grid-2">
      <div class="content-card"><div class="section-head"><div><h2>Import faktur z Excelu</h2><div class="section-copy">Jeden soubor = jeden rok, každý list = jedna faktura.</div></div></div><form method="post" action="{{ url_for('import_invoices_excel_page') }}" enctype="multipart/form-data" class="stack"><div class="field"><label>Soubor Excel</label><input type="file" name="excel_file" accept=".xlsx,.xlsm" required></div><div class="field"><label>Fakturující firma pro import</label><select name="issuer_id">{% for row in issuers %}<option value="{{ row.id }}">{{ row.company_name }}</option>{% endfor %}</select></div><div><button class="button" type="submit">Importovat faktury</button></div></form></div>
      <div class="content-card"><div class="section-head"><div><h2>Záloha a obnova</h2><div class="section-copy">Zálohuj databázi i PDF doklady do jednoho ZIP balíčku nebo obnov celý stav ze zálohy.</div></div></div><div class="stack"><a class="button" href="{{ url_for('backup_data_page') }}" onclick="return startBackupDownload(this)"><span class="backup-button-idle">Stáhnout kompletní ZIP zálohu</span><span class="backup-button-busy" style="display:none;">Vytvářím ZIP zálohu...</span></a><div id="backupStatusMessage" class="muted" style="display:none;">Aplikace právě připravuje zálohu. Podle velikosti souborů to může chvíli trvat.</div><form method="post" action="{{ url_for('restore_data_page') }}" enctype="multipart/form-data" class="stack"><div class="field"><label>ZIP nebo JSON záloha</label><input type="file" name="backup_file" accept=".zip,.json" required></div><div><button class="button-secondary" type="submit">Obnovit ze zálohy</button></div></form></div></div>
    </section>

    <section class="content-card"><div class="section-head"><div><h2>Správa databáze</h2><div class="section-copy">Nebezpečné operace nad celou databází, doklady i složkou exports.</div></div></div><div class="toolbar"><form method="post" action="{{ url_for('wipe_data_page') }}" onsubmit="return confirm('Opravdu smazat všechny faktury, odběratele, firmy, uložené doklady i celou složku exports včetně záloh?');"><button class="button-danger" type="submit">Smazat vše z databáze i soubory</button></form></div></section>
    """
    return render_page(
        "Nastavení",
        "Google Drive, firmy a datové nástroje webové aplikace.",
        body,
        "settings",
        drive_folder=get_app_setting(conn, "google_drive_folder", ""),
        auto_export=get_app_setting(conn, "google_drive_auto_export", "0") == "1",
        issuers=list_issuers(conn),
    )


def main() -> None:
    host = os.environ.get("APP_HOST", "127.0.0.1").strip() or "127.0.0.1"
    port_raw = os.environ.get("APP_PORT", "8000").strip() or "8000"
    try:
        port = int(port_raw)
    except ValueError:
        port = 8000
    public_url = os.environ.get("APP_PUBLIC_URL", f"http://{host}:{port}").strip() or f"http://{host}:{port}"
    print(f"Fakturace Studio běží na {public_url}")
    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()

