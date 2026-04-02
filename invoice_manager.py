#!/usr/bin/env python3
"""
Jednoduchy program pro spravu faktur a jejich generovani.
Pouziti: python invoice_manager.py --help
"""

from __future__ import annotations

import argparse
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

DB_FILE = Path("invoices.db")


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT,
    address TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    issue_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('draft', 'sent', 'paid', 'overdue')),
    currency TEXT NOT NULL DEFAULT 'CZK',
    note TEXT,
    paid_date TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(customer_id) REFERENCES customers(id)
);

CREATE TABLE IF NOT EXISTS invoice_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL,
    description TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    vat_rate REAL NOT NULL DEFAULT 21,
    FOREIGN KEY(invoice_id) REFERENCES invoices(id)
);
"""


@dataclass
class InvoiceTotals:
    subtotal: float
    vat_total: float
    grand_total: float


def connect_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)
    conn.commit()


def parse_date(input_date: Optional[str], default: Optional[date] = None) -> date:
    if not input_date:
        if default is None:
            raise ValueError("Datum chybi.")
        return default
    try:
        return datetime.strptime(input_date, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("Neplatny format datumu. Pouzij YYYY-MM-DD.") from exc


def add_customer(conn: sqlite3.Connection, name: str, email: str, address: str) -> int:
    cur = conn.execute(
        """
        INSERT INTO customers (name, email, address, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (name.strip(), email.strip(), address.strip(), datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    return int(cur.lastrowid)


def ensure_customer_exists(conn: sqlite3.Connection, customer_id: int) -> None:
    row = conn.execute("SELECT id FROM customers WHERE id = ?", (customer_id,)).fetchone()
    if row is None:
        raise ValueError(f"Zakaznik s ID {customer_id} neexistuje.")


def ensure_invoice_exists(conn: sqlite3.Connection, invoice_id: int) -> None:
    row = conn.execute("SELECT id FROM invoices WHERE id = ?", (invoice_id,)).fetchone()
    if row is None:
        raise ValueError(f"Faktura s ID {invoice_id} neexistuje.")


def new_invoice(
    conn: sqlite3.Connection,
    customer_id: int,
    issue_date: date,
    due_date: date,
    currency: str,
    note: str,
) -> int:
    ensure_customer_exists(conn, customer_id)
    cur = conn.execute(
        """
        INSERT INTO invoices (customer_id, issue_date, due_date, status, currency, note, created_at)
        VALUES (?, ?, ?, 'draft', ?, ?, ?)
        """,
        (
            customer_id,
            issue_date.isoformat(),
            due_date.isoformat(),
            currency.upper().strip(),
            note.strip(),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    conn.commit()
    return int(cur.lastrowid)


def add_item(
    conn: sqlite3.Connection,
    invoice_id: int,
    description: str,
    quantity: float,
    unit_price: float,
    vat_rate: float,
) -> int:
    ensure_invoice_exists(conn, invoice_id)
    cur = conn.execute(
        """
        INSERT INTO invoice_items (invoice_id, description, quantity, unit_price, vat_rate)
        VALUES (?, ?, ?, ?, ?)
        """,
        (invoice_id, description.strip(), quantity, unit_price, vat_rate),
    )
    conn.commit()
    return int(cur.lastrowid)


def compute_totals(conn: sqlite3.Connection, invoice_id: int) -> InvoiceTotals:
    items = conn.execute(
        """
        SELECT quantity, unit_price, vat_rate
        FROM invoice_items
        WHERE invoice_id = ?
        """,
        (invoice_id,),
    ).fetchall()

    subtotal = 0.0
    vat_total = 0.0
    for item in items:
        base = float(item["quantity"]) * float(item["unit_price"])
        vat = base * (float(item["vat_rate"]) / 100.0)
        subtotal += base
        vat_total += vat

    return InvoiceTotals(
        subtotal=round(subtotal, 2),
        vat_total=round(vat_total, 2),
        grand_total=round(subtotal + vat_total, 2),
    )


def list_invoices(conn: sqlite3.Connection, status: Optional[str]) -> list[sqlite3.Row]:
    if status:
        return conn.execute(
            """
            SELECT i.id, c.name AS customer_name, i.issue_date, i.due_date, i.status, i.currency
            FROM invoices i
            JOIN customers c ON c.id = i.customer_id
            WHERE i.status = ?
            ORDER BY i.id DESC
            """,
            (status,),
        ).fetchall()

    return conn.execute(
        """
        SELECT i.id, c.name AS customer_name, i.issue_date, i.due_date, i.status, i.currency
        FROM invoices i
        JOIN customers c ON c.id = i.customer_id
        ORDER BY i.id DESC
        """
    ).fetchall()


def get_invoice_detail(conn: sqlite3.Connection, invoice_id: int) -> tuple[sqlite3.Row, list[sqlite3.Row], InvoiceTotals]:
    invoice = conn.execute(
        """
        SELECT i.*, c.name AS customer_name, c.email AS customer_email, c.address AS customer_address
        FROM invoices i
        JOIN customers c ON c.id = i.customer_id
        WHERE i.id = ?
        """,
        (invoice_id,),
    ).fetchone()

    if invoice is None:
        raise ValueError(f"Faktura s ID {invoice_id} neexistuje.")

    items = conn.execute(
        """
        SELECT id, description, quantity, unit_price, vat_rate
        FROM invoice_items
        WHERE invoice_id = ?
        ORDER BY id ASC
        """,
        (invoice_id,),
    ).fetchall()

    totals = compute_totals(conn, invoice_id)
    return invoice, items, totals


def mark_paid(conn: sqlite3.Connection, invoice_id: int, paid_date: date) -> None:
    ensure_invoice_exists(conn, invoice_id)
    conn.execute(
        """
        UPDATE invoices
        SET status = 'paid', paid_date = ?
        WHERE id = ?
        """,
        (paid_date.isoformat(), invoice_id),
    )
    conn.commit()


def export_invoice_html(conn: sqlite3.Connection, invoice_id: int, output_path: Path) -> Path:
    invoice, items, totals = get_invoice_detail(conn, invoice_id)

    rows_html = "\n".join(
        (
            f"<tr><td>{i+1}</td><td>{item['description']}</td>"
            f"<td>{float(item['quantity']):.2f}</td><td>{float(item['unit_price']):.2f}</td>"
            f"<td>{float(item['vat_rate']):.1f}%</td>"
            f"<td>{float(item['quantity']) * float(item['unit_price']):.2f}</td></tr>"
        )
        for i, item in enumerate(items)
    )

    html = f"""<!DOCTYPE html>
<html lang=\"cs\">
<head>
<meta charset=\"utf-8\" />
<title>Faktura #{invoice['id']}</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 32px; color: #222; }}
h1 {{ margin-bottom: 4px; }}
.meta, .customer {{ margin-bottom: 18px; }}
table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
th {{ background: #f5f5f5; }}
.total {{ margin-top: 16px; width: 300px; margin-left: auto; }}
.total td {{ border: none; padding: 4px 8px; }}
.total .grand {{ font-weight: bold; border-top: 1px solid #aaa; }}
</style>
</head>
<body>
<h1>Faktura #{invoice['id']}</h1>
<div class=\"meta\">
<div><strong>Vystaveno:</strong> {invoice['issue_date']}</div>
<div><strong>Splatnost:</strong> {invoice['due_date']}</div>
<div><strong>Stav:</strong> {invoice['status']}</div>
</div>
<div class=\"customer\">
<div><strong>Odberatel:</strong> {invoice['customer_name']}</div>
<div><strong>Email:</strong> {invoice['customer_email'] or '-'}</div>
<div><strong>Adresa:</strong> {invoice['customer_address'] or '-'}</div>
</div>
<table>
<thead>
<tr><th>#</th><th>Polozka</th><th>Mnozstvi</th><th>Cena/ks</th><th>DPH</th><th>Mezisoucet</th></tr>
</thead>
<tbody>
{rows_html}
</tbody>
</table>
<table class=\"total\">
<tr><td>Zaklad:</td><td>{totals.subtotal:.2f} {invoice['currency']}</td></tr>
<tr><td>DPH:</td><td>{totals.vat_total:.2f} {invoice['currency']}</td></tr>
<tr class=\"grand\"><td>Celkem:</td><td>{totals.grand_total:.2f} {invoice['currency']}</td></tr>
</table>
<p>Poznamka: {invoice['note'] or '-'}</p>
</body>
</html>
"""

    output_path.write_text(html, encoding="utf-8")
    return output_path


def report(conn: sqlite3.Connection, date_from: date, date_to: date) -> list[sqlite3.Row]:
    return conn.execute(
        """
        SELECT i.id, c.name AS customer_name, i.issue_date, i.status, i.currency
        FROM invoices i
        JOIN customers c ON c.id = i.customer_id
        WHERE i.issue_date BETWEEN ? AND ?
        ORDER BY i.issue_date ASC
        """,
        (date_from.isoformat(), date_to.isoformat()),
    ).fetchall()


def cmd_init(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    print(f"Databaze pripravena: {args.db}")


def cmd_add_customer(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    customer_id = add_customer(conn, args.name, args.email, args.address)
    print(f"Zakaznik vytvoren. ID: {customer_id}")


def cmd_new_invoice(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    issue = parse_date(args.issue_date, default=date.today())
    due = issue + timedelta(days=args.due_days)
    invoice_id = new_invoice(conn, args.customer_id, issue, due, args.currency, args.note)
    print(f"Faktura vytvorena. ID: {invoice_id}")


def cmd_add_item(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    item_id = add_item(conn, args.invoice_id, args.description, args.quantity, args.unit_price, args.vat_rate)
    print(f"Polozka pridana. ID polozky: {item_id}")


def cmd_list(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    rows = list_invoices(conn, args.status)
    if not rows:
        print("Zadne faktury.")
        return

    print("ID | Zakaznik | Vystaveno | Splatnost | Stav | Mena")
    print("-" * 72)
    for row in rows:
        print(
            f"{row['id']} | {row['customer_name']} | {row['issue_date']} | "
            f"{row['due_date']} | {row['status']} | {row['currency']}"
        )


def cmd_show(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    invoice, items, totals = get_invoice_detail(conn, args.invoice_id)

    print(f"Faktura #{invoice['id']}")
    print(f"Zakaznik: {invoice['customer_name']}")
    print(f"Vystaveno: {invoice['issue_date']} | Splatnost: {invoice['due_date']} | Stav: {invoice['status']}")
    print("Polozky:")
    if not items:
        print("  (zadne)")
    for idx, item in enumerate(items, start=1):
        line_total = float(item["quantity"]) * float(item["unit_price"])
        print(
            f"  {idx}. {item['description']} - {float(item['quantity']):.2f} ks x "
            f"{float(item['unit_price']):.2f} ({float(item['vat_rate']):.1f}% DPH) = {line_total:.2f}"
        )

    print(f"Zaklad: {totals.subtotal:.2f} {invoice['currency']}")
    print(f"DPH:    {totals.vat_total:.2f} {invoice['currency']}")
    print(f"Celkem: {totals.grand_total:.2f} {invoice['currency']}")


def cmd_mark_paid(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    paid = parse_date(args.paid_date, default=date.today())
    mark_paid(conn, args.invoice_id, paid)
    print(f"Faktura {args.invoice_id} oznacena jako zaplacena ({paid.isoformat()}).")


def cmd_export_html(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    output = Path(args.output) if args.output else Path(f"invoice_{args.invoice_id}.html")
    path = export_invoice_html(conn, args.invoice_id, output)
    print(f"Faktura exportovana do HTML: {path}")


def cmd_report(args: argparse.Namespace) -> None:
    conn = connect_db(Path(args.db))
    init_db(conn)
    date_from = parse_date(args.date_from)
    date_to = parse_date(args.date_to)
    rows = report(conn, date_from, date_to)

    if not rows:
        print("V zadanym obdobi nejsou zadne faktury.")
        return

    print("ID | Zakaznik | Datum | Stav | Mena")
    print("-" * 56)
    for row in rows:
        print(f"{row['id']} | {row['customer_name']} | {row['issue_date']} | {row['status']} | {row['currency']}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Sprava a generovani faktur")
    parser.add_argument("--db", default=str(DB_FILE), help="Cesta k SQLite databazi")

    sub = parser.add_subparsers(dest="command", required=True)

    p_init = sub.add_parser("init", help="Inicializuje databazi")
    p_init.set_defaults(func=cmd_init)

    p_customer = sub.add_parser("add-customer", help="Prida zakaznika")
    p_customer.add_argument("--name", required=True, help="Jmeno nebo firma")
    p_customer.add_argument("--email", default="", help="Email")
    p_customer.add_argument("--address", default="", help="Adresa")
    p_customer.set_defaults(func=cmd_add_customer)

    p_new_inv = sub.add_parser("new-invoice", help="Vytvori novou fakturu")
    p_new_inv.add_argument("--customer-id", type=int, required=True)
    p_new_inv.add_argument("--issue-date", default="", help="Datum vystaveni YYYY-MM-DD")
    p_new_inv.add_argument("--due-days", type=int, default=14, help="Splatnost za X dni")
    p_new_inv.add_argument("--currency", default="CZK", help="Mena, napr. CZK")
    p_new_inv.add_argument("--note", default="", help="Poznamka")
    p_new_inv.set_defaults(func=cmd_new_invoice)

    p_item = sub.add_parser("add-item", help="Prida polozku na fakturu")
    p_item.add_argument("--invoice-id", type=int, required=True)
    p_item.add_argument("--description", required=True)
    p_item.add_argument("--quantity", type=float, required=True)
    p_item.add_argument("--unit-price", type=float, required=True)
    p_item.add_argument("--vat-rate", type=float, default=21.0)
    p_item.set_defaults(func=cmd_add_item)

    p_list = sub.add_parser("list", help="Vypise faktury")
    p_list.add_argument("--status", choices=["draft", "sent", "paid", "overdue"], default=None)
    p_list.set_defaults(func=cmd_list)

    p_show = sub.add_parser("show-invoice", help="Zobrazi detail faktury")
    p_show.add_argument("--invoice-id", type=int, required=True)
    p_show.set_defaults(func=cmd_show)

    p_paid = sub.add_parser("mark-paid", help="Oznaci fakturu jako zaplacenou")
    p_paid.add_argument("--invoice-id", type=int, required=True)
    p_paid.add_argument("--paid-date", default="", help="Datum uhrazeni YYYY-MM-DD")
    p_paid.set_defaults(func=cmd_mark_paid)

    p_export = sub.add_parser("export-html", help="Exportuje fakturu do HTML")
    p_export.add_argument("--invoice-id", type=int, required=True)
    p_export.add_argument("--output", default="", help="Vystupni HTML soubor")
    p_export.set_defaults(func=cmd_export_html)

    p_report = sub.add_parser("report", help="Prehled faktur v intervalu")
    p_report.add_argument("--date-from", required=True, help="Od YYYY-MM-DD")
    p_report.add_argument("--date-to", required=True, help="Do YYYY-MM-DD")
    p_report.set_defaults(func=cmd_report)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    try:
        args.func(args)
    except ValueError as err:
        print(f"Chyba: {err}")
        raise SystemExit(1)
    except sqlite3.Error as err:
        print(f"Databazova chyba: {err}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
