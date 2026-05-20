"""Couche de persistance SQLite pour l'extracteur de factures IAM YAZAKI.

Schéma :
    invoices    (1 ligne par facture sauvegardée)
       |
       +-- contracts   (1 ligne par page de contrat)
              |
              +-- line_items (1 ligne par ligne FRAIS MENSUELS / PONCTUELS)

`invoices.raw_json` conserve aussi le JSON brut complet (sauvegarde).
"""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator


DB_PATH = Path(__file__).resolve().parent / "data" / "invoices.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    saved_at TEXT NOT NULL,
    plant TEXT NOT NULL,
    department TEXT NOT NULL,
    project TEXT NOT NULL,
    source_file TEXT,
    client_number TEXT,
    invoice_number TEXT,
    invoice_date TEXT,
    period_start TEXT,
    period_end TEXT,
    total REAL NOT NULL DEFAULT 0,
    frais_abonnement_services REAL DEFAULT 0,
    frais_ponctuels REAL DEFAULT 0,
    montant_ht REAL DEFAULT 0,
    montant_tva REAL DEFAULT 0,
    montant_ttc REAL DEFAULT 0,
    montant_du REAL DEFAULT 0,
    raw_json TEXT NOT NULL,
    UNIQUE (plant, department, project, invoice_number)
);

CREATE TABLE IF NOT EXISTS contracts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL,
    page_number INTEGER,
    document_page TEXT,
    contract_type TEXT,
    phone_number TEXT,
    period_start TEXT,
    period_end TEXT,
    total_contrat REAL NOT NULL DEFAULT 0,
    FOREIGN KEY (invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS line_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contract_id INTEGER NOT NULL,
    section TEXT NOT NULL CHECK (section IN ('mensuel', 'ponctuel')),
    description TEXT,
    amount REAL DEFAULT 0,
    date_start TEXT,
    date_end TEXT,
    FOREIGN KEY (contract_id) REFERENCES contracts(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_invoices_pdp
    ON invoices(plant, department, project);
CREATE INDEX IF NOT EXISTS idx_invoices_invoice_number
    ON invoices(invoice_number);
CREATE INDEX IF NOT EXISTS idx_contracts_invoice
    ON contracts(invoice_id);
CREATE INDEX IF NOT EXISTS idx_line_items_contract
    ON line_items(contract_id);
"""


# --------------------------------------------------------------- low-level helpers
@contextmanager
def get_conn() -> Iterator[sqlite3.Connection]:
    """Connexion gérée : commit auto, rollback en cas d'erreur, FK activées."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    """Crée le schéma s'il n'existe pas."""
    with get_conn() as conn:
        conn.executescript(SCHEMA)


# ------------------------------------------------------------------ public API
def save_invoice(record: dict[str, Any]) -> int:
    """Sauvegarde (ou remplace) une facture complète.

    Si une facture avec la même (Usine, Département, Projet, N° Facture)
    existe déjà, elle est remplacée. Retourne l'id de la facture.
    """
    init_db()
    gs = record.get("global_summary") or {}

    with get_conn() as conn:
        # Remplacement éventuel
        existing = conn.execute(
            "SELECT id FROM invoices "
            "WHERE plant=? AND department=? AND project=? AND invoice_number=?",
            (
                record.get("plant", ""),
                record.get("department", ""),
                record.get("project", ""),
                record.get("invoice_number"),
            ),
        ).fetchone()
        if existing:
            conn.execute("DELETE FROM invoices WHERE id = ?", (existing["id"],))

        cur = conn.execute(
            """
            INSERT INTO invoices (
                saved_at, plant, department, project, source_file,
                client_number, invoice_number, invoice_date,
                period_start, period_end, total,
                frais_abonnement_services, frais_ponctuels,
                montant_ht, montant_tva, montant_ttc, montant_du,
                raw_json
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                record.get("saved_at")
                or datetime.now().isoformat(timespec="seconds"),
                record.get("plant", ""),
                record.get("department", ""),
                record.get("project", ""),
                record.get("source_file"),
                record.get("client_number"),
                record.get("invoice_number"),
                record.get("invoice_date"),
                record.get("period_start"),
                record.get("period_end"),
                float(record.get("total") or 0),
                float(gs.get("frais_abonnement_services") or 0),
                float(gs.get("frais_ponctuels") or 0),
                float(gs.get("montant_ht") or 0),
                float(gs.get("montant_tva") or 0),
                float(gs.get("montant_ttc") or 0),
                float(gs.get("montant_du") or 0),
                json.dumps(record, ensure_ascii=False),
            ),
        )
        invoice_id = int(cur.lastrowid)

        for c in record.get("contracts", []):
            cur2 = conn.execute(
                """
                INSERT INTO contracts (
                    invoice_id, page_number, document_page, contract_type,
                    phone_number, period_start, period_end, total_contrat
                ) VALUES (?,?,?,?,?,?,?,?)
                """,
                (
                    invoice_id,
                    c.get("page_number"),
                    c.get("document_page"),
                    c.get("contract_type"),
                    c.get("phone_number"),
                    c.get("period_start"),
                    c.get("period_end"),
                    float(c.get("total_contrat") or 0),
                ),
            )
            contract_id = int(cur2.lastrowid)

            for section_key, section_name in (
                ("frais_mensuels", "mensuel"),
                ("frais_ponctuels", "ponctuel"),
            ):
                for li in c.get(section_key, []) or []:
                    conn.execute(
                        """
                        INSERT INTO line_items (
                            contract_id, section, description, amount,
                            date_start, date_end
                        ) VALUES (?,?,?,?,?,?)
                        """,
                        (
                            contract_id,
                            section_name,
                            li.get("description"),
                            float(li.get("amount") or 0),
                            li.get("date_start"),
                            li.get("date_end"),
                        ),
                    )

        return invoice_id


def list_invoices() -> list[dict[str, Any]]:
    """Liste légère des factures (sans le JSON brut), triée par date d'enreg."""
    init_db()
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT
                i.id, i.saved_at, i.plant, i.department, i.project,
                i.invoice_number, i.invoice_date,
                i.period_start, i.period_end,
                i.total, i.montant_ht, i.montant_ttc,
                (SELECT COUNT(*) FROM contracts c WHERE c.invoice_id = i.id)
                    AS contracts_count
            FROM invoices i
            ORDER BY i.saved_at DESC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def get_invoice_raw(invoice_id: int) -> dict[str, Any] | None:
    """Retourne le JSON complet (tel que sauvegardé) d'une facture."""
    init_db()
    with get_conn() as conn:
        row = conn.execute(
            "SELECT raw_json FROM invoices WHERE id = ?", (invoice_id,)
        ).fetchone()
        if not row:
            return None
        return json.loads(row["raw_json"])


def delete_invoice(invoice_id: int) -> bool:
    """Supprime une facture (cascade vers contracts/line_items)."""
    init_db()
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
        return cur.rowcount > 0


def distinct_values(field: str) -> list[str]:
    """Valeurs distinctes pour un champ (plant / department / project)."""
    if field not in ("plant", "department", "project"):
        raise ValueError(f"champ non autorisé : {field}")
    init_db()
    with get_conn() as conn:
        rows = conn.execute(
            f"SELECT DISTINCT {field} FROM invoices "
            f"WHERE {field} IS NOT NULL AND {field} != '' "
            f"ORDER BY {field}"
        ).fetchall()
        return [r[0] for r in rows]


def aggregate_total(
    plant: str | None = None,
    department: str | None = None,
    project: str | None = None,
) -> dict[str, Any]:
    """Somme des `total` filtré par usine / département / projet.

    Retourne {'total': float, 'invoices_count': int, 'contracts_count': int}.
    """
    init_db()
    sql = (
        "SELECT COALESCE(SUM(i.total), 0) AS total, "
        "COUNT(i.id) AS invoices_count, "
        "COALESCE(SUM((SELECT COUNT(*) FROM contracts c WHERE c.invoice_id = i.id)), 0) "
        "    AS contracts_count "
        "FROM invoices i WHERE 1=1"
    )
    params: list[Any] = []
    if plant:
        sql += " AND i.plant = ?"
        params.append(plant)
    if department:
        sql += " AND i.department = ?"
        params.append(department)
    if project:
        sql += " AND i.project = ?"
        params.append(project)

    with get_conn() as conn:
        row = conn.execute(sql, params).fetchone()
        return {
            "total": round(float(row["total"] or 0), 2),
            "invoices_count": int(row["invoices_count"] or 0),
            "contracts_count": int(row["contracts_count"] or 0),
        }


def db_stats() -> dict[str, int]:
    """Statistiques rapides : nb factures / contrats / lignes."""
    init_db()
    with get_conn() as conn:
        return {
            "invoices": int(conn.execute("SELECT COUNT(*) FROM invoices").fetchone()[0]),
            "contracts": int(conn.execute("SELECT COUNT(*) FROM contracts").fetchone()[0]),
            "line_items": int(conn.execute("SELECT COUNT(*) FROM line_items").fetchone()[0]),
        }


def db_size_bytes() -> int:
    """Taille du fichier .db (0 si inexistant)."""
    return DB_PATH.stat().st_size if DB_PATH.exists() else 0


if __name__ == "__main__":  # diagnostic rapide
    init_db()
    print(f"DB : {DB_PATH}")
    print(f"Stats : {db_stats()}")
    print(f"Taille : {db_size_bytes()} octets")
