"""Utilitaire de calcul des dépenses.

Étant donné une facture (JSON ou ligne de DB) et un budget, retourne le
budget restant.

Règle métier :
    dépenses = budget - total
où `total` est la somme du "Total Contrat" de chaque page de contrat
(intersection de toutes les lignes sauf la première / récap. global
avec la colonne "Total Contrat").
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def _coerce_total(invoice_json: dict[str, Any]) -> float:
    """Extrait le total d'un objet JSON, avec repli sur le calcul."""
    if "total" in invoice_json and isinstance(invoice_json["total"], (int, float)):
        return float(invoice_json["total"])
    contracts = invoice_json.get("contracts", [])
    return round(
        sum(float(c.get("total_contrat", 0) or 0) for c in contracts), 2
    )


def calculate_expenses(
    invoice_json: dict[str, Any] | str | Path,
    budget: float,
) -> dict[str, float]:
    """Calcule budget − total.

    `invoice_json` peut être un dict, une chaîne JSON ou un chemin vers
    un fichier .json. Retourne {'budget', 'total', 'expenses'}.
    """
    if isinstance(invoice_json, (str, Path)):
        path = Path(invoice_json)
        # On considère que c'est un chemin si ça ressemble à un chemin
        # (a une extension OU contient un séparateur) — pour éviter de
        # tenter de parser un nom de fichier inexistant comme JSON brut.
        looks_like_path = (
            isinstance(invoice_json, Path)
            or path.suffix.lower() == ".json"
            or "/" in str(invoice_json)
            or "\\" in str(invoice_json)
        )
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
        elif looks_like_path:
            raise FileNotFoundError(
                f"Fichier JSON introuvable : {invoice_json!s}"
            )
        else:
            data = json.loads(str(invoice_json))
    else:
        data = invoice_json

    total = _coerce_total(data)
    budget = float(budget)
    return {
        "budget": budget,
        "total": total,
        "expenses": round(budget - total, 2),
    }


def calculate_expenses_from_db(
    budget: float,
    plant: str | None = None,
    department: str | None = None,
    project: str | None = None,
) -> dict[str, Any]:
    """Agrège le total depuis la base SQLite, puis calcule budget − total.

    Filtres optionnels : usine / département / projet.
    """
    from database import aggregate_total

    agg = aggregate_total(plant=plant, department=department, project=project)
    budget = float(budget)
    expenses = round(budget - agg["total"], 2)
    return {
        "budget": budget,
        "total": agg["total"],
        "expenses": expenses,
        "invoices_count": agg["invoices_count"],
        "contracts_count": agg["contracts_count"],
        "filters": {"plant": plant, "department": department, "project": project},
    }


def _print_available_json_files() -> None:
    data_dir = Path(__file__).resolve().parent / "data"
    files = sorted(
        [p for p in data_dir.glob("*.json") if p.name != "invoices.json"]
    ) if data_dir.exists() else []
    if not files:
        print("(Aucun fichier .json sous ./data/)")
        return
    print("Fichiers JSON disponibles :")
    for p in files:
        print(f"  data/{p.name}")


def _print_db_values() -> None:
    try:
        from database import distinct_values, db_stats
    except Exception as exc:  # noqa: BLE001
        print(f"(Erreur d'accès à la base : {exc})")
        return
    stats = db_stats()
    print(f"Base SQLite : {stats['invoices']} facture(s), "
          f"{stats['contracts']} contrat(s), {stats['line_items']} ligne(s).")
    for field, label in (
        ("plant", "Usines"),
        ("department", "Départements"),
        ("project", "Projets"),
    ):
        vals = distinct_values(field)
        if vals:
            print(f"  {label:<13} : {vals}")
        else:
            print(f"  {label:<13} : (aucun)")


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description=(
            "Calcule le budget restant à partir d'un fichier JSON, "
            "d'une facture en base SQLite, ou d'un agrégat SQLite."
        ),
        epilog=(
            "Exemples :\n"
            "  python expenses.py list\n"
            "  python expenses.py json data/Meknes_IT_Mobile_Telecom_0001848477042026.json 60000\n"
            "  python expenses.py db 60000 --plant Meknes --department IT --project \"Mobile Telecom\"\n"
            "  python expenses.py db 100000     # tous les enregistrements"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser(
        "list",
        help="Affiche les fichiers JSON disponibles et les valeurs en base.",
    )

    p_json = sub.add_parser(
        "json", help="Calcul depuis un fichier JSON local."
    )
    p_json.add_argument("json_path", help="Chemin vers le fichier .json")
    p_json.add_argument("budget", type=float, help="Budget en MAD")

    p_db = sub.add_parser(
        "db", help="Calcul agrégé depuis la base SQLite."
    )
    p_db.add_argument("budget", type=float, help="Budget en MAD")
    p_db.add_argument("--plant", help="Filtre Usine (ex : Meknes)")
    p_db.add_argument("--department", help="Filtre Département (ex : IT)")
    p_db.add_argument("--project", help="Filtre Projet (ex : 'Mobile Telecom')")

    args = parser.parse_args()

    if args.cmd == "list":
        _print_available_json_files()
        print()
        _print_db_values()
        sys.exit(0)

    if args.cmd == "json":
        try:
            out = calculate_expenses(args.json_path, args.budget)
        except FileNotFoundError as exc:
            print(f"[Erreur] {exc}\n", file=sys.stderr)
            _print_available_json_files()
            sys.exit(1)
        except json.JSONDecodeError as exc:
            print(f"[Erreur] Fichier JSON invalide : {exc}", file=sys.stderr)
            sys.exit(1)
    else:  # db
        out = calculate_expenses_from_db(
            args.budget,
            plant=args.plant,
            department=args.department,
            project=args.project,
        )
        if out["invoices_count"] == 0:
            print(
                "[Avertissement] Aucune facture ne correspond aux filtres. "
                "Voici les valeurs disponibles en base :\n",
                file=sys.stderr,
            )
            _print_db_values()
            print(file=sys.stderr)

    print(json.dumps(out, indent=2, ensure_ascii=False))
