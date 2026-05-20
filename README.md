# YAZAKI · Extracteur de factures IAM

Application web moderne (Python + Streamlit) pour extraire les données
structurées de factures **Maroc Telecom (IAM)** multipages, taguer chaque
facture par **Usine / Département / Projet** (listes déroulantes en
cascade), enregistrer le tout en JSON, et calculer le budget restant
(`dépenses = budget − total`).

> Palette : **noir / blanc / rouge**. Interface intégralement en français.

## Fonctionnalités

- Téléversement glisser-déposer d'une facture IAM PDF multipage
  (testé sur la facture YAZAKI de 86 pages).
- Analyse de chaque page de contrat et extraction de :
  - `N° d'Appel`, type de contrat, période facturée
  - Toutes les lignes des sections `FRAIS MENSUELS` et `FRAIS PONCTUELS`
  - `TOTAL CONTRAT` par contrat
  - Récapitulatif global (Montant HT / TVA / TTC / Dû)
- Calcul automatique du **Total** = somme du `TOTAL CONTRAT` de chaque
  page (intersection de toutes les lignes sauf la première /
  récapitulatif global avec la colonne *Total Contrat*).
- Barre latérale avec **listes en cascade** : Usine → Département →
  Projet (modifiables dans `config.json`).
- Enregistrement dans une **base SQLite** (`data/invoices.db`) avec un
  schéma normalisé (`invoices` → `contracts` → `line_items`), **plus**
  un fichier JSON par facture sous `data/` et l'ajout à l'index
  principal `data/invoices.json` (sauvegarde / portabilité).
- Onglet **Calculer les dépenses** : 5 sources possibles
  - **agrégat depuis SQLite** filtré par Usine / Département / Projet
    (somme du total sur N factures),
  - facture choisie depuis SQLite,
  - fichier JSON enregistré sous `data/`,
  - JSON téléversé à la volée,
  - facture en cours d'extraction.
- Onglet **Historique** : tableau de toutes les factures en base
  (totaux, dates, catégorisation), téléchargement JSON et suppression
  ligne par ligne. Statistiques live (nb. factures / contrats / lignes
  / taille DB).

## Installation

```bash
pip install -r requirements.txt
```

## Lancement de l'application

```bash
streamlit run app.py
```

Ouvrez ensuite **http://localhost:8501** dans votre navigateur.

1. **Barre latérale** → choisir **Usine → Département → Projet**.
2. Onglet **Extraire la facture** → déposer le PDF IAM, cliquer sur
   **Extraire les données**.
3. Vérifier le tableau des contrats et le Total calculé.
4. Cliquer **Enregistrer** (ou **Télécharger le JSON**).
5. Onglet **Calculer les dépenses** → choisir le JSON, saisir un
   budget, lire le budget restant.

## Tests rapides en ligne de commande

Analyser une facture et afficher les totaux :

```bash
python invoice_parser.py "C:/Users/ismai/Downloads/IAM 03-2026.PDF"
```

Lister les fichiers JSON dispo + les valeurs présentes en base (à
faire en premier pour connaître les bons filtres) :

```bash
python expenses.py list
```

Calculer le budget restant à partir d'un JSON enregistré
(remplacer le nom par un fichier réel listé par la commande ci-dessus) :

```bash
python expenses.py json data/Meknes_IT_Mobile_Telecom_0001848477042026.json 60000
```

Calculer depuis la base SQLite (agrégat) en filtrant — **les valeurs
doivent correspondre exactement à ce qui est en base** (la commande
`list` les affiche) :

```bash
python expenses.py db 60000 --plant Meknes --department IT --project "Mobile Telecom"
python expenses.py db 100000                  # tous les enregistrements
```

> Note : les libellés stockés sont ceux choisis dans les listes
> déroulantes au moment de la sauvegarde. Si vous avez sauvegardé avant
> la traduction de `config.json`, vos anciennes factures portent les
> anciens libellés (`Meknes`, `IT`, …) — `expenses.py list` vous
> montrera toujours les libellés réellement utilisés.

Stats rapides de la base :

```bash
python database.py
```

## Modifier les listes déroulantes

Ouvrez `config.json` et ajoutez / supprimez des usines, départements
ou projets. Le mécanisme de cascade est automatique : les départements
affichés dépendent de l'usine choisie, et les projets dépendent du
département.

```json
{
  "plants": {
    "Meknès": {
      "Informatique": ["Téléphonie Mobile", "Internet & Lignes Fixes"],
      "Ressources Humaines": ["Recrutement", "Formation"]
    },
    "Tanger": {
      "Informatique": ["Téléphonie Mobile"]
    }
  }
}
```

## Utiliser `calculate_expenses` depuis Python

```python
from expenses import calculate_expenses

result = calculate_expenses(
    "data/Meknes_Informatique_Telephonie_Mobile_0001848477042026.json",
    60000,
)
print(result)
# {'budget': 60000.0, 'total': 48261.83, 'expenses': 11738.17}
```

La fonction accepte un chemin, une chaîne JSON ou un dict.

## Structure du projet

```
yzaki/
├── app.py               # Application web Streamlit
├── invoice_parser.py    # PDF -> données structurées
├── database.py          # Persistance SQLite (invoices/contracts/line_items)
├── expenses.py          # budget - total (JSON ou agrégat SQLite)
├── config.json          # usines / départements / projets
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml      # thème (noir / blanc / rouge)
└── data/                # créé au premier enregistrement
    ├── invoices.db      # base SQLite principale
    ├── invoices.json    # index JSON principal (sauvegarde)
    └── <usine>_<département>_<projet>_<n°facture>.json
```

### Schéma SQLite

```
invoices (id, saved_at, plant, department, project, source_file,
          client_number, invoice_number, invoice_date,
          period_start, period_end, total,
          frais_abonnement_services, frais_ponctuels,
          montant_ht, montant_tva, montant_ttc, montant_du,
          raw_json)
   └── contracts (id, invoice_id, page_number, document_page,
                  contract_type, phone_number,
                  period_start, period_end, total_contrat)
          └── line_items (id, contract_id, section ['mensuel'|'ponctuel'],
                          description, amount, date_start, date_end)
```

Unicité : `(plant, department, project, invoice_number)`. Sauvegarder
deux fois la même facture remplace l'ancienne (avec cascade sur
`contracts` et `line_items`).
