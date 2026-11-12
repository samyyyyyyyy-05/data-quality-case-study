# Data Quality Control — Case Study

## Objectif

Cette application Streamlit contrôle trois fichiers Excel contenant des positions, des prix précédents et des cashflows. Elle détecte les anomalies, affiche les résultats et génère un reporting Excel.

## Fichiers du rendu

```text
app.py
positions.xlsx
previous_prices.xlsx
cashflows.xlsx
anomalies_expected.xlsx
reporting.xlsx
README.md
```

- `app.py` : code Python et application Streamlit.
- `anomalies_expected.xlsx` : anomalies volontairement introduites et résultats attendus.
- `reporting.xlsx` : reporting généré par l'application.

## Lancement de l'application

Python doit être installé. Ouvrir un terminal dans le dossier contenant `app.py`
cd [CHEMIN DU DOSSIER EXACT]

puis exécuter :

```bash
python -m pip install pandas openpyxl streamlit
python -m streamlit run app.py
```

L'application s'ouvre normalement dans le navigateur. Sinon, ouvrir :

```text
http://localhost:8501
```

## Utilisation

1. Charger `positions.xlsx`, `previous_prices.xlsx` et `cashflows.xlsx`.
2. Cliquer sur **Lancer les contrôles**.
3. Consulter l'Executive Summary, les anomalies, les mouvements de prix et les cashflows à 30 jours.
4. Filtrer les anomalies par portefeuille, type d'actif ou sévérité.
5. Télécharger le reporting Excel complet.

## Contrôles réalisés

- présence des fichiers et des colonnes obligatoires ;
- validité des dates et des nombres ;
- valeurs manquantes, y compris les textes `null` et `none` ;
- doublons exacts ;
- cohérence des ISIN entre les fichiers ;
- cashflows sans position correspondante ;
- nominal ou prix nul/négatif ;
- valeurs métiers non reconnues ;
- maturité antérieure à la date de valorisation ;
- mouvements de prix supérieurs aux seuils ;
- cashflows passés encore au statut `Forecast` ;
- cashflows attendus dans les 30 jours.

Seuils de variation : 5 % pour `CLO Debt` et `ABS`, 15 % pour `CLO Equity`.

## Reporting généré

Le fichier Excel contient cinq onglets :

- `Executive Summary` ;
- `Exceptions` ;
- `Price Movements` ;
- `Cashflows 30D` ;
- `Data Quality`.

Avec les données fournies, l'application analyse 25 positions et produit 15 lignes d'exception. Celles-ci correspondent aux 13 scénarios attendus : chaque ligne appartenant à un doublon est signalée séparément.

## Hypothèses principales

- la première date de valorisation valide du fichier des positions sert de date de référence ;
- les cashflows sont rapprochés avec le couple `Portfolio` / `ISIN` ;
- les montants EUR et USD restent séparés, sans conversion de devise ;
- les données sources sont conservées : les anomalies sont signalées mais ne sont pas corrigées automatiquement ;
- les cashflows à venir couvrent les 30 jours suivant la date de valorisation et les statuts `Forecast` ou `Confirmed`.

## Limites

- seuls les fichiers `.xlsx` portant les noms attendus sont acceptés ;
- les seuils et valeurs autorisées sont définis dans le code ;
- les quasi-doublons ne sont pas détectés ;
- aucune connexion à une base de données ou à une source de prix externe ;
- l'application est fournie pour une exécution locale et ne dispose pas d'URL déployée.
