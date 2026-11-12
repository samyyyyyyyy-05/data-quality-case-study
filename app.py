from pathlib import Path
from io import BytesIO
import tempfile

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Data Quality Control",
    layout="wide",
)

if "controls_run" not in st.session_state:
    st.session_state.controls_run = False

# ÉTAPE 1 — Fichiers et contrôles techniques

FILES = {
    "positions.xlsx": [
        "ValuationDate", "Portfolio", "ISIN", "SecurityName",
        "AssetType", "Currency", "Nominal", "BookPrice",
        "MarketPrice", "Rating", "MaturityDate",
    ],
    "previous_prices.xlsx": [
        "ISIN", "PreviousDate", "PreviousPrice", "PriceSource",
    ],
    "cashflows.xlsx": [
        "Portfolio", "ISIN", "CashflowDate", "CashflowType",
        "Currency", "Amount", "Status",
    ],
}

DATES = {
    "positions.xlsx": ["ValuationDate", "MaturityDate"],
    "previous_prices.xlsx": ["PreviousDate"],
    "cashflows.xlsx": ["CashflowDate"],
}

NUMBERS = {
    "positions.xlsx": ["Nominal", "BookPrice", "MarketPrice"],
    "previous_prices.xlsx": ["PreviousPrice"],
    "cashflows.xlsx": ["Amount"],
}


def is_missing(value):
    """Reconnaît les cellules vides et les textes null/none."""
    return pd.isna(value) or (
        isinstance(value, str)
        and value.strip().lower() in {"", "nan", "null", "none"}
    )


def load_inputs(folder):
    """Charge les fichiers et vérifie leur structure et leurs formats."""
    data = {}

    for file_name, required_columns in FILES.items():
        path = folder / file_name

        if not path.exists():
            raise ValueError(f"Fichier absent : {file_name}")

        try:
            table = pd.read_excel(path)
        except Exception as error:
            raise ValueError(
                f"Fichier illisible : {file_name}. Détail : {error}"
            ) from error

        if table.empty:
            raise ValueError(f"Fichier vide : {file_name}")

        table.columns = table.columns.astype(str).str.strip()

        missing_columns = [
            column for column in required_columns
            if column not in table.columns
        ]
        if missing_columns:
            raise ValueError(
                f"Colonnes absentes dans {file_name} : {missing_columns}"
            )

        for column in DATES[file_name]:
            for index, value in table[column].items():
                if not is_missing(value) and pd.isna(
                    pd.to_datetime(value, errors="coerce")
                ):
                    raise ValueError(
                        f"Date invalide : {file_name}, ligne {index + 2}, "
                        f"colonne {column}"
                    )

        for column in NUMBERS[file_name]:
            for index, value in table[column].items():
                if not is_missing(value) and pd.isna(
                    pd.to_numeric(value, errors="coerce")
                ):
                    raise ValueError(
                        f"Nombre invalide : {file_name}, ligne {index + 2}, "
                        f"colonne {column}"
                    )

        data[file_name] = table

    return data


# ÉTAPE 2 — Anomalies de qualité

EXCEPTION_COLUMNS = [
    "FileName", "RowNumber", "Portfolio", "ISIN", "SecurityName",
    "AnomalyType", "ObservedValue", "Severity", "Comment",
]

SEVERITIES = {
    "Missing ISIN": "Critical",
    "Missing MarketPrice": "Critical",
    "Missing Rating": "Warning",
    "Duplicate position": "Critical",
    "Duplicate previous price": "Warning",
    "Duplicate cashflow": "Critical",
    "No price history": "Information",
    "Price without position": "Warning",
    "Cashflow without position": "Critical",
    "Nominal zero or negative": "Warning",
    "BookPrice zero or negative": "Critical",
    "MarketPrice zero or negative": "Critical",
    "Maturity before valuation date": "Warning",
    "Price movement above threshold": "Warning",
    "Past cashflow still Forecast": "Warning",
    "Unrecognized AssetType": "Warning",
    "Unrecognized Currency": "Warning",
    "Unrecognized CashflowType": "Warning",
    "Unrecognized Status": "Warning",
}

RECOMMENDED_ACTIONS = {
    "Missing ISIN": "Reject or quarantine the position",
    "Missing MarketPrice": "Block valuation input and request review",
    "Missing Rating": "Request rating review",
    "Duplicate position": "Flag both rows as duplicate",
    "Duplicate previous price": "Review duplicate price history",
    "Duplicate cashflow": "Flag both rows as duplicate",
    "No price history": "Identify as a new title",
    "Price without position": "Review unmatched price",
    "Cashflow without position": "Flag unmatched cashflow",
    "Nominal zero or negative": "Flag zero or negative exposure",
    "BookPrice zero or negative": "Reject or review book price",
    "MarketPrice zero or negative": "Reject or quarantine the price",
    "Maturity before valuation date": "Review maturity date",
    "Price movement above threshold": "Flag price movement for review",
    "Past cashflow still Forecast": "Review stale cashflow status",
    "Unrecognized AssetType": "Request data correction",
    "Unrecognized Currency": "Request data correction",
    "Unrecognized CashflowType": "Request data correction",
    "Unrecognized Status": "Request data correction",
}

def build_exceptions(data):
    """Regroupe les anomalies de qualité dans un tableau."""
    positions = data["positions.xlsx"]
    prices = data["previous_prices.xlsx"]
    cashflows = data["cashflows.xlsx"]
    exceptions = []

    def add(file_name, index, row, anomaly, value, comment):
        exceptions.append({
            "FileName": file_name,
            "RowNumber": index + 2,
            "Portfolio": row.get("Portfolio", ""),
            "ISIN": row.get("ISIN", ""),
            "SecurityName": row.get("SecurityName", ""),
            "AnomalyType": anomaly,
            "ObservedValue": value,
            "Severity": SEVERITIES.get(anomaly, "Warning"),
            "Comment": comment,
        })

    # Cellules vides, y compris les textes "null" et "none"
    for file_name, table in data.items():
        for index, row in table.iterrows():
            for column in table.columns:
                if is_missing(row[column]):
                    observed_value = row.get("SecurityName", "Blank")
                    if not observed_value:
                        observed_value = "Blank"

                    add(
                        file_name,
                        index,
                        row,
                        f"Missing {column}",
                        observed_value,
                        f"{column} est vide.",
                    )

    # Doublons exacts : chaque ligne du doublon est signalée
    duplicate_names = {
        "positions.xlsx": "Duplicate position",
        "previous_prices.xlsx": "Duplicate previous price",
        "cashflows.xlsx": "Duplicate cashflow",
    }

    for file_name, table in data.items():
        duplicates = table[table.duplicated(keep=False)]

        for index, row in duplicates.iterrows():
            add(
                file_name,
                index,
                row,
                duplicate_names[file_name],
                "Exact duplicate",
                "Cette ligne apparaît plusieurs fois.",
            )

    # Positions sans historique de prix et prix sans position
    position_isins = {
        str(value).strip()
        for value in positions["ISIN"]
        if not is_missing(value)
    }
    price_isins = {
        str(value).strip()
        for value in prices["ISIN"]
        if not is_missing(value)
    }

    for index, row in positions.iterrows():
        isin = row["ISIN"]

        if not is_missing(isin) and str(isin).strip() not in price_isins:
            add(
                "positions.xlsx",
                index,
                row,
                "No price history",
                isin,
                "Le titre n'a pas de prix précédent.",
            )

    for index, row in prices.iterrows():
        isin = row["ISIN"]

        if not is_missing(isin) and str(isin).strip() not in position_isins:
            add(
                "previous_prices.xlsx",
                index,
                row,
                "Price without position",
                isin,
                "Prix précédent sans position correspondante.",
            )

    # Cashflows sans position correspondante
    position_keys = {
        (str(row["Portfolio"]).strip(), str(row["ISIN"]).strip())
        for _, row in positions.iterrows()
        if not is_missing(row["Portfolio"]) and not is_missing(row["ISIN"])
    }

    for index, row in cashflows.iterrows():
        portfolio = row["Portfolio"]
        isin = row["ISIN"]

        if not is_missing(portfolio) and not is_missing(isin):
            key = (str(portfolio).strip(), str(isin).strip())

            if key not in position_keys:
                add(
                    "cashflows.xlsx",
                    index,
                    row,
                    "Cashflow without position",
                    f"{key[0]} / {key[1]}",
                    "Le cashflow n'a pas de position correspondante.",
                )

    return pd.DataFrame(
        exceptions,
        columns=EXCEPTION_COLUMNS,
    ).fillna("")


# ÉTAPE 3 — Contrôles métiers

def add_business_checks(loaded_data, exceptions_df):
    """Ajoute les contrôles métiers et prépare les tableaux de résultats."""
    positions = loaded_data["positions.xlsx"]
    prices = loaded_data["previous_prices.xlsx"]
    cashflows = loaded_data["cashflows.xlsx"]
    new_exceptions = []

    valuation_date = pd.to_datetime(
        positions["ValuationDate"],
        errors="coerce",
    ).dropna().iloc[0]

    thresholds = {
        "CLO Debt": 0.05,
        "ABS": 0.05,
        "CLO Equity": 0.15,
    }

    allowed_values = {
        "AssetType": ["CLO Debt", "CLO Equity", "ABS"],
        "Currency": ["EUR", "USD"],
        "CashflowType": ["Coupon", "Principal", "Fee"],
        "Status": ["Forecast", "Confirmed", "Paid"],
    }

    def add(file_name, index, row, anomaly, value, comment):
        new_exceptions.append({
            "FileName": file_name,
            "RowNumber": index + 2,
            "Portfolio": row.get("Portfolio", ""),
            "ISIN": row.get("ISIN", ""),
            "SecurityName": row.get("SecurityName", ""),
            "AnomalyType": anomaly,
            "ObservedValue": value,
            "Severity": SEVERITIES.get(anomaly, "Warning"),
            "Comment": comment,
        })

    # Contrôles des positions
    for index, row in positions.iterrows():
        nominal = pd.to_numeric(row["Nominal"], errors="coerce")

        if pd.notna(nominal) and nominal <= 0:
            add(
                "positions.xlsx",
                index,
                row,
                "Nominal zero or negative",
                nominal,
                "Le nominal est nul ou négatif.",
            )

        for column in ["BookPrice", "MarketPrice"]:
            price = pd.to_numeric(row[column], errors="coerce")

            if pd.notna(price) and price <= 0:
                add(
                    "positions.xlsx",
                    index,
                    row,
                    f"{column} zero or negative",
                    price,
                    f"{column} doit être supérieur à zéro.",
                )

        for column in ["AssetType", "Currency"]:
            if (
                not is_missing(row[column])
                and row[column] not in allowed_values[column]
            ):
                add(
                    "positions.xlsx",
                    index,
                    row,
                    f"Unrecognized {column}",
                    row[column],
                    f"Valeur attendue : {', '.join(allowed_values[column])}.",
                )

        maturity_date = pd.to_datetime(
            row["MaturityDate"],
            errors="coerce",
        )

        if pd.notna(maturity_date) and maturity_date < valuation_date:
            add(
                "positions.xlsx",
                index,
                row,
                "Maturity before valuation date",
                row["MaturityDate"],
                "La maturité précède la date de valorisation.",
            )

    # Contrôles des cashflows
    for index, row in cashflows.iterrows():
        for column in ["CashflowType", "Status"]:
            if (
                not is_missing(row[column])
                and row[column] not in allowed_values[column]
            ):
                add(
                    "cashflows.xlsx",
                    index,
                    row,
                    f"Unrecognized {column}",
                    row[column],
                    f"Valeur attendue : {', '.join(allowed_values[column])}.",
                )

    # Comparaison des prix actuels et précédents
    previous_prices = prices.dropna(
        subset=["ISIN"],
    ).drop_duplicates("ISIN")

    previous_price_by_isin = previous_prices.set_index(
        "ISIN",
    )["PreviousPrice"].to_dict()

    price_movements = []

    for index, row in positions.iterrows():
        isin = row["ISIN"]

        if is_missing(isin) or isin not in previous_price_by_isin:
            continue

        current_price = pd.to_numeric(
            row["MarketPrice"],
            errors="coerce",
        )
        previous_price = pd.to_numeric(
            previous_price_by_isin[isin],
            errors="coerce",
        )

        if (
            pd.isna(current_price)
            or pd.isna(previous_price)
            or current_price <= 0
            or previous_price <= 0
        ):
            continue

        change = (current_price - previous_price) / previous_price
        asset_type = row["AssetType"]
        threshold = thresholds.get(asset_type)

        price_movements.append({
            "Portfolio": row["Portfolio"],
            "ISIN": isin,
            "AssetType": asset_type,
            "CurrentPrice": current_price,
            "PreviousPrice": previous_price,
            "AbsoluteChange": current_price - previous_price,
            "ChangePercent": change,
            "Threshold": threshold,
        })

        if threshold is not None and abs(change) > threshold:
            add(
                "positions.xlsx / previous_prices.xlsx",
                index,
                row,
                "Price movement above threshold",
                f"{change:.2%}",
                f"Seuil applicable : {threshold:.0%}.",
            )

    # Cashflows passés encore indiqués comme Forecast
    for index, row in cashflows.iterrows():
        cashflow_date = pd.to_datetime(
            row["CashflowDate"],
            errors="coerce",
        )

        if (
            pd.notna(cashflow_date)
            and cashflow_date < valuation_date
            and row["Status"] == "Forecast"
        ):
            add(
                "cashflows.xlsx",
                index,
                row,
                "Past cashflow still Forecast",
                row["CashflowDate"],
                "La date est passée mais le statut est encore Forecast.",
            )

    # Cashflows prévus dans les 30 jours
    cashflows_30d = cashflows.copy()
    cashflows_30d["CashflowDate"] = pd.to_datetime(
        cashflows_30d["CashflowDate"],
        errors="coerce",
    )

    upcoming_cashflows = cashflows_30d[
        cashflows_30d["CashflowDate"].between(
            valuation_date,
            valuation_date + pd.Timedelta(days=30),
        )
        & cashflows_30d["Status"].isin(["Forecast", "Confirmed"])
    ].copy()

    exceptions_df = pd.concat(
        [exceptions_df, pd.DataFrame(new_exceptions)],
        ignore_index=True,
    )

    return (
        exceptions_df,
        pd.DataFrame(price_movements),
        upcoming_cashflows,
    )


# INTERFACE STREAMLIT

st.title("Data Quality Control")
st.write("Chargez les trois fichiers Excel du case study.")

files = st.file_uploader(
    "Sélectionnez les fichiers",
    type=["xlsx"],
    accept_multiple_files=True,
)

if files:
    files_by_name = {file.name: file for file in files}

    st.write("Fichiers sélectionnés :")
    for file in files:
        st.write(f"- {file.name}")

    if st.button("Lancer les contrôles"):
        st.session_state.controls_run = True

    if st.session_state.controls_run:
        required_files = [
            "positions.xlsx",
            "previous_prices.xlsx",
            "cashflows.xlsx",
        ]

        missing_files = [
            name for name in required_files
            if name not in files_by_name
        ]

        if missing_files:
            st.error("Il manque : " + ", ".join(missing_files))
        else:
            try:
                with tempfile.TemporaryDirectory() as temp_folder:
                    temp_folder = Path(temp_folder)

                    for name in required_files:
                        (temp_folder / name).write_bytes(
                            files_by_name[name].getvalue()
                        )

                    loaded_data = load_inputs(temp_folder)
                    exceptions_df = build_exceptions(loaded_data)

                    (
                        exceptions_df,
                        price_movements_df,
                        upcoming_cashflows_df,
                    ) = add_business_checks(
                        loaded_data,
                        exceptions_df,
                    )

                    exceptions_df["RecommendedAction"] = (exceptions_df["AnomalyType"].map(RECOMMENDED_ACTIONS).fillna("Review exception"))

                # Résumé pour l'application
                critical_count = int(
                    (exceptions_df["Severity"] == "Critical").sum()
                )
                warning_count = int(
                    (exceptions_df["Severity"] == "Warning").sum()
                )
                information_count = int(
                    (exceptions_df["Severity"] == "Information").sum()
                )
                
                if critical_count > 0:
                    overall_status = "Review required"
                elif warning_count > 0:
                    overall_status = "Review recommended"
                else:
                    overall_status = "No exception detected"                

                valuation_date = pd.to_datetime(
                    loaded_data["positions.xlsx"]["ValuationDate"].iloc[0]
                ).strftime("%Y-%m-%d")

                summary = {
                    "Positions analysées": len(
                        loaded_data["positions.xlsx"]
                    ),
                    "Anomalies détectées": len(exceptions_df),
                    "Critical": critical_count,
                    "Warning": warning_count,
                    "Information": information_count,
                    "Cashflows à 30 jours": len(upcoming_cashflows_df),
                    "Date de valorisation": valuation_date,
                    "Overall status": overall_status,
                }

                for currency, total in loaded_data[
                    "positions.xlsx"
                ].groupby("Currency")["Nominal"].sum().items():
                    summary[
                        f"Nominal total {currency} (doublons inclus)"
                    ] = total

                summary["Périmètre des chiffres"] = (
                    "Lignes sources conservées, y compris les doublons "
                    "et anomalies signalés."
                )

                summary_df = pd.DataFrame(
                    summary.items(),
                    columns=["Indicateur", "Valeur"],
                )

                st.success("Les contrôles sont terminés.")
                
                if overall_status == "Review required":
                    st.warning(
                        "Statut général : Review required. "
                        "Des anomalies Critical nécessitent une revue."
                    )
                elif overall_status == "Review recommended":
                    st.info(
                        "Statut général : Review recommended. "
                        "Des anomalies Warning ont été détectées."
                    )
                else:
                    st.success(
                        "Statut général : No exception detected."
                    )
                    
                st.subheader("Executive Summary")
                st.dataframe(
                    summary_df,
                    hide_index=True,
                    use_container_width=True,
                )

                metric1, metric2, metric3 = st.columns(3)
                metric1.metric(
                    "Positions analysées",
                    len(loaded_data["positions.xlsx"]),
                )
                metric2.metric(
                    "Anomalies détectées",
                    len(exceptions_df),
                )
                metric3.metric(
                    "Cashflows à 30 jours",
                    len(upcoming_cashflows_df),
                )

                sev1, sev2, sev3 = st.columns(3)
                sev1.metric("Critical", critical_count)
                sev2.metric("Warning", warning_count)
                sev3.metric("Information", information_count)

                # Ajouter le type d'actif aux anomalies pour filtrer
                positions = loaded_data["positions.xlsx"]
                asset_types = {}

                for _, row in positions.iterrows():
                    if not is_missing(row["Portfolio"]) and not is_missing(
                        row["ISIN"]
                    ):
                        key = (
                            str(row["Portfolio"]).strip(),
                            str(row["ISIN"]).strip(),
                        )
                        asset_types[key] = row["AssetType"]

                exceptions_view = exceptions_df.copy()
                exceptions_view["AssetType"] = exceptions_view.apply(
                    lambda row: asset_types.get(
                        (
                            str(row["Portfolio"]).strip(),
                            str(row["ISIN"]).strip(),
                        ),
                        "",
                    ),
                    axis=1,
                )

                st.subheader("Anomalies")

                filter1, filter2, filter3 = st.columns(3)

                portfolio_options = ["Tous"] + sorted(
                    str(value)
                    for value in exceptions_view["Portfolio"].dropna().unique()
                    if str(value).strip()
                )
                severity_options = ["Toutes"] + [
                    value
                    for value in ["Critical", "Warning", "Information"]
                    if value in exceptions_view["Severity"].values
                ]
                asset_type_options = ["Tous"] + sorted(
                    str(value)
                    for value in exceptions_view["AssetType"].unique()
                    if str(value).strip()
                )

                selected_portfolio = filter1.selectbox(
                    "Portefeuille",
                    portfolio_options,
                )
                selected_severity = filter2.selectbox(
                    "Sévérité",
                    severity_options,
                )
                selected_asset_type = filter3.selectbox(
                    "Type d’actif",
                    asset_type_options,
                )

                filtered_exceptions = exceptions_view.copy()

                if selected_portfolio != "Tous":
                    filtered_exceptions = filtered_exceptions[
                        filtered_exceptions["Portfolio"]
                        == selected_portfolio
                    ]

                if selected_severity != "Toutes":
                    filtered_exceptions = filtered_exceptions[
                        filtered_exceptions["Severity"]
                        == selected_severity
                    ]

                if selected_asset_type != "Tous":
                    filtered_exceptions = filtered_exceptions[
                        filtered_exceptions["AssetType"]
                        == selected_asset_type
                    ]

                st.dataframe(
                    filtered_exceptions,
                    hide_index=True,
                    use_container_width=True,
                )

                st.subheader("Mouvements de prix")
                st.dataframe(
                    price_movements_df,
                    hide_index=True,
                    use_container_width=True,
                )

                st.subheader("Cashflows à venir (30 jours)")
                st.dataframe(
                    upcoming_cashflows_df,
                    hide_index=True,
                    use_container_width=True,
                )

                # Préparer le statut des mouvements pour le reporting
                movements_df = price_movements_df.copy()
                movements_df["Status"] = movements_df.apply(
                    lambda row: (
                        "Not checked"
                        if pd.isna(row["Threshold"])
                        else "Exception"
                        if abs(row["ChangePercent"]) > row["Threshold"]
                        else "OK"
                    ),
                    axis=1,
                )

                # Nombre d'anomalies par contrôle
                counts = exceptions_df["AnomalyType"].value_counts()
                checks = list(dict.fromkeys([
                    *SEVERITIES.keys(),
                    *counts.index,
                ]))

                quality_df = pd.DataFrame({
                    "Control": checks,
                    "ExceptionsFound": [
                        int(counts.get(check, 0))
                        for check in checks
                    ],
                })

                # Créer le reporting Excel avec cinq onglets
                excel_file = BytesIO()

                sheets = {
                    "Executive Summary": summary_df,
                    "Exceptions": exceptions_df,
                    "Price Movements": movements_df,
                    "Cashflows 30D": upcoming_cashflows_df,
                    "Data Quality": quality_df,
                }

                with pd.ExcelWriter(
                    excel_file,
                    engine="openpyxl",
                ) as writer:
                    for sheet_name, table in sheets.items():
                        table.to_excel(
                            writer,
                            sheet_name=sheet_name,
                            index=False,
                        )
                        sheet = writer.sheets[sheet_name]
                        sheet.freeze_panes = "A2"
                        sheet.auto_filter.ref = sheet.dimensions

                        for cell in sheet[1]:
                            cell.fill = PatternFill(
                                "solid",
                                fgColor="1F4E78",
                            )
                            cell.font = Font(
                                color="FFFFFF",
                                bold=True,
                            )
                            cell.alignment = Alignment(
                                wrap_text=True,
                            )

                        for column in sheet.columns:
                            letter = get_column_letter(
                                column[0].column
                            )
                            width = max(
                                len(str(cell.value or ""))
                                for cell in column
                            )
                            sheet.column_dimensions[letter].width = min(
                                max(width + 2, 14),
                                48,
                            )

                        for column in sheet.iter_cols(min_row=2):
                            header = sheet.cell(
                                1,
                                column[0].column,
                            ).value

                            for cell in column:
                                if header in [
                                    "ChangePercent",
                                    "Threshold",
                                ]:
                                    cell.number_format = "0.00%"
                                elif header == "CashflowDate":
                                    cell.number_format = "yyyy-mm-dd"
                                elif header in [
                                    "Nominal",
                                    "Amount",
                                    "CurrentPrice",
                                    "PreviousPrice",
                                    "AbsoluteChange",
                                ]:
                                    cell.number_format = "#,##0.00"

                    # Surligner les anomalies Critical
                    sheet = writer.sheets["Exceptions"]
                    severity_column = list(
                        exceptions_df.columns
                    ).index("Severity")

                    for row in sheet.iter_rows(min_row=2):
                        if row[severity_column].value == "Critical":
                            for cell in row:
                                cell.fill = PatternFill(
                                    "solid",
                                    fgColor="FCE8E6",
                                )

                st.download_button(
                    label="Télécharger le reporting Excel complet",
                    data=excel_file.getvalue(),
                    file_name="reporting.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

            except Exception as error:
                st.error(
                    f"Les contrôles n'ont pas pu aboutir : {error}"
                )
