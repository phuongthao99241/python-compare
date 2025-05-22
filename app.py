import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Vergleichstool", layout="wide")
st.title("🔍 Vertrags-/Asset-Datenvergleich (Test vs. Prod)")

# Upload-Felder
file_test = st.file_uploader("📂 Test-Datei hochladen", type=["xlsx"], key="test")
file_prod = st.file_uploader("📂 Prod-Datei hochladen", type=["xlsx"], key="prod")

# Bereinigungsfunktion inkl. Kontonummern-Auffüllung
@st.cache_data
def clean_and_prepare(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Header-Zeilen extrahieren
    header_1 = df_raw.iloc[0]  # Kontenbezeichnung
    header_2 = df_raw.iloc[1]  # Vertrags-ID, Asset-ID etc.
    header_3 = df_raw.iloc[2]  # Kontonummer
    header_4 = df_raw.iloc[3]  # Soll/Haben

    # 🛠️ Leere Kontonummern nach rechts auffüllen
    header_3 = header_3.fillna(method="ffill")

    # Datenbereich extrahieren
    df_data = df_raw.iloc[4:].copy()
    df_data.reset_index(drop=True, inplace=True)

    # Neue Spaltennamen generieren
    columns_combined = []
    for i in range(len(header_1)):
        if i <= 9:  # Metadaten-Spalten (Vertrags-ID bis Währung)
            columns_combined.append(header_2[i])
        else:
            beschreibung = header_1[i]
            if pd.isna(beschreibung):
                beschreibung = columns_combined[i - 1].split(" - ")[0]
            konto_nr = header_3[i]
            soll_haben = header_4[i]
            name = f"{beschreibung} - {konto_nr} - {soll_haben}"
            columns_combined.append(name)

    df_data.columns = columns_combined

    # Key-Spalte zur Identifikation
    df_data["Vertrags-ID"] = df_data["Vertrags-ID"].astype(str)
    df_data["Asset-ID"] = df_data["Asset-ID"].astype(str)
    df_data["Key"] = df_data["Vertrags-ID"] + "_" + df_data["Asset-ID"]

    return df_data.set_index("Key")

# Wenn beide Dateien vorhanden sind
if file_test and file_prod:
    df_test = clean_and_prepare(file_test)
    df_prod = clean_and_prepare(file_prod)

    # 👉 Download bereinigter Dateien
    col1, col2 = st.columns(2)
    with col1:
        output_test = io.BytesIO()
        with pd.ExcelWriter(output_test, engine="xlsxwriter") as writer:
            df_test.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Test")
        st.download_button(
            label="⬇️ Bereinigte Test-Datei herunterladen",
            data=output_test.getvalue(),
            file_name="bereinigt_test.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        output_prod = io.BytesIO()
        with pd.ExcelWriter(output_prod, engine="xlsxwriter") as writer:
            df_prod.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Prod")
        st.download_button(
            label="⬇️ Bereinigte Prod-Datei herunterladen",
            data=output_prod.getvalue(),
            file_name="bereinigt_prod.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Vergleich
    all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
    common_cols = df_test.columns.intersection(df_prod.columns).difference(["Vertrags-ID", "Asset-ID"])

    results = []
    for key in all_keys:
        row_result = {
            "Vertrags-ID": key.split("_")[0],
            "Asset-ID": "_".join(key.split("_")[1:])
        }

        if key not in df_test.index:
            row_result["Unterschiede"] = "Nur in Prod"
        elif key not in df_prod.index:
            row_result["Unterschiede"] = "Nur in Test"
        else:
            diffs = []
            for col in common_cols:
                val_test = df_test.loc[key, col]
                val_prod = df_prod.loc[key, col]

                if isinstance(val_test, pd.Series):
                    val_test = val_test.iloc[0]
                if isinstance(val_prod, pd.Series):
                    val_prod = val_prod.iloc[0]

                if pd.isna(val_test) and pd.isna(val_prod):
                    continue
                elif pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                    diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")

            row_result["Unterschiede"] = "; ".join(diffs) if diffs else "Keine"

        results.append(row_result)

    df_diff = pd.DataFrame(results)

    st.success(f"✅ Vergleich abgeschlossen. {len(df_diff)} Zeilen analysiert.")
    st.dataframe(df_diff, use_container_width=True)

    # Vergleichs-Datei zum Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_diff.to_excel(writer, index=False, sheet_name="Vergleich")
    st.download_button(
        label="📥 Vergleichsergebnis herunterladen",
        data=output.getvalue(),
        file_name="vergleichsergebnis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
