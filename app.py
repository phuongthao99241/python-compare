# compare_app.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Vergleichstool", layout="wide")
st.title("🔍 Vertrags-/Asset-Datenvergleich (Test vs. Prod)")

# Upload-Felder
file_test = st.file_uploader("📂 Test-Datei hochladen", type=["xlsx"], key="test")
file_prod = st.file_uploader("📂 Prod-Datei hochladen", type=["xlsx"], key="prod")

# Hilfsfunktion zum Einlesen und Vorverarbeiten
@st.cache_data
def process_excel(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    df_clean = df_raw.iloc[3:].copy()
    df_clean.columns = df_raw.iloc[0]
    df_clean.reset_index(drop=True, inplace=True)
    df_clean["Vertrags-ID"] = df_clean["Vertrags-ID"].astype(str)
    df_clean["Asset-ID"] = df_clean["Asset-ID"].astype(str)
    df_clean["Key"] = df_clean["Vertrags-ID"] + "_" + df_clean["Asset-ID"]
    return df_clean.set_index("Key")

# Wenn beide Dateien hochgeladen wurden
if file_test and file_prod:
    df_test = process_excel(file_test)
    df_prod = process_excel(file_prod)

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
    
    # Interaktive Tabelle mit Filter
    st.dataframe(df_diff, use_container_width=True)

    # Download-Link generieren
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_diff.to_excel(writer, index=False, sheet_name="Vergleich")
    st.download_button(
        label="📥 Vergleich als Excel herunterladen",
        data=output.getvalue(),
        file_name="vergleichsergebnis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
