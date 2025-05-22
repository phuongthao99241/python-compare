import streamlit as st
import pandas as pd
import io
import re

# 🌐 Sprachwahl
st.set_page_config(page_title=t["title"], layout="wide")
lang = st.sidebar.selectbox("🌐 Sprache / Language", options=["Deutsch", "English"], index=0)

TEXTE = {
    "Deutsch": {
        "title": "🔍 Vertrags-/Asset-Datenvergleich (Test vs. Prod)",
        "upload_test": "📂 Test-Datei hochladen",
        "upload_prod": "📂 Prod-Datei hochladen",
        "only_test": "⚠️ Spalten **nur in Test**:",
        "only_prod": "⚠️ Spalten **nur in Prod**:",
        "all_match": "✅ Alle Spalten stimmen überein.",
        "download_test": "⬇️ Bereinigte Test-Datei",
        "download_prod": "⬇️ Bereinigte Prod-Datei",
        "compare_done": "✅ Vergleich abgeschlossen. {n} Zeilen analysiert.",
        "download_result": "📥 Vergleichsergebnis herunterladen",
        "contract_id": "Vertrags-ID",
        "asset_id": "Asset-ID"
    },
    "English": {
        "title": "🔍 Contract/Asset Data Comparison (Test vs. Prod)",
        "upload_test": "📂 Upload Test File",
        "upload_prod": "📂 Upload Prod File",
        "only_test": "⚠️ Columns **only in Test**:",
        "only_prod": "⚠️ Columns **only in Prod**:",
        "all_match": "✅ All columns match.",
        "download_test": "⬇️ Download cleaned Test file",
        "download_prod": "⬇️ Download cleaned Prod file",
        "compare_done": "✅ Comparison complete. {n} rows analyzed.",
        "download_result": "📥 Download comparison result",
        "contract_id": "Contract ID",
        "asset_id": "Asset ID"
    }
}
t = TEXTE[lang]

# 🔧 Layout

st.title(t["title"])

# 📂 Upload
file_test = st.file_uploader(t["upload_test"], type=["xlsx"], key="test")
file_prod = st.file_uploader(t["upload_prod"], type=["xlsx"], key="prod")

# 🧹 Bereinigung
@st.cache_data
def clean_and_prepare(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    header_1 = df_raw.iloc[1]  # Kontobezeichnung
    header_2 = df_raw.iloc[2]  # Kontonummer
    header_3 = df_raw.iloc[3]  # Soll/Haben

    header_1 = header_1.fillna(method="ffill")
    header_2 = header_2.fillna(method="ffill")

    df_data = df_raw.iloc[4:].copy()
    df_data.reset_index(drop=True, inplace=True)

    columns_combined = []
    for i in range(len(header_1)):
        if i < 9:
            columns_combined.append(header_1[i])
        else:
            beschreibung = re.sub(r'\s+', ' ', str(header_1[i]).strip())
            konto_nr = str(header_2[i]).strip()
            soll_haben = str(header_3[i]).strip()
            name = f"{beschreibung} - {konto_nr}_IFRS16 - {soll_haben}"
            columns_combined.append(name)

    df_data.columns = columns_combined

    # Schlüssel zur Identifikation
    df_data["Vertrags-ID"] = df_data["Vertrags-ID"].astype(str)
    df_data["Asset-ID"] = df_data["Asset-ID"].astype(str)
    df_data["Key"] = df_data["Vertrags-ID"] + "_" + df_data["Asset-ID"]

    return df_data.set_index("Key")

# Wenn beide Dateien geladen
if file_test and file_prod:
    df_test = clean_and_prepare(file_test)
    df_prod = clean_and_prepare(file_prod)

    # 🔄 Spaltennamen übersetzen
    df_test = df_test.rename(columns={"Vertrags-ID": t["contract_id"], "Asset-ID": t["asset_id"]})
    df_prod = df_prod.rename(columns={"Vertrags-ID": t["contract_id"], "Asset-ID": t["asset_id"]})

    # 🔍 Spaltenvergleich
    columns_test = set(df_test.columns) - {t["contract_id"], t["asset_id"], "Key"}
    columns_prod = set(df_prod.columns) - {t["contract_id"], t["asset_id"], "Key"}

    only_in_test = sorted(columns_test - columns_prod)
    only_in_prod = sorted(columns_prod - columns_test)

    if only_in_test:
        st.warning(t["only_test"])
        st.code("\n".join(only_in_test), language="")

    if only_in_prod:
        st.warning(t["only_prod"])
        st.code("\n".join(only_in_prod), language="")

    if not only_in_test and not only_in_prod:
        st.info(t["all_match"])

    # 📥 Downloads
    col1, col2 = st.columns(2)
    with col1:
        output_test = io.BytesIO()
        with pd.ExcelWriter(output_test, engine="xlsxwriter") as writer:
            df_test.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Test")
        st.download_button(t["download_test"], data=output_test.getvalue(), file_name="bereinigt_test.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col2:
        output_prod = io.BytesIO()
        with pd.ExcelWriter(output_prod, engine="xlsxwriter") as writer:
            df_prod.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Prod")
        st.download_button(t["download_prod"], data=output_prod.getvalue(), file_name="bereinigt_prod.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 🔍 Vergleich
    all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
    common_cols = df_test.columns.intersection(df_prod.columns).difference([t["contract_id"], t["asset_id"]])

    results = []
    for key in all_keys:
        row_result = {
            t["contract_id"]: key.split("_")[0],
            t["asset_id"]: "_".join(key.split("_")[1:])
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

                if isinstance(val_test, pd.Series): val_test = val_test.iloc[0]
                if isinstance(val_prod, pd.Series): val_prod = val_prod.iloc[0]

                if pd.isna(val_test) and pd.isna(val_prod):
                    continue
                elif pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                    diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")

            row_result["Unterschiede"] = "; ".join(diffs) if diffs else "Keine"

        results.append(row_result)

    df_diff = pd.DataFrame(results)
    st.success(t["compare_done"].format(n=len(df_diff)))
    st.dataframe(df_diff, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_diff.to_excel(writer, index=False, sheet_name="Vergleich")
    st.download_button(t["download_result"], data=output.getvalue(), file_name="vergleichsergebnis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
