
import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Hotel Cashflow", layout="wide")
st.title("Hotel Cashflow - Web App v6.3")

uploaded_spese = st.file_uploader("Carica file Spese (.xlsx)", type=["xlsx"], key="spese")
uploaded_fiscale = st.file_uploader("Carica file Incassi Fiscali (.xlsx)", type=["xlsx"], key="fiscale")
uploaded_export = st.file_uploader("Carica file Incassi Export (.xlsx)", type=["xlsx"], key="export")

df_spese, df_fiscale, df_export = None, None, None

if uploaded_spese:
    df_spese = pd.read_excel(uploaded_spese)
    df_spese["Importo"] = pd.to_numeric(df_spese["Imponibile"], errors="coerce").fillna(0) + pd.to_numeric(df_spese["IVA"], errors="coerce").fillna(0)
    st.success("File Spese caricato correttamente.")
    st.dataframe(df_spese.head())

if uploaded_fiscale and uploaded_export:
    df_fiscale = pd.read_excel(uploaded_fiscale)
    df_export = pd.read_excel(uploaded_export)
    st.success("File Incassi caricati correttamente.")
    st.dataframe(df_fiscale.head(1))
    st.dataframe(df_export.head(1))

def esporta_excel():
    output = BytesIO()
    if df_spese is None or df_fiscale is None or df_export is None:
        st.error("Carica tutti e tre i file per procedere.")
        return None

    # Spese
    df_spese["Categoria"] = df_spese["Categoria"].astype(str).str.strip().str.title()
    df_spese["Mese"] = pd.to_datetime(df_spese["Data"], errors="coerce").dt.month_name()

    mesi_tradotti = {
        "January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile",
        "May": "Maggio", "June": "Giugno", "July": "Luglio", "August": "Agosto",
        "September": "Settembre", "October": "Ottobre", "November": "Novembre", "December": "Dicembre"
    }
    df_spese["Mese"] = df_spese["Mese"].map(mesi_tradotti)

    # Incassi unificati
    df_fiscale_clean = df_fiscale.iloc[:, [0, 2]].copy()
    df_fiscale_clean.columns = ["Data", "Importo"]
    df_fiscale_clean["Fonte"] = "Fiscale"

    df_export_clean = df_export.iloc[:, [0, 7, 26]].copy()
    df_export_clean.columns = ["Data", "Corrispettivi", "Pagamenti"]
    df_export_clean["Corrispettivi"] = pd.to_numeric(df_export_clean["Corrispettivi"], errors="coerce").fillna(0)
    df_export_clean["Pagamenti"] = pd.to_numeric(df_export_clean["Pagamenti"], errors="coerce").fillna(0)
    df_export_clean["Importo"] = df_export_clean["Corrispettivi"] + df_export_clean["Pagamenti"]
    df_export_clean = df_export_clean[["Data", "Importo"]]
    df_export_clean["Fonte"] = "Export"

    df_incassi = pd.concat([df_fiscale_clean, df_export_clean], ignore_index=True)
    df_incassi["Data"] = pd.to_datetime(df_incassi["Data"], errors="coerce")
    df_incassi["Mese"] = df_incassi["Data"].dt.month_name().str.capitalize()
    df_incassi["Mese"] = df_incassi["Mese"].map(mesi_tradotti)

    mesi_ordine = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
                   "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
    df_incassi["Mese"] = pd.Categorical(df_incassi["Mese"], categories=mesi_ordine, ordered=True)

    # Cashflow
    spese_mese = df_spese.groupby(["Mese", "Categoria"])["Importo"].sum().unstack(fill_value=0)
    spese_mese = spese_mese.reindex(mesi_ordine).fillna(0)
    spese_mese["Totale Spese"] = spese_mese.sum(axis=1)

    incassi_mese = df_incassi.groupby("Mese")["Importo"].sum().reindex(mesi_ordine).fillna(0).to_frame()
    incassi_mese.columns = ["Totale Incassi"]

    cashflow = spese_mese.join(incassi_mese)
    cashflow["Risultato Netto"] = cashflow["Totale Incassi"] - cashflow["Totale Spese"]
    cashflow["Cumulato"] = cashflow["Risultato Netto"].cumsum()
    cashflow = cashflow.reset_index()

    # Export
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_spese.to_excel(writer, sheet_name="Dettaglio Spese", index=False)
        df_incassi.to_excel(writer, sheet_name="Dettaglio Incassi", index=False)
        cashflow.to_excel(writer, sheet_name="Cashflow Mensile", index=False)

        workbook = writer.book
        euro_fmt = workbook.add_format({'num_format': '€#,##0.00'})
        writer.sheets["Dettaglio Spese"].set_column("C:C", 18, euro_fmt)
        writer.sheets["Dettaglio Incassi"].set_column("B:B", 18, euro_fmt)
        writer.sheets["Cashflow Mensile"].set_column("B:G", 18, euro_fmt)

    output.seek(0)
    return output

if uploaded_spese and uploaded_fiscale and uploaded_export:
    if st.button("Genera ed Esporta Excel"):
        file_excel = esporta_excel()
        if file_excel:
            st.success("File Excel generato correttamente!")
            st.download_button(label="Scarica Excel", data=file_excel, file_name="cashflow_riepilogo_v6.3.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
