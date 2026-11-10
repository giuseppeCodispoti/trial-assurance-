def run():
    import pandas as pd
    import plotly.express as px
    import streamlit as st
    import math
    from io import BytesIO
    from datetime import datetime

    # ------------------------------------------------------------
    # CONFIGURAZIONE PAGINA
    # ------------------------------------------------------------

    orario_caricamento = datetime.now().strftime("%d-%m-%Y")
    st.markdown(f"#### ✅ PROGRAMMAZIONE ASSURANCE – TRIAL del **{orario_caricamento}**")

    # ------------------------------------------------------------
    # UPLOAD DEI DUE FILE
    # ------------------------------------------------------------
    col1, col2 = st.columns(2)

    with col1:
        file_csv = st.file_uploader("Carica il file CSV PIU", type=["csv"])

    with col2:
        file_excel = st.file_uploader("Carica il file Excel EDAM", type=["xlsx", "xls"])

    if file_csv is not None and file_excel is not None:
        try:
            # ----------------------------------------------------
            # 1. LETTURA CSV
            # ----------------------------------------------------
            try:
                df_csv = pd.read_csv(
                    file_csv, sep=";", dtype=str, encoding="utf-8-sig"
                )
                if len(df_csv.columns) == 1:
                    file_csv.seek(0)
                    df_csv = pd.read_csv(
                        file_csv, sep=",", dtype=str, encoding="utf-8-sig"
                    )
            except UnicodeDecodeError:
                file_csv.seek(0)
                df_csv = pd.read_csv(
                    file_csv, sep=";", dtype=str, encoding="latin-1"
                )

            # ----------------------------------------------------
            # 2. LETTURA EXCEL
            # ----------------------------------------------------
            df_imprese = pd.read_excel(file_excel, dtype=str)

            # Pulizia nomi colonne
            df_csv.columns = df_csv.columns.astype(str).str.strip()
            df_imprese.columns = df_imprese.columns.astype(str).str.strip()

            # Controllo colonne obbligatorie
            colonne_csv = ["HashTags", "IdCtt"]
            colonne_excel = ["ID WR", "Codice Centrale", "Impresa"]

            mancanti_csv = [c for c in colonne_csv if c not in df_csv.columns]
            mancanti_excel = [c for c in colonne_excel if c not in df_imprese.columns]

            if mancanti_csv:
                st.error(f"Nel CSV mancano queste colonne: {', '.join(mancanti_csv)}")
                st.stop()

            if mancanti_excel:
                st.error(f"Nel file Excel mancano queste colonne: {', '.join(mancanti_excel)}")
                st.stop()

            # ----------------------------------------------------
            # 3. FILTRO HASHTAGS
            # Tiene le WR con TRIAL oppure LATENZAONFIELD,
            # anche quando sono presenti entrambi.
            # ----------------------------------------------------
            filtro = df_csv["HashTags"].fillna("").str.contains(
                r"trial|LATENZAONFIELD",
                case=False,
                regex=True,
                na=False
            )
            df_trial = df_csv.loc[filtro].copy()

            # Normalizza la chiave WR
            df_trial["IdCtt"] = df_trial["IdCtt"].fillna("").astype(str).str.strip()
            df_imprese["ID WR"] = df_imprese["ID WR"].fillna("").astype(str).str.strip()

            # Elimina chiavi vuote
            df_trial = df_trial[df_trial["IdCtt"] != ""].copy()
            df_imprese = df_imprese[df_imprese["ID WR"] != ""].copy()

            # ----------------------------------------------------
            # 4. MERGE CSV + EXCEL
            # CSV IdCtt = Excel ID WR
            # La colonna Impresa viene presa ESCLUSIVAMENTE dall'Excel.
            # ----------------------------------------------------
            df_imprese_merge = df_imprese[
                ["ID WR", "Codice Centrale", "Impresa"]
            ].copy()

            # Valori nulli della colonna Impresa -> "sociale"
            df_imprese_merge["Impresa"] = df_imprese_merge["Impresa"].fillna("sociale")

            # Evita che le eventuali WR duplicate nell'Excel
            # moltiplichino le righe durante il merge.
            duplicati_excel = int(
                df_imprese_merge.duplicated(subset=["ID WR"], keep=False).sum()
            )
            df_imprese_merge = df_imprese_merge.drop_duplicates(
                subset=["ID WR"], keep="first"
            )

            # Il CSV può contenere una propria colonna "Impresa":
            # la eliminiamo per utilizzare esclusivamente quella dell'Excel.
            df_trial = df_trial.drop(columns=["Impresa"], errors="ignore")

            df_trial = df_trial.merge(
                df_imprese_merge,
                left_on="IdCtt",
                right_on="ID WR",
                how="left",
                indicator=True
            )

            wr_non_abbinate = df_trial.loc[
                df_trial["_merge"] == "left_only", "IdCtt"
            ].nunique()

            df_trial.drop(columns=["_merge"], inplace=True)

            # Anche le WR non abbinate o con Impresa nulla vengono
            # classificate come "sociale".
            df_trial["Impresa"] = df_trial["Impresa"].fillna("sociale")

            # ----------------------------------------------------
            # 5. CALCOLO AT DAL CODICE CENTRALE
            # ----------------------------------------------------
            df_trial["distretto"] = (
                df_trial["Codice Centrale"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str[:3]
            )

            df_trial["at"] = df_trial["distretto"].map({
                "964": "bagnato",
                "965": "votano",
                "966": "bagnato"
            }).fillna("carbone")

            # ----------------------------------------------------
            # 6. IMPRESA
            # I valori nulli vengono valorizzati come "sociale".
            # ----------------------------------------------------
            df_trial["Impresa"] = df_trial["Impresa"].fillna("sociale")

            # ----------------------------------------------------
            # 7. PIVOT AT / IMPRESA
            # Conteggio delle WR uniche abbinate.
            # ----------------------------------------------------
            df_abbinate = df_trial[df_trial["ID WR"].notna()].copy()

            pivot = (
                df_abbinate
                .groupby(["at", "Impresa"])["ID WR"]
                .nunique()
                .reset_index(name="Trial")
            )

            pivot_sorted = pivot.sort_values(
                by=["at", "Impresa"],
                kind="stable"
            ).reset_index(drop=True)

            # Mostra AT una sola volta per gruppo
            pivot_sorted["at"] = pivot_sorted.groupby("at")["at"].transform(
                lambda x: x.mask(x.index != x.index[0], "")
            )

            # Totale complessivo
            totale = pd.DataFrame(
                [["Totale complessivo", "", int(pivot["Trial"].sum())]],
                columns=["at", "Impresa", "Trial"]
            )

            pivot_formattata = pd.concat(
                [pivot_sorted, totale],
                ignore_index=True
            )


            # ----------------------------------------------------
            # 9. TABELLA PROGRAMMAZIONE
            # ----------------------------------------------------
  
            st.dataframe(
                pivot_formattata,
                use_container_width=True,
                hide_index=True
            )
            # ----------------------------------------------------
            # 9B. DOWNLOAD PIVOT COMPLESSIVA
            # ----------------------------------------------------
            output_pivot = BytesIO()

            with pd.ExcelWriter(output_pivot, engine="xlsxwriter") as writer:
                pivot_formattata.to_excel(
                    writer,
                    index=False,
                    sheet_name="PIVOT TRIAL"
                )

                workbook = writer.book
                worksheet = writer.sheets["PIVOT TRIAL"]

                # Formati Excel
                header_format = workbook.add_format({
                    "bold": True,
                    "bg_color": "#FFD700",
                    "border": 1,
                    "align": "center"
                })

                total_format = workbook.add_format({
                    "bold": True,
                    "border": 2,
                    "align": "center",
                    "bg_color": "#ADD8E6"
                })

                border_format = workbook.add_format({"border": 1})
                number_format = workbook.add_format({
                    "border": 1,
                    "align": "center"
                })

                # Intestazioni
                for col_num, value in enumerate(pivot_formattata.columns):
                    worksheet.write(0, col_num, value, header_format)

                # Dati e totale complessivo
                for row_num in range(1, len(pivot_formattata) + 1):
                    is_total = (
                        str(pivot_formattata.iloc[row_num - 1, 0])
                        .strip().lower() == "totale complessivo"
                    )

                    for col_num in range(len(pivot_formattata.columns)):
                        value = pivot_formattata.iloc[row_num - 1, col_num]

                        if is_total:
                            fmt = total_format
                        elif col_num == 2:
                            fmt = number_format
                        else:
                            fmt = border_format

                        worksheet.write(row_num, col_num, value, fmt)

                worksheet.set_column("A:A", 22)
                worksheet.set_column("B:B", 35)
                worksheet.set_column("C:C", 12)
                worksheet.freeze_panes(1, 0)

            output_pivot.seek(0)

            st.download_button(
                label="📥 Scarica Pivot complessiva Trial",
                data=output_pivot,
                file_name="pivot_complessiva_trial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ----------------------------------------------------
            # 10. ESPORTAZIONE EXCEL DEL DETTAGLIO WR FILTRATE
            # Esporta le righe originali del CSV che contengono
            # TRIAL e/o LATENZAONFIELD, arricchite con AT e Impresa.
            # Non esporta la pivot riepilogativa e non include tecnici.
            # ----------------------------------------------------
            colonne_excel_dettaglio = [
                c for c in df_trial.columns
                if c not in ["distretto"]
            ]

            dettaglio_wr = df_trial[colonne_excel_dettaglio].copy()

            output_dettaglio = BytesIO()

            with pd.ExcelWriter(output_dettaglio, engine="xlsxwriter") as writer:
                dettaglio_wr.to_excel(
                    writer,
                    index=False,
                    sheet_name="DETTAGLIO_WR"
                )

                workbook = writer.book
                worksheet = writer.sheets["DETTAGLIO_WR"]

                header_format = workbook.add_format({
                    "bold": True,
                    "bg_color": "#FFD700",
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter"
                })

                border_format = workbook.add_format({
                    "border": 1,
                    "valign": "vcenter"
                })

                # Intestazioni e bordi
                for col_num, value in enumerate(dettaglio_wr.columns):
                    worksheet.write(0, col_num, value, header_format)
                    # Larghezza leggibile, senza allargare eccessivamente il file
                    worksheet.set_column(col_num, col_num, min(max(len(str(value)) + 2, 14), 30))

                for row_num in range(1, len(dettaglio_wr) + 1):
                    for col_num in range(len(dettaglio_wr.columns)):
                        value = dettaglio_wr.iloc[row_num - 1, col_num]

                        # XlsxWriter non accetta NaN o valori infiniti:
                        # li convertiamo in celle vuote prima della scrittura.
                        if pd.isna(value):
                            value = None
                        elif pd.api.types.is_number(value) and not math.isfinite(float(value)):
                            value = None

                        worksheet.write(row_num, col_num, value, border_format)

                worksheet.freeze_panes(1, 0)
                worksheet.autofilter(0, 0, len(dettaglio_wr), len(dettaglio_wr.columns) - 1)

            output_dettaglio.seek(0)

            st.download_button(
                label="📥 Scarica dettaglio WR Trial",
                data=output_dettaglio,
                file_name="dettaglio_wr_trial_latenzaonfield.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ----------------------------------------------------
            # 11. GRAFICO DISTRIBUZIONE TRIAL PER IMPRESA
            # ----------------------------------------------------
            df_pie = pivot.groupby(
                "Impresa", as_index=False
            )["Trial"].sum()

            df_pie = df_pie[df_pie["Trial"] > 0]

            if not df_pie.empty:
                fig = px.pie(
                    df_pie,
                    names="Impresa",
                    values="Trial",
                    title="Distribuzione tt Trial per Impresa"
                )
                st.plotly_chart(fig, use_container_width=True)

            # ----------------------------------------------------


        except Exception as e:
            st.error(f"Errore durante l'elaborazione dei file: {e}")

    elif file_csv is not None or file_excel is not None:
        st.info("Carica entrambi i file per avviare l'elaborazione.")
    else:
        st.info("In attesa del caricamento del CSV WR e del file Excel Risorse.")
