def run():
    import math
    from datetime import datetime
    from io import BytesIO

    import pandas as pd
    import plotly.express as px
    import streamlit as st

    # ------------------------------------------------------------
    # CONFIGURAZIONE PAGINA
    # ------------------------------------------------------------
    data_elaborazione = datetime.now().strftime("%d-%m-%Y")
    st.markdown(
        f"#### ✅ PROGRAMMAZIONE ASSURANCE – CHIUSURA TRIAL del **{data_elaborazione}**"
    )

    # ------------------------------------------------------------
    # UPLOAD DEI DUE FILE EXCEL
    # ------------------------------------------------------------
    col1, col2 = st.columns(2)

    with col1:
        file_trial = st.file_uploader(
            "Carica il dettaglio Excel WR Trial", type=["xlsx", "xls"]
        )

    with col2:
        file_chiusi = st.file_uploader(
            "Carica il file Excel TT chiusi", type=["xlsx", "xls"]
        )

    def normalizza_id(colonna):
        """Normalizza gli ID WR per evitare differenze dovute a spazi o tipi."""
        return colonna.fillna("").astype(str).str.strip()

    def excel_bytes(dataframe, nome_foglio="RIEPILOGO"):
        """Crea un file Excel formattato e compatibile con valori nulli/infiniti."""
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            dataframe.to_excel(
                writer,
                index=False,
                sheet_name=nome_foglio
            )

            workbook = writer.book
            worksheet = writer.sheets[nome_foglio]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#FFD700",
                "border": 1,
                "align": "center",
                "valign": "vcenter"
            })

            cell_format = workbook.add_format({
                "border": 1,
                "valign": "vcenter"
            })

            number_format = workbook.add_format({
                "border": 1,
                "align": "center"
            })

            percent_format = workbook.add_format({
                "border": 1,
                "align": "center",
                "num_format": '0.0"%"'
            })

            total_format = workbook.add_format({
                "bold": True,
                "bg_color": "#ADD8E6",
                "border": 2,
                "align": "center"
            })

            total_percent_format = workbook.add_format({
                "bold": True,
                "bg_color": "#ADD8E6",
                "border": 2,
                "align": "center",
                "num_format": '0.0"%"'
            })

            # Formattazione intestazioni e larghezza colonne
            for col_num, name in enumerate(dataframe.columns):
                worksheet.write(0, col_num, name, header_format)

                width = min(max(len(str(name)) + 3, 14), 28)
                worksheet.set_column(col_num, col_num, width)

            # Formattazione celle
            for row_num in range(1, len(dataframe) + 1):

                is_total = (
                    str(
                        dataframe.iloc[row_num - 1].get("at", "")
                    ).strip().lower()
                    == "totale complessivo"
                )

                for col_num, name in enumerate(dataframe.columns):

                    value = dataframe.iloc[row_num - 1, col_num]

                    # Gestione valori nulli e infiniti
                    if pd.isna(value):
                        value = None

                    elif isinstance(value, (int, float)) and not math.isfinite(float(value)):
                        value = None

                    if name == "Chiusura %":
                        fmt = total_percent_format if is_total else percent_format

                    elif is_total:
                        fmt = total_format

                    elif name in ["Risorse", "TT Trial", "TT chiusi"]:
                        fmt = number_format

                    else:
                        fmt = cell_format

                    worksheet.write(row_num, col_num, value, fmt)

            worksheet.freeze_panes(1, 0)

            worksheet.autofilter(
                0,
                0,
                len(dataframe),
                len(dataframe.columns) - 1
            )

        output.seek(0)

        return output

    # ------------------------------------------------------------
    # ELABORAZIONE
    # ------------------------------------------------------------
    if file_trial is not None and file_chiusi is not None:

        try:

            # ----------------------------------------------------
            # 1. LETTURA FILE
            # ----------------------------------------------------
            df_trial = pd.read_excel(file_trial, dtype=str)
            df_chiusi = pd.read_excel(file_chiusi, dtype=str)

            df_trial.columns = df_trial.columns.astype(str).str.strip()
            df_chiusi.columns = df_chiusi.columns.astype(str).str.strip()

            # ----------------------------------------------------
            # 2. CONTROLLO COLONNE OBBLIGATORIE
            # ----------------------------------------------------
            colonne_trial = ["ID WR", "at", "Impresa"]
            colonne_chiusi = ["ID WR"]

            mancanti_trial = [
                c for c in colonne_trial
                if c not in df_trial.columns
            ]

            mancanti_chiusi = [
                c for c in colonne_chiusi
                if c not in df_chiusi.columns
            ]

            if mancanti_trial:
                st.error(
                    "Nel dettaglio Trial mancano queste colonne: "
                    + ", ".join(mancanti_trial)
                )
                st.stop()

            if mancanti_chiusi:
                st.error(
                    "Nel file TT chiusi manca la colonna: "
                    + ", ".join(mancanti_chiusi)
                )
                st.stop()

            # ----------------------------------------------------
            # 3. PREPARAZIONE DETTAGLIO TRIAL
            # Il dettaglio Trial è la base:
            # AT e Impresa provengono da qui.
            # ----------------------------------------------------
            df_trial["ID WR"] = normalizza_id(df_trial["ID WR"])

            df_trial = df_trial[
                df_trial["ID WR"] != ""
            ].copy()

            df_trial["at"] = (
                df_trial["at"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df_trial["Impresa"] = (
                df_trial["Impresa"]
                .fillna("sociale")
                .astype(str)
                .str.strip()
            )

            df_trial.loc[
                df_trial["Impresa"] == "",
                "Impresa"
            ] = "sociale"

            # Una sola riga per WR:
            # evita di duplicare il denominatore.
            duplicati_trial = int(
                df_trial.duplicated(
                    subset=["ID WR"],
                    keep=False
                ).sum()
            )

            df_base = df_trial.drop_duplicates(
                subset=["ID WR"],
                keep="first"
            ).copy()

            if (df_base["at"] == "").any():
                st.warning(
                    "Alcune WR Trial non hanno il campo AT valorizzato."
                )

            # ----------------------------------------------------
            # 4. PREPARAZIONE FILE TT CHIUSI
            # Non importiamo né mostriamo nominativi
            # o dati dei tecnici.
            # La chiave di collegamento è ID WR.
            # ----------------------------------------------------
            df_chiusi["ID WR"] = normalizza_id(
                df_chiusi["ID WR"]
            )

            df_chiusi = df_chiusi[
                df_chiusi["ID WR"] != ""
            ].copy()

            duplicati_chiusi = int(
                df_chiusi.duplicated(
                    subset=["ID WR"],
                    keep=False
                ).sum()
            )

            df_chiusi = df_chiusi.drop_duplicates(
                subset=["ID WR"],
                keep="first"
            )

            # ----------------------------------------------------
            # 5. MERGE TT CHIUSI CON IL DETTAGLIO TRIAL
            # Inner join: contiamo solo chiusure riferite
            # a WR presenti nel Trial.
            # AT e Impresa sono ereditate dal dettaglio Trial.
            # ----------------------------------------------------
            df_match = df_chiusi[["ID WR"]].merge(
                df_base[["ID WR", "at", "Impresa"]],
                on="ID WR",
                how="inner"
            )

            wr_chiuse_fuori_trial = int(
                df_chiusi.loc[
                    ~df_chiusi["ID WR"].isin(df_base["ID WR"]),
                    "ID WR"
                ].nunique()
            )

            # ----------------------------------------------------
            # 6. RIEPILOGO PER AT E IMPRESA
            # ----------------------------------------------------
            trial_group = (
                df_base.groupby(
                    ["at", "Impresa"],
                    dropna=False
                )
                .agg(**{
                    "TT Trial": ("ID WR", "nunique")
                })
                .reset_index()
            )

            chiusi_group = (
                df_match.groupby(
                    ["at", "Impresa"],
                    dropna=False
                )
                .agg(**{
                    "TT chiusi": ("ID WR", "nunique")
                })
                .reset_index()
            )

            riepilogo = trial_group.merge(
                chiusi_group,
                on=["at", "Impresa"],
                how="left"
            )

            riepilogo["TT chiusi"] = (
                riepilogo["TT chiusi"]
                .fillna(0)
                .astype(int)
            )

            riepilogo["TT Trial"] = (
                riepilogo["TT Trial"].astype(int)
            )

            # Percentuale espressa su scala 0-100 (es. 25.0, non 0.25).
            riepilogo["Chiusura %"] = (
                riepilogo["TT chiusi"] /
                riepilogo["TT Trial"] * 100
            ).fillna(0)

            # ----------------------------------------------------
            # CALCOLO RISORSE
            # Se nel dettaglio Trial è presente la colonna Risorse,
            # la sommiamo per AT/Impresa.
            # ----------------------------------------------------
            if "Risorse" in df_base.columns:

                df_base["Risorse_num"] = pd.to_numeric(
                    df_base["Risorse"],
                    errors="coerce"
                ).fillna(0)

                risorse_group = (
                    df_base.groupby(
                        ["at", "Impresa"],
                        dropna=False
                    )["Risorse_num"]
                    .sum()
                    .reset_index(name="Risorse")
                )

                riepilogo = riepilogo.merge(
                    risorse_group,
                    on=["at", "Impresa"],
                    how="left"
                )

            # ----------------------------------------------------
            # ORDINAMENTO E COLONNE FINALI
            # ----------------------------------------------------
            riepilogo = riepilogo.sort_values(
                ["at", "Impresa"],
                kind="stable"
            ).reset_index(drop=True)

            colonne_finali = ["at", "Impresa"]

            if "Risorse" in riepilogo.columns:
                colonne_finali.append("Risorse")

            colonne_finali += [
                "TT Trial",
                "TT chiusi",
                "Chiusura %"
            ]

            riepilogo = riepilogo[colonne_finali]

            # ----------------------------------------------------
            # TOTALE COMPLESSIVO
            # La percentuale è calcolata sui totali,
            # non come media delle percentuali delle imprese.
            # ----------------------------------------------------
            totale_trial = int(
                df_base["ID WR"].nunique()
            )

            totale_chiusi = int(
                df_match["ID WR"].nunique()
            )

            riga_totale = {
                "at": "Totale complessivo",
                "Impresa": ""
            }

            if "Risorse" in riepilogo.columns:
                riga_totale["Risorse"] = int(
                    riepilogo["Risorse"].sum()
                )

            riga_totale["TT Trial"] = totale_trial
            riga_totale["TT chiusi"] = totale_chiusi

            riga_totale["Chiusura %"] = (
                totale_chiusi / totale_trial * 100
                if totale_trial
                else 0
            )

            riepilogo = pd.concat(
                [
                    riepilogo,
                    pd.DataFrame([riga_totale])
                ],
                ignore_index=True
            )

            # ----------------------------------------------------
            # 7. RISULTATI E CONTROLLI
            # ----------------------------------------------------
            st.subheader("Riepilogo chiusura Trial")

            k1, k2, k3 = st.columns(3)

            k1.metric(
                "TT Trial",
                f"{totale_trial:,}".replace(",", ".")
            )

            k2.metric(
                "TT chiusi",
                f"{totale_chiusi:,}".replace(",", ".")
            )

            percentuale_totale = (
                totale_chiusi / totale_trial * 100
                if totale_trial
                else 0
            )

            k3.metric(
                "Chiusura complessiva",
                f"{percentuale_totale:.1f}%"
            )

            if duplicati_trial:
                st.warning(
                    f"Nel dettaglio Trial sono presenti "
                    f"{duplicati_trial} righe appartenenti "
                    "a WR duplicate. Per il calcolo è stata "
                    "mantenuta una sola riga per ID WR."
                )

            if duplicati_chiusi:
                st.warning(
                    f"Nel file TT chiusi sono presenti "
                    f"{duplicati_chiusi} righe appartenenti "
                    "a WR duplicate. Ogni ID WR è stata "
                    "conteggiata una sola volta."
                )

            if wr_chiuse_fuori_trial:
                st.info(
                    f"Sono state escluse {wr_chiuse_fuori_trial} "
                    "WR chiuse che non risultano nel dettaglio "
                    "Trial caricato."
                )

            # ----------------------------------------------------
            # TABELLA PERCENTUALE DI CHIUSURA
            # ----------------------------------------------------
            st.subheader(
                "Tabella percentuale di chiusura per AT e Impresa"
            )

            # Mostra AT una sola volta per gruppo, come nella
            # tabella della Programmazione Trial.
            riepilogo_visualizzato = riepilogo.copy()
            riepilogo_visualizzato["at"] = (
                riepilogo_visualizzato["at"].fillna("").astype(str).str.strip()
            )
            maschera_totale = (
                riepilogo_visualizzato["at"].str.casefold()
                == "totale complessivo"
            )
            maschera_dati = ~maschera_totale
            riepilogo_visualizzato.loc[
                maschera_dati
                & riepilogo_visualizzato.loc[maschera_dati, "at"].duplicated(),
                "at"
            ] = ""

            formato_colonne = {
                "at": st.column_config.TextColumn(
                    "AT",
                    width=135
                ),

                "Impresa": st.column_config.TextColumn(
                    "Impresa",
                    width=260
                ),

                "Risorse": st.column_config.NumberColumn(
                    "Risorse",
                    width=95
                ),

                "TT Trial": st.column_config.NumberColumn(
                    "TT Trial",
                    width=100
                ),

                "TT chiusi": st.column_config.NumberColumn(
                    "TT chiusi",
                    width=105
                ),

                "Chiusura %": st.column_config.NumberColumn(
                    "Chiusura %",
                    format="%.1f%%",
                    width=115
                ),
            }

            st.dataframe(
                riepilogo_visualizzato,
                use_container_width=False,
                hide_index=True,
                column_config=formato_colonne
            )

            # ----------------------------------------------------
            # 8. GRAFICO: PERCENTUALE DI CHIUSURA PER IMPRESA
            # ----------------------------------------------------
            # Esclude sempre la riga totale dal grafico.
            grafico = riepilogo[
                riepilogo["at"].astype(str).str.strip().str.casefold()
                != "totale complessivo"
            ].copy()

            if not grafico.empty:

                # Percentuali memorizzate sulla scala 0-100.
                fig = px.bar(
                    grafico,
                    x="Impresa",
                    y="Chiusura %",
                    color="at",
                    barmode="group",
                    text=grafico["Chiusura %"].map(
                        lambda x: f"{x:.1f}%"
                    ),
                    title=(
                        "Percentuale di chiusura Trial "
                        "per Impresa e AT"
                    ),
                    labels={
                        "Chiusura %": "Chiusura",
                        "Impresa": "Impresa",
                        "at": "AT"
                    }
                )

                # Grafico più alto e testi delle percentuali leggibili.
                fig.update_traces(
                    textposition="outside",
                    textfont=dict(size=14),
                    cliponaxis=False
                )

                massimo = max(
                    100,
                    float(grafico["Chiusura %"].max()) * 1.18
                )

                fig.update_layout(
                    height=680,
                    margin=dict(l=55, r=35, t=85, b=150),
                    yaxis=dict(
                        title="Chiusura %",
                        ticksuffix="%",
                        range=[0, massimo],
                        dtick=20
                    ),
                    xaxis=dict(
                        title="Impresa",
                        tickangle=-35,
                        automargin=True
                    ),
                    legend_title_text="AT",
                    uniformtext_minsize=11,
                    uniformtext_mode="show"
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

            # 9. DOWNLOAD RIEPILOGO EXCEL
            # ----------------------------------------------------
            output_excel = excel_bytes(
                riepilogo,
                "CHIUSURA TRIAL"
            )

            st.download_button(
                label="📥 Scarica riepilogo chiusura Trial Excel",
                data=output_excel,
                file_name="riepilogo_chiusura_trial.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )



        except Exception as e:
            st.error(f"Errore durante l'elaborazione dei file: {e}")

    else:
        st.info(
            "Carica il dettaglio Excel Trial e il file Excel TT chiusi "
            "per calcolare la percentuale."
        )
