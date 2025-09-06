import streamlit as st
import pandas as pd

# Import matplotlib com fallback
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

uploaded_file = st.file_uploader("📂 Suba sua planilha (.xlsx ou .csv)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # Detectar abas (se for Excel)
        if uploaded_file.name.endswith(".xlsx"):
            try:
                xls = pd.ExcelFile(uploaded_file)
                aba = st.selectbox("📑 Escolha a aba", xls.sheet_names)
                df = pd.read_excel(xls, sheet_name=aba)
            except Exception as e:
                st.error(f"❌ Erro ao carregar aba do Excel: {e}")
                st.stop()
        else:
            try:
                df = pd.read_csv(uploaded_file, sep=None, engine="python")
            except Exception as e:
                st.error(f"❌ Erro ao carregar CSV: {e}")
                st.stop()

        st.subheader("🔎 Pré-visualização")
        st.dataframe(df.head())

        cols = df.columns.tolist()
        st.write("📋 Colunas detectadas:", cols)

        # ==============================
        # Relação categórica
        # ==============================
        st.subheader("📊 Relação entre duas colunas categóricas")
        col_a = st.selectbox("👉 Primeira coluna categórica", cols, key="cata")
        col_b = st.selectbox("👉 Segunda coluna categórica", cols, key="catb")

        relacao, diag = None, None
        try:
            if col_a and col_b:
                relacao = df.groupby([col_a, col_b]).size().reset_index(name="QTD")

                if not relacao.empty:
                    pivot = relacao.pivot(index=col_a, columns=col_b, values="QTD").fillna(0)
                    st.bar_chart(pivot)

                    if MATPLOTLIB_OK:
                        st.subheader(f"🥧 Distribuição de {col_b}")
                        dist = df[col_b].value_counts()
                        st.pyplot(dist.plot.pie(autopct="%1.1f%%", figsize=(5, 5)).get_figure())

                    maior = relacao.loc[relacao["QTD"].idxmax()]
                    diag = (
                        f"⚠️ Diagnóstico Preventivo:\n\n"
                        f"- A combinação **{maior[col_a]} x {maior[col_b]}** apresentou **{maior['QTD']} ocorrências**.\n"
                        f"- Recomenda-se intensificar manutenção preventiva em **{maior[col_a]}**, "
                        f"com foco em evitar novos casos de **{maior[col_b]}**."
                    )
                    st.success(diag)
                else:
                    st.warning("⚠️ Não há dados suficientes para gerar a relação.")
        except Exception as e:
            st.error(f"❌ Erro ao gerar gráficos categóricos: {e}")

        # ==============================
        # Exportar Excel
        # ==============================
        if st.button("📥 Gerar Relatório Excel"):
            try:
                saida = "relatorio_dinamico.xlsx"
                with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
                    wb = writer.book
                    df.to_excel(writer, sheet_name="Base", index=False)

                    if df_corr is not None and corr_val is not None:
                        df_corr.to_excel(writer, sheet_name="Correlação", index=False)
                        ws = writer.sheets["Correlação"]
                        ws.write(len(df_corr) + 2, 0, "Coeficiente de Correlação (Pearson):")
                        ws.write(len(df_corr) + 2, 1, corr_val)
                        if insight:
                            ws.write(len(df_corr) + 3, 0, "Insight:")
                            ws.write(len(df_corr) + 3, 1, insight)

                        chart = wb.add_chart({"type": "scatter"})
                        chart.add_series({
                            "categories": ["Correlação", 1, 0, len(df_corr), 0],
                            "values": ["Correlação", 1, 1, len(df_corr), 1],
                            "name": f"{col_x} vs {col_y}"
                        })
                        chart.set_title({"name": f"{col_x} x {col_y}"})
                        ws.insert_chart("E2", chart)

                    if relacao is not None:
                        relacao.to_excel(writer, sheet_name="Relação", index=False)
                        ws = writer.sheets["Relação"]
                        chart = wb.add_chart({"type": "column"})
                        chart.add_series({
                            "categories": ["Relação", 1, 0, len(relacao), 0],
                            "values": ["Relação", 1, 2, len(relacao), 2],
                            "name": f"{col_a} x {col_b}"
                        })
                        chart.set_title({"name": f"{col_a} x {col_b}"})
                        ws.insert_chart("E2", chart)
                        if diag:
                            ws.write(len(relacao) + 3, 0, "Diagnóstico Preventivo:")
                            ws.write(len(relacao) + 4, 0, diag)

                with open(saida, "rb") as f:
                    st.download_button("⬇️ Baixar Relatório", f, file_name=saida)
            except Exception as e:
                st.error(f"❌ Erro ao gerar relatório Excel: {e}")

    except Exception as e:
        st.error(f"❌ Erro geral: {e}")
