import streamlit as st
import pandas as pd

# NOVO: Altair pra gráfico de barras sem "value/color"
import altair as alt

# Import matplotlib com fallback (pizza continua igual)
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
                    # ---------- TROCA: sai st.bar_chart, entra Altair bonitão ----------
                    chart = (
                        alt.Chart(relacao)
                           .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                           .encode(
                               x=alt.X(f"{col_a}:N", title=col_a),
                               y=alt.Y("QTD:Q", title="Quantidade"),
                               color=alt.Color(f"{col_b}:N", title=col_b,
                                               legend=alt.Legend(orient="bottom")),
                               tooltip=[
                                   alt.Tooltip(f"{col_a}:N", title=col_a),
                                   alt.Tooltip(f"{col_b}:N", title=col_b),
                                   alt.Tooltip("QTD:Q", title="Quantidade"),
                               ],
                           )
                           .properties(height=340)
                           .configure_axis(grid=True)
                           .configure_view(stroke=None)
                    )
                    st.altair_chart(chart, use_container_width=True)
                    # ---------------------------------------------------------------

                    # 🍕 Pizza (mantida exatamente como estava)
                    if MATPLOTLIB_OK:
                        st.subheader(f"🥧 Distribuição de {col_b}")
                        dist = df[col_b].value_counts(normalize=True) * 100  # porcentagem
                        dist = dist.sort_values(ascending=False)

                        # Agrupar categorias pequenas (<5%)
                        outros = dist[dist < 5].sum()
                        dist = dist[dist >= 5]
                        if outros > 0:
                            dist["Outros"] = outros

                        fig, ax = plt.subplots(figsize=(7, 6))
                        wedges, texts, autotexts = ax.pie(
                            dist,
                            autopct="%1.1f%%",
                            startangle=90,
                            counterclock=False,
                            colors=plt.cm.tab20.colors
                        )
                        # Legenda fora do gráfico
                        ax.legend(
                            wedges,
                            dist.index,
                            title=f"{col_b}",
                            loc="center left",
                            bbox_to_anchor=(1, 0, 0.5, 1)
                        )
                        ax.set_title(f"Distribuição de {col_b}", fontsize=14)
                        st.pyplot(fig)

                    # ✅ Diagnóstico final (mantido do jeito que estava)
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

                    if relacao is not None:
                        relacao.to_excel(writer, sheet_name="Relação", index=False)
                        ws = writer.sheets["Relação"]
                        chart = wb.add_chart({"type": "column"})
                        chart.add_series({
                            "categories": ["Relação", 1, 0, len(relacao), 0],
                            "values":     ["Relação", 1, 2, len(relacao), 2],
                            "name":       f"{col_a} x {col_b}"
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
