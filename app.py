import streamlit as st
import pandas as pd

# NOVO: Altair pra barras sem "value/color"
import altair as alt

# Import matplotlib com fallback (pizza continua igual)
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

from difflib import get_close_matches

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

# ------------------ IA: Manual ------------------
st.sidebar.markdown("### 📘 Manual (opcional)")
kb_file = st.sidebar.file_uploader("Suba o manual (CSV ou XLSX) com colunas: termo, conclusao, solucoes",
                                   type=["csv", "xlsx"], key="kb")

def _norm(s: str) -> str:
    s = str(s).lower()
    keep = []
    for ch in s:
        if ch.isalnum() or ch.isspace() or ch in "-_/":
            keep.append(ch)
    return "".join(keep).strip()

@st.cache_data
def load_kb(file):
    if not file:
        return pd.DataFrame()
    if file.name.lower().endswith(".xlsx"):
        df = pd.read_excel(file)
    else:
        try:
            df = pd.read_csv(file, sep=None, engine="python")
        except Exception:
            df = pd.read_csv(file)
    # normaliza nomes de colunas esperados
    cols = {c.strip().lower(): c for c in df.columns}
    rename = {}
    if "termo" in cols: rename[cols["termo"]] = "termo"
    if "conclusao" in cols: rename[cols["conclusao"]] = "conclusao"
    if "solucoes" in cols: rename[cols["solucoes"]] = "solucoes"
    df = df.rename(columns=rename)
    for need in ["termo", "conclusao", "solucoes"]:
        if need not in df.columns:
            df[need] = ""
    df["termo_norm"] = df["termo"].map(_norm)
    return df[["termo", "termo_norm", "conclusao", "solucoes"]]

# Base embutida (fallback) — edite à vontade depois
KB_DEFAULT = pd.DataFrame([
    {"termo": "LOW_PRESSURE", "conclusao": "Pressão baixa no circuito de tinta/jet.",
     "solucoes": "Verificar nível de solvente;Checar mangueiras e engates;Executar rotina de pressurização;Inspecionar vazamentos;Reiniciar bomba/recirculação"},
    {"termo": "ALTA VISCOSIDADE", "conclusao": "Viscosidade de tinta acima da janela.",
     "solucoes": "Checar make-up/solvente;Executar rotina de diluição;Verificar sensor de viscosidade;Avaliar temperatura ambiente;Padronizar tampas e reabastecimento"},
    {"termo": "NOZZLE CLOG / ENTUPIMENTO DE BICO", "conclusao": "Bico/jet possivelmente obstruído.",
     "solucoes": "Limpar cabeça de impressão;Aplicar flush conforme manual;Verificar filtro;Checar qualidade da tinta;Agendar limpeza preventiva"},
    {"termo": "MISALIGNMENT / DESALINHADO", "conclusao": "Cabeça desalinhada em relação ao produto.",
     "solucoes": "Ajustar distância/ângulo da cabeça;Fixar suportes;Testar impressão e leitura;Revisar gabarito/guia do produto"},
    {"termo": "IMPRESSÃO CLARA / FADED", "conclusao": "Marcação com baixa densidade/contraste.",
     "solucoes": "Checar velocidade vs. setpoint;Ajustar tensão/tempo de gota (conforme manual);Verificar tinta/solvente;Limpar bico e eletrodos;Revisar distância cabeça-produto"},
])

KB_DEFAULT["termo_norm"] = KB_DEFAULT["termo"].map(_norm)

kb_user = load_kb(kb_file)
KB_ALL = kb_user if not kb_user.empty else KB_DEFAULT

def kb_lookup(term: str, cutoff=0.78):
    """Busca termo no manual (fuzzy). Retorna dict {conclusao, solucoes[]} ou None."""
    if KB_ALL.empty:
        return None
    term_norm = _norm(term)
    universe = KB_ALL["termo_norm"].tolist()
    match = get_close_matches(term_norm, universe, n=1, cutoff=cutoff)
    if not match:
        return None
    row = KB_ALL.loc[KB_ALL["termo_norm"] == match[0]].iloc[0]
    sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
    return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": ("manual" if not kb_user.empty else "embutido")}

# ------------------------------------------------

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

        # Qual coluna consultar no manual? (padrão: a segunda — ex.: DEFEITO/CAUSA)
        coluna_manual = st.radio("Consultar manual usando qual coluna?",
                                 [col_b, col_a], index=0, horizontal=True)

        relacao, diag = None, None
        try:
            if col_a and col_b:
                # contagem (robusta a NaN)
                tmp = df[[col_a, col_b]].copy()
                tmp[col_a] = tmp[col_a].astype(str).fillna("—")
                tmp[col_b] = tmp[col_b].astype(str).fillna("—")

                relacao = (
                    tmp.groupby([col_a, col_b], dropna=False)
                       .size()
                       .reset_index(name="QTD")
                )

                if not relacao.empty:
                    # ---------- BARRAS (Altair) ----------
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

                    # ---------- PIZZA (igual ao seu) ----------
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
                        ax.legend(
                            wedges,
                            dist.index,
                            title=f"{col_b}",
                            loc="center left",
                            bbox_to_anchor=(1, 0, 0.5, 1)
                        )
                        ax.set_title(f"Distribuição de {col_b}", fontsize=14)
                        st.pyplot(fig)

                    # ---------- DIAGNÓSTICO ORIGINAL ----------
                    top = relacao.sort_values("QTD", ascending=False).iloc[0]
                    diag = (
                        f"⚠️ Diagnóstico Preventivo:\n\n"
                        f"- A combinação **{top[col_a]} x {top[col_b]}** apresentou **{int(top['QTD'])} ocorrências**.\n"
                        f"- Recomenda-se intensificar manutenção preventiva em **{top[col_a]}**, "
                        f"com foco em evitar novos casos de **{top[col_b]}**."
                    )
                    st.success(diag)

                    # ---------- 🧠 IA (Manual) ----------
                    termo_para_consulta = str(top[coluna_manual])
                    kb_res = kb_lookup(termo_para_consulta)

                    with st.expander("🧠 Conclusão automática (Manual)", expanded=True):
                        if kb_res:
                            st.markdown(f"**Conclusão:** {kb_res['conclusao']}")
                            if kb_res["solucoes"]:
                                st.markdown("**Possíveis soluções:**")
                                st.markdown("\n".join([f"- {s}" for s in kb_res["solucoes"]]))
                            st.caption(f"Fonte: {kb_res['fonte']}")
                        else:
                            st.info("Não encontrei esse item no manual. Suba um CSV/XLSX com colunas **termo, conclusao, solucoes** para recomendações específicas.")
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
