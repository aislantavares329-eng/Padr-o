# app.py — Analisador Dinâmico de Planilhas (flex por tipo de coluna)
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pandas.api.types import is_numeric_dtype

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

# ---------- leitura segura ----------
def read_any(file):
    if file.name.lower().endswith(".xlsx"):
        xls = pd.ExcelFile(file)
        aba = st.selectbox("📑 Escolha a aba", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=aba)
    else:
        try:
            df = pd.read_csv(file, sep=None, engine="python")
        except Exception:
            df = pd.read_csv(file)
    df.columns = [str(c).strip() for c in df.columns]
    return df

uploaded_file = st.file_uploader("📂 Suba sua planilha (.xlsx ou .csv)", type=["xlsx", "csv"])
if not uploaded_file:
    st.stop()

# ---------- base e preview ----------
df = read_any(uploaded_file)
st.subheader("🔎 Pré-visualização")
st.dataframe(df.head(), use_container_width=True)

cols = df.columns.tolist()
st.write("📋 Colunas detectadas:", cols)

# ==============================
# Relação dinâmica entre duas colunas (auto por tipo)
# ==============================
st.subheader("📊 Relação entre duas colunas")

col_a = st.selectbox("👉 Coluna A (eixo X)", cols, key="col_a")
col_b = st.selectbox("👉 Coluna B (cor/medida)", cols, key="col_b")

relacao = None   # para export
diag = None

def _safe_str(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("—")

if col_a and col_b:
    A_num = is_numeric_dtype(df[col_a])
    B_num = is_numeric_dtype(df[col_b])

    base = df[[col_a, col_b]].copy()

    # ===== Caso 1: A categórica, B categórica → contagem empilhada =====
    if not A_num and not B_num:
        base[col_a] = _safe_str(base[col_a])
        base[col_b] = _safe_str(base[col_b])

        # Top N pra B (resto -> "Outros")
        topn = st.slider(f"Top N categorias de {col_b}", 3, 20, 8, 1)
        tops = base[col_b].value_counts().nlargest(topn).index
        base.loc[~base[col_b].isin(tops), col_b] = "Outros"

        # Tabela de relação
        relacao = (
            base.groupby([col_a, col_b], dropna=False)
                .size()
                .reset_index(name="Medida")  # nome neutro p/ export/tooltip
        )
        total_por_a = relacao.groupby(col_a)["Medida"].transform("sum")
        relacao["Percentual"] = (relacao["Medida"] / total_por_a.replace({0: np.nan})) * 100

        # Gráfico barras empilhadas vs 100%
        modo = st.radio("Modo do gráfico", ["Quantidade", "Percentual (100%)"],
                        horizontal=True, key="modo_catcat")
        y_enc = (alt.Y("sum(Medida):Q", title="Quantidade")
                 if modo == "Quantidade"
                 else alt.Y("sum(Medida):Q", title="Percentual", stack="normalize"))

        chart = (
            alt.Chart(relacao)
              .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
              .encode(
                  x=alt.X(f"{col_a}:N", title=col_a, sort='-y'),
                  y=y_enc,
                  color=alt.Color(f"{col_b}:N", title=col_b,
                                  legend=alt.Legend(orient="bottom"),
                                  scale=alt.Scale(scheme="tableau10")),
                  order=alt.Order("Medida:Q", sort="descending"),
                  tooltip=[
                      alt.Tooltip(f"{col_a}:N", title=col_a),
                      alt.Tooltip(f"{col_b}:N", title=col_b),
                      alt.Tooltip("Medida:Q", title="Quantidade"),
                      alt.Tooltip("Percentual:Q", title="Percentual", format=".1f"),
                  ],
              )
              .properties(height=340)
              .configure_axis(grid=True)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

        # Donut robusto da distribuição de B
        st.subheader(f"🥧 Distribuição de {col_b}")
        dist = (
            base[col_b]
              .astype(str)
              .value_counts(dropna=False)
              .reset_index(name="Medida")
              .rename(columns={"index": col_b})
        )
        # garante numérico e evita crash de tipo
        dist["Medida"] = pd.to_numeric(dist["Medida"], errors="coerce").fillna(0).astype(float)
        total = float(dist["Medida"].sum())
        dist["Percentual"] = 0.0 if total == 0 else (dist["Medida"] / total) * 100.0
        dist = dist.sort_values("Medida", ascending=False)

        donut = (
            alt.Chart(dist)
              .mark_arc(outerRadius=120, innerRadius=60)
              .encode(
                  theta=alt.Theta("Medida:Q"),
                  color=alt.Color(f"{col_b}:N", title=col_b, scale=alt.Scale(scheme="tableau10")),
                  tooltip=[
                      alt.Tooltip(f"{col_b}:N", title=col_b),
                      alt.Tooltip("Medida:Q", title="Quantidade"),
                      alt.Tooltip("Percentual:Q", title="Percentual", format=".1f"),
                  ],
              )
              .properties(height=340)
              .configure_view(stroke=None)
        )
        st.altair_chart(donut, use_container_width=True)

        maior = relacao.loc[relacao["Medida"].idxmax()]
        diag = (f"⚠️ Diagnóstico:\n\n"
                f"- **{maior[col_a]} × {maior[col_b]}** teve **{int(maior['Medida'])} ocorrências** "
                f"({maior['Percentual']:.1f}%).")

    # ===== Caso 2: A categórica, B numérica → barra agregada =====
    elif not A_num and B_num:
        agg = st.radio(f"Agregação de {col_b}", ["Soma", "Média", "Mediana", "Contagem"], horizontal=True)
        agg_map = {"Soma": "sum", "Média": "mean", "Mediana": "median", "Contagem": "size"}
        fn = agg_map[agg]

        relacao = (
            base.groupby(col_a, dropna=False)
                .agg(Medida=(col_b, fn))
                .reset_index()
        )
        medida_titulo = f"{agg} de {col_b}" if agg != "Contagem" else "Contagem"

        chart = (
            alt.Chart(relacao)
              .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
              .encode(
                  x=alt.X(f"{col_a}:N", title=col_a, sort='-y'),
                  y=alt.Y("Medida:Q", title=medida_titulo),
                  tooltip=[
                      alt.Tooltip(f"{col_a}:N", title=col_a),
                      alt.Tooltip("Medida:Q",   title=medida_titulo),
                  ],
              )
              .properties(height=340)
              .configure_axis(grid=True)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

    # ===== Caso 3: A numérica, B categórica → boxplot =====
    elif A_num and not B_num:
        base[col_b] = _safe_str(base[col_b])
        chart = (
            alt.Chart(base)
              .mark_boxplot()
              .encode(
                  x=alt.X(f"{col_b}:N", title=col_b),
                  y=alt.Y(f"{col_a}:Q",  title=col_a),
                  color=alt.Color(f"{col_b}:N", legend=None),
              )
              .properties(height=340)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

    # ===== Caso 4: A numérica, B numérica → dispersão =====
    else:
        show_trend = st.toggle("Mostrar linha de tendência", value=True)
        points = (
            alt.Chart(base)
              .mark_circle(opacity=0.6, size=80)
              .encode(
                  x=alt.X(f"{col_a}:Q", title=col_a),
                  y=alt.Y(f"{col_b}:Q", title=col_b),
                  tooltip=[
                      alt.Tooltip(f"{col_a}:Q", title=col_a),
                      alt.Tooltip(f"{col_b}:Q", title=col_b),
                  ],
              )
        )
        chart = points + points.transform_regression(col_a, col_b).mark_line() if show_trend else points
        st.altair_chart(chart.properties(height=360).configure_view(stroke=None),
                        use_container_width=True)

# ==============================
# Exportar Excel (se houver "relacao")
# ==============================
if st.button("📥 Gerar Relatório Excel"):
    try:
        saida = "relatorio_dinamico.xlsx"
        with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
            wb = writer.book
            df.to_excel(writer, sheet_name="Base", index=False)

            if relacao is not None and not relacao.empty:
                relacao.to_excel(writer, sheet_name="Relação", index=False)
                ws = writer.sheets["Relação"]

                # Colunas: 0 = col_a, 1 = col_b (se existir), 2 = Medida
                chart = wb.add_chart({"type": "column"})
                chart.add_series({
                    "categories": ["Relação", 1, 0, len(relacao), 0],  # eixo X = col_a
                    "values":     ["Relação", 1, 2, len(relacao), 2],  # Medida
                    "name":       f"{col_a}" + (f" × {col_b}" if col_b in relacao.columns else ""),
                })
                chart.set_title({"name": "Relação"})
                chart.set_legend({"position": "bottom"})
                ws.insert_chart("E2", chart)

                if diag:
                    ws.write(len(relacao) + 3, 0, "Diagnóstico:")
                    ws.write(len(relacao) + 4, 0, diag)

        with open(saida, "rb") as f:
            st.download_button("⬇️ Baixar Relatório", f, file_name=saida)
    except Exception as e:
        st.error(f"❌ Erro ao gerar relatório Excel: {e}")
