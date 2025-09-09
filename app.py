# ==============================
# Relação dinâmica entre duas colunas (auto por tipo)
# ==============================
import altair as alt
from pandas.api.types import is_numeric_dtype

st.subheader("📊 Relação entre duas colunas")
col_a = st.selectbox("👉 Coluna A (eixo X)", cols, key="col_a")
col_b = st.selectbox("👉 Coluna B (cor/medida)", cols, key="col_b")

relacao, diag = None, None  # manter compatível com a exportação

def _safe_str(s):
    return s.astype(str).fillna("—")

if col_a and col_b:
    A_num = is_numeric_dtype(df[col_a])
    B_num = is_numeric_dtype(df[col_b])

    base = df[[col_a, col_b]].copy()

    # ===== CASO 1: A categórica, B categórica → contagem empilhada =====
    if not A_num and not B_num:
        base[col_a] = _safe_str(base[col_a])
        base[col_b] = _safe_str(base[col_b])

        # Top N opcional pra B (resto vira "Outros")
        topn = st.slider(f"Top N categorias de {col_b}", 3, 20, 8, 1)
        tops = base[col_b].value_counts().nlargest(topn).index
        base.loc[~base[col_b].isin(tops), col_b] = "Outros"

        relacao = (
            base.groupby([col_a, col_b], dropna=False)
                .size()
                .reset_index(name="Medida")  # sempre "Medida" pra export/tooltip
        )
        total_por_a = relacao.groupby(col_a)["Medida"].transform("sum")
        relacao["Percentual"] = (relacao["Medida"] / total_por_a) * 100

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
                      alt.Tooltip("Medida:Q",   title="Quantidade"),
                      alt.Tooltip("Percentual:Q", title="Percentual", format=".1f")
                  ],
              )
              .properties(height=340)
              .configure_axis(grid=True)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

        # Donut de B
        st.subheader(f"🥧 Distribuição de {col_b}")
        dist = (base[col_b].value_counts().reset_index()
                .rename(columns={"index": col_b, col_b: "Medida"}))
        total = dist["Medida"].sum()
        dist["Percentual"] = dist["Medida"] / total * 100
        donut = (
            alt.Chart(dist)
              .mark_arc(outerRadius=120, innerRadius=60)
              .encode(
                  theta=alt.Theta("Medida:Q"),
                  color=alt.Color(f"{col_b}:N", title=col_b, scale=alt.Scale(scheme="tableau10")),
                  tooltip=[
                      alt.Tooltip(f"{col_b}:N", title=col_b),
                      alt.Tooltip("Medida:Q", title="Quantidade"),
                      alt.Tooltip("Percentual:Q", title="Percentual", format=".1f")
                  ]
              )
              .properties(height=340)
              .configure_view(stroke=None)
        )
        st.altair_chart(donut, use_container_width=True)

        maior = relacao.loc[relacao["Medida"].idxmax()]
        diag = (f"⚠️ Diagnóstico:\n\n"
                f"- **{maior[col_a]} × {maior[col_b]}** teve **{maior['Medida']} ocorrências** "
                f"({maior['Percentual']:.1f}%).")

    # ===== CASO 2: A categórica, B numérica → barra com agregação (Soma/Média/Mediana/Contagem) =====
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
                      alt.Tooltip("Medida:Q",   title=medida_titulo)
                  ]
              )
              .properties(height=340)
              .configure_axis(grid=True)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

    # ===== CASO 3: A numérica, B categórica → boxplot por categoria =====
    elif A_num and not B_num:
        base[col_b] = _safe_str(base[col_b])
        chart = (
            alt.Chart(base)
              .mark_boxplot()
              .encode(
                  x=alt.X(f"{col_b}:N", title=col_b),
                  y=alt.Y(f"{col_a}:Q",  title=col_a),
                  color=alt.Color(f"{col_b}:N", legend=None)
              )
              .properties(height=340)
              .configure_view(stroke=None)
        )
        st.altair_chart(chart, use_container_width=True)

    # ===== CASO 4: A numérica, B numérica → dispersão com linha de tendência =====
    else:
        show_trend = st.toggle("Mostrar linha de tendência", value=True)
        points = (
            alt.Chart(base)
              .mark_circle(opacity=0.6, size=80)
              .encode(
                  x=alt.X(f"{col_a}:Q", title=col_a),
                  y=alt.Y(f"{col_b}:Q", title=col_b),
                  tooltip=[alt.Tooltip(f"{col_a}:Q", title=col_a),
                           alt.Tooltip(f"{col_b}:Q", title=col_b)]
              )
        )
        if show_trend:
            trend = points.transform_regression(col_a, col_b).mark_line()
            chart = points + trend
        else:
            chart = points
        st.altair_chart(chart.properties(height=360).configure_view(stroke=None),
                        use_container_width=True)
