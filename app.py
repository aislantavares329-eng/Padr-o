# app.py — Analisador Dinâmico de Planilhas
# Barras com Altair, pizza original, diagnóstico original
# IA do manual com PDF (pypdf/PyPDF2) + CSV fallback, alias + fuzzy + jaccard

import streamlit as st
import pandas as pd
import altair as alt
import numpy as np
import unicodedata
from difflib import get_close_matches

# === matplotlib só para a pizza (mantém layout antigo) ===
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

# === PDF readers (tentativas) ===
PDF_BACKEND = None
try:
    import pypdf  # moderno
    PDF_BACKEND = "pypdf"
except Exception:
    try:
        import PyPDF2  # alternativo
        PDF_BACKEND = "PyPDF2"
    except Exception:
        PDF_BACKEND = None

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

# ===========================
# Helpers de normalização
# ===========================
def _norm(s: str) -> str:
    """normaliza texto (sem acento, minúsculo, só alfa-num/esp/-_/)."""
    s = unicodedata.normalize("NFD", str(s).lower())
    s = "".join(ch for ch in s if ch.isalnum() or ch.isspace() or ch in "-_/")
    s = " ".join(s.split())
    return s

def _tokens(s: str):
    return set(_norm(s).split())

# ===========================
# Sidebar: Manual (CSV + PDF)
# ===========================
st.sidebar.markdown("### 📘 Manual (opcional)")
kb_file = st.sidebar.file_uploader(
    "Suba o manual em **CSV/XLSX** (colunas: termo, conclusao, solucoes)",
    type=["csv", "xlsx"],
    key="kb"
)
pdf_files = st.sidebar.file_uploader(
    "Ou suba **um ou mais PDFs** do manual",
    type=["pdf"],
    key="pdfs",
    accept_multiple_files=True
)
usar_pdf = st.sidebar.toggle("🔎 Usar manual em PDF com prioridade", value=True, help="Se ligado, a IA busca soluções direto do(s) PDF(s). Se não achar, cai no CSV/fallback.")

# ===========================
# Carregamento do CSV manual
# ===========================
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
    cols = {c.strip().lower(): c for c in df.columns}
    rename = {}
    for need in ["termo", "conclusao", "solucoes"]:
        if need in cols:
            rename[cols[need]] = need
    df = df.rename(columns=rename)
    for need in ["termo", "conclusao", "solucoes"]:
        if need not in df.columns:
            df[need] = ""
    df["termo_norm"] = df["termo"].map(_norm)
    df["tokens"] = df["termo"].map(_tokens)
    return df[["termo", "termo_norm", "tokens", "conclusao", "solucoes"]]

KB_DEFAULT = pd.DataFrame([
    {"termo": "LOW_PRESSURE", "conclusao": "Pressão baixa no circuito de tinta/jet.",
     "solucoes": "Verificar nível de solvente;Checar mangueiras e engates;Executar rotina de pressurização;Inspecionar vazamentos;Testar transdutor/bomba"},
    {"termo": "ALTA VISCOSIDADE", "conclusao": "Viscosidade acima da janela.",
     "solucoes": "Checar make-up/solvente;Executar rotina de diluição;Verificar sensor de viscosidade;Ajustar temperatura ambiente"},
    {"termo": "NOZZLE CLOG / ENTUPIMENTO DE BICO", "conclusao": "Bico/jato possivelmente obstruído.",
     "solucoes": "Limpar cabeça de impressão;Aplicar flush;Trocar/limpar filtro;Checar qualidade da tinta;Agendar preventiva"},
    {"termo": "MISALIGNMENT / DESALINHADO", "conclusao": "Cabeça desalinhada do produto.",
     "solucoes": "Ajustar distância/ângulo;Fixar suportes;Revisar gabarito/guia;Testar leitura"},
    {"termo": "IMPRESSÃO CLARA / FADED", "conclusao": "Baixa densidade/contraste de marcação.",
     "solucoes": "Ajustar velocidade/atraso;Checar tinta/solvente;Limpar bico/eletrodos;Revisar distância cabeça-produto"},
])
KB_DEFAULT["termo_norm"] = KB_DEFAULT["termo"].map(_norm)
KB_DEFAULT["tokens"] = KB_DEFAULT["termo"].map(_tokens)

kb_user = load_kb(kb_file)
KB_ALL = (pd.concat([kb_user, KB_DEFAULT], ignore_index=True)
          if not kb_user.empty else KB_DEFAULT)

ALIASES = {
    "falha de jato": "nozzle clog / entupimento de bico",
    "cabeçote sujo": "cabeçote requer limpeza ao desligar",
    "ausencia de impressao": "impressão clara / faded",
    "perda de modulacao": "impressão clara / faded",
}
ALIASES_NORM = { _norm(k): _norm(v) for k, v in ALIASES.items() }

# ===========================
# PDF → texto por página
# ===========================
@st.cache_data(show_spinner=False)
def read_pdfs(files):
    pages = []  # [{source, page, text}]
    if not files:
        return pages
    for f in files:
        try:
            if PDF_BACKEND == "pypdf":
                reader = pypdf.PdfReader(f)
                for i, p in enumerate(reader.pages, start=1):
                    txt = p.extract_text() or ""
                    pages.append({"source": f.name, "page": i, "text": txt})
            elif PDF_BACKEND == "PyPDF2":
                reader = PyPDF2.PdfReader(f)
                for i, p in enumerate(reader.pages, start=1):
                    txt = p.extract_text() or ""
                    pages.append({"source": f.name, "page": i, "text": txt})
        except Exception:
            continue
    return pages

PDF_PAGES = read_pdfs(pdf_files)

def _score_jaccard(a_tokens:set, b_tokens:set) -> float:
    if not a_tokens or not b_tokens:
        return 0.0
    return len(a_tokens & b_tokens) / len(a_tokens | b_tokens)

def _extract_steps(text:str, max_lines=12):
    """Puxa linhas com padrão de instrução ('-','•','·','*','1.','2.' ou verbos comuns)."""
    steps = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    verbs = ("verificar","checar","inspecionar","executar","limpar","ajustar",
             "substituir","alinhar","recolocar","testar","revisar","aguardar")
    for ln in lines:
        ln_l = _norm(ln)
        if ln.startswith(("-", "•", "·", "*")) or ln[:2].isdigit() or any(ln_l.startswith(v) for v in verbs):
            steps.append(ln.lstrip("-•·* ").strip())
        if len(steps) >= max_lines:
            break
    # remova duplicadas curtas
    uniq = []
    seen = set()
    for s in steps:
        key = _norm(s)
        if key and key not in seen:
            uniq.append(s)
            seen.add(key)
    return uniq

def kb_lookup_pdf(term:str):
    """Busca termo nos PDFs. Retorna dict com conclusao, solucoes, fonte (arquivo/página)."""
    if not PDF_PAGES:
        return None
    q_norm = _norm(term)
    q_tok  = _tokens(term)

    # 1) alias direto → re-busca
    q_norm = ALIASES_NORM.get(q_norm, q_norm)
    q_tok  = _tokens(q_norm)

    # rankeia páginas por jaccard + presença bruta do termo
    scored = []
    for p in PDF_PAGES:
        txt = p["text"]
        tnorm = _norm(txt)
        tokens = _tokens(tnorm)
        s = _score_jaccard(q_tok, tokens)
        if q_norm in tnorm:
            s += 0.25  # boost se match literal
        if s > 0:
            scored.append((s, p))
    if not scored:
        return None
    scored.sort(key=lambda x: x[0], reverse=True)
    best = scored[0][1]

    # conclusão = primeira linha próxima do termo (ou início da página)
    text = best["text"]
    tnorm = _norm(text)
    pos = tnorm.find(q_norm) if q_norm in tnorm else 0
    window = text[max(0, pos-300): pos+800]  # pega bloco ao redor
    # conclusão: a primeira sentença do bloco
    concl = window.strip().splitlines()[0].strip()
    if len(concl) < 8:  # muito curta? pega a próxima
        for ln in window.splitlines():
            if len(ln.strip()) > 8:
                concl = ln.strip()
                break

    steps = _extract_steps(window)
    return {
        "conclusao": concl.strip(),
        "solucoes": steps,
        "fonte": f"{best['source']} p.{best['page']}"
    }

def kb_lookup_csv(term: str, cutoff_close=0.65, cutoff_jacc=0.35):
    if KB_ALL.empty:
        return None
    term_norm = _norm(term)

    # alias
    target_norm = ALIASES_NORM.get(term_norm, term_norm)
    hit = KB_ALL[KB_ALL["termo_norm"] == target_norm]
    if not hit.empty:
        row = hit.iloc[0]
        sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
        return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": "manual CSV/alias"}

    # close-match
    cand = get_close_matches(term_norm, KB_ALL["termo_norm"].tolist(), n=1, cutoff=cutoff_close)
    if cand:
        row = KB_ALL.loc[KB_ALL["termo_norm"] == cand[0]].iloc[0]
        sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
        return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": "manual CSV/fuzzy"}

    # jaccard
    tA = _tokens(term_norm)
    if tA:
        jacc = KB_ALL["tokens"].apply(lambda tB: _score_jaccard(tA, tB))
        idx = int(jacc.idxmax())
        if jacc.iloc[idx] >= cutoff_jacc:
            row = KB_ALL.iloc[idx]
            sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
            return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": f"manual CSV/jaccard {jacc.iloc[idx]:.2f}"}
    return None

def kb_lookup(term: str):
    """Resolução final: PDF (se habilitado e disponível) → CSV → fallback embutido."""
    if usar_pdf and PDF_PAGES and PDF_BACKEND is not None:
        hit = kb_lookup_pdf(term)
        if hit and (hit["conclusao"] or hit["solucoes"]):
            return hit
    # Cai para CSV/fallback
    return kb_lookup_csv(term)

# ===========================
# Leitura da planilha
# ===========================
uploaded_file = st.file_uploader("📂 Suba sua planilha (.xlsx ou .csv)", type=["xlsx", "csv"])

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
    return df

if uploaded_file is not None:
    try:
        df = read_any(uploaded_file)

        st.subheader("🔎 Pré-visualização")
        st.dataframe(df.head(), use_container_width=True)

        cols = df.columns.tolist()
        st.write("📋 Colunas detectadas:", cols)

        # ==============================
        # Relação categórica (igual à sua)
        # ==============================
        st.subheader("📊 Relação entre duas colunas categóricas")
        col_a = st.selectbox("👉 Primeira coluna categórica", cols, key="cata")
        col_b = st.selectbox("👉 Segunda coluna categórica", cols, key="catb")

        coluna_manual = st.radio("Consultar manual usando qual coluna?",
                                 [col_b, col_a], index=0, horizontal=True)

        relacao, diag = None, None
        try:
            if col_a and col_b:
                # contagem robusta (mantém NaN como "—")
                tmp = df[[col_a, col_b]].copy()
                tmp[col_a] = tmp[col_a].astype(str).fillna("—")
                tmp[col_b] = tmp[col_b].astype(str).fillna("—")

                relacao = (
                    tmp.groupby([col_a, col_b], dropna=False)
                       .size()
                       .reset_index(name="QTD")
                )

                if not relacao.empty:
                    # Barras (Altair) — sem "value/color"
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

                    # Pizza (igual ao seu original)
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

                    # Diagnóstico original
                    top = relacao.sort_values("QTD", ascending=False).iloc[0]
                    diag = (
                        f"⚠️ Diagnóstico Preventivo:\n\n"
                        f"- A combinação **{top[col_a]} x {top[col_b]}** apresentou **{int(top['QTD'])} ocorrências**.\n"
                        f"- Recomenda-se intensificar manutenção preventiva em **{top[col_a]}**, "
                        f"com foco em evitar novos casos de **{top[col_b]}**."
                    )
                    st.success(diag)

                    # IA do Manual (PDF prioritário)
                    termo_para_consulta = str(top[coluna_manual])
                    kb_res = kb_lookup(termo_para_consulta)

                    with st.expander("🧠 Conclusão automática (Manual)", expanded=True):
                        if kb_res:
                            if usar_pdf and PDF_PAGES and PDF_BACKEND is not None and "p." in kb_res["fonte"]:
                                st.markdown(f"**Conclusão (do PDF):** {kb_res['conclusao']}")
                            else:
                                st.markdown(f"**Conclusão:** {kb_res['conclusao']}")
                            if kb_res["solucoes"]:
                                st.markdown("**Possíveis soluções:**")
                                st.markdown("\n".join([f"- {s}" for s in kb_res["solucoes"]]))
                            st.caption(f"Fonte: {kb_res['fonte']}")
                        else:
                            msg = "Não encontrei no PDF nem no CSV. Suba manual/CSV com colunas **termo, conclusao, solucoes**."
                            if usar_pdf and PDF_BACKEND is None and pdf_files:
                                msg = "Biblioteca de PDF não disponível. Adicione `pypdf` ao requirements."
                            st.info(msg)
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
