# app.py — Analisador Dinâmico de Planilhas
# Barras (Altair), pizza (matplotlib), diagnóstico original,
# IA do manual com PDF robusta (regex + heurística) + CSV fallback.
# Filtros anti-tabela/índice; plano de ação amarrado ao diagnóstico; modo 5W2H.

import io
import re
import unicodedata
from difflib import get_close_matches
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

# --- só para a pizza (mantém igual ao seu código) ---
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

# --- backends de PDF (usa o que estiver instalado) ---
PDF_BACKEND = None
try:
    import pypdf
    PDF_BACKEND = "pypdf"
except Exception:
    try:
        import PyPDF2
        PDF_BACKEND = "PyPDF2"
    except Exception:
        PDF_BACKEND = None

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

# ===========================
# Sidebar: Manual + Plano
# ===========================

pdf_files = st.sidebar.file_uploader(
    "Suba **um ou mais PDFs** do manual",
    type=["pdf"], key="pdfs", accept_multiple_files=True
)
usar_pdf = st.sidebar.toggle("🔎 Usar manual em PDF com prioridade", value=True)
st.sidebar.caption(f"PDF backend: {PDF_BACKEND or 'nenhum'}")

st.sidebar.markdown("### 🧰 Plano de ação")
modo_5w2h = st.sidebar.toggle("📋 Modo 5W2H", value=False)
responsavel_padrao = st.sidebar.text_input("Responsável padrão", "Manutenção")
prazo_dias = st.sidebar.number_input("Prazo (dias)", min_value=1, max_value=60, value=7, step=1)
custo_padrao = st.sidebar.text_input("Custo (estimado)", "—")

# ===========================
# Normalização / tokens
# ===========================
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s).lower())
    s = "".join(ch for ch in s if ch.isalnum() or ch.isspace() or ch in "-_/")
    return " ".join(s.split())

def _tokens(s: str) -> set:
    return set(_norm(s).split())

def _score_jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)

# ===========================
# Carregar CSV do manual (se houver). Aqui vamos passar None.
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
    ren = {}
    for need in ["termo", "conclusao", "solucoes"]:
        if need in cols:
            ren[cols[need]] = need
    df = df.rename(columns=ren)
    for need in ["termo", "conclusao", "solucoes"]:
        if need not in df.columns:
            df[need] = ""
    df["termo_norm"] = df["termo"].map(_norm)
    df["tokens"] = df["termo"].map(lambda s: _tokens(str(s)))
    return df[["termo", "termo_norm", "tokens", "conclusao", "solucoes"]]

# Fallback embutido (pode editar/expandir)
KB_DEFAULT = pd.DataFrame([
    {"termo":"LOW_PRESSURE","conclusao":"Pressão baixa no circuito de tinta/jet.",
     "solucoes":"Verificar nível de solvente;Checar mangueiras e engates;Executar rotina de pressurização;Inspecionar vazamentos;Testar transdutor/bomba"},
    {"termo":"ALTA VISCOSIDADE","conclusao":"Viscosidade acima da janela.",
     "solucoes":"Checar make-up/solvente;Executar rotina de diluição;Verificar sensor de viscosidade;Ajustar temperatura ambiente"},
    {"termo":"NOZZLE CLOG / ENTUPIMENTO DE BICO","conclusao":"Bico/jato possivelmente obstruído.",
     "solucoes":"Limpar cabeça de impressão;Aplicar flush;Trocar/limpar filtro;Checar qualidade da tinta;Agendar preventiva"},
    {"termo":"MISALIGNMENT / DESALINHADO","conclusao":"Cabeça desalinhada do produto.",
     "solucoes":"Ajustar distância/ângulo;Fixar suportes;Revisar gabarito/guia;Testar leitura"},
    {"termo":"IMPRESSÃO CLARA / FADED","conclusao":"Baixa densidade/contraste de marcação.",
     "solucoes":"Ajustar velocidade/atraso;Checar tinta/solvente;Limpar bico/eletrodos;Revisar distância cabeça-produto"},
])
KB_DEFAULT["termo_norm"] = KB_DEFAULT["termo"].map(_norm)
KB_DEFAULT["tokens"] = KB_DEFAULT["termo"].map(lambda s: _tokens(str(s)))

# >>>>>>> correção: sem 'kb_file'; usamos None e seguimos com o fallback/CSV interno
kb_user = load_kb(None)
KB_ALL = (pd.concat([kb_user, KB_DEFAULT], ignore_index=True)
          if not kb_user.empty else KB_DEFAULT)

# Aliases para casar nomes da sua planilha com o manual
ALIASES = {
    "falha de jato": "nozzle clog / entupimento de bico",
    "jato falhando": "nozzle clog / entupimento de bico",
    "cabeçote sujo": "cabeçote requer limpeza ao desligar",
    "ausencia de impressao": "impressão clara / faded",
    "perda de modulacao": "impressão clara / faded",
    "baixa pressao": "low pressure",
}
ALIASES_NORM = { _norm(k): _norm(v) for k, v in ALIASES.items() }

# ===========================
# Ler PDFs (texto por página)
# ===========================
@st.cache_data(show_spinner=False)
def read_pdfs(files, backend):
    pages = []  # [{source,page,text,norm,tokens}]
    if not files or backend is None:
        return pages
    for f in files:
        try:
            data = f.getvalue()
            if backend == "pypdf":
                reader = pypdf.PdfReader(io.BytesIO(data))
            else:
                reader = PyPDF2.PdfReader(io.BytesIO(data))
            for i, p in enumerate(reader.pages, start=1):
                try:
                    txt = p.extract_text() or ""
                except Exception:
                    txt = ""
                norm = _norm(txt)
                pages.append({"source":f.name, "page":i, "text":txt,
                              "norm":norm, "tokens":_tokens(norm)})
        except Exception:
            continue
    return pages

PDF_PAGES = read_pdfs(pdf_files, PDF_BACKEND)

# ===========================
# Helpers anti-índice/tabela/cabeçalho
# ===========================
_BAD_WORDS = {"página","pagina","page","tabela","table","índice","indice",
              "sumário","sumario","conteúdo","conteudo","seção","secao",
              "capítulo","capitulo","manual"}

def _looks_like_toc_entry(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return False
    if re.search(r"\.{2,}\s*\d+(\s*[–—-]\s*\d+)?$", s):
        return True
    if re.search(r"[\.·\s]{3,}\d+(\s*[–—-]\s*\d+)?$", s):
        return True
    if re.search(r"\b(se[cç][aã]o|section|cap[ií]tulo|chapter)\b.*\d+\s*[–—-]\s*\d+$", s, re.I):
        return True
    return False

def _looks_like_table_line(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return False
    if re.fullmatch(r"(?:[\d\.\-\/\s]|N/?A)+", s, flags=re.I):
        return True
    if re.search(r"(?:\d[\d\./-]*\s+){3,}\d[\d\./-]*", s):
        return True
    tokens = s.split()
    if tokens and sum(t.isupper() and len(t) <= 4 for t in tokens) / len(tokens) > 0.7:
        return True
    return False

def _is_texty(s: str) -> bool:
    s = (s or "").strip()
    if not s or len(s) < 8:
        return False
    if _looks_like_toc_entry(s) or _looks_like_table_line(s):
        return False
    sn = s.lower()
    if any(w in sn for w in _BAD_WORDS):
        return False
    letters = sum(ch.isalpha() for ch in s)
    digits  = sum(ch.isdigit()  for ch in s)
    if digits > letters * 0.6:
        return False
    if sum(1 for t in s.split() if any(c.isalpha() for c in t)) < 2:
        return False
    return True

def _first_informative_line(block: str) -> str:
    KEYS = ("causa","cause","sintoma","symptom","descrição","description",
            "ação","acao","procedimento","solução","solucao","falha","fault",
            "avaria","problem","issue")
    lines_all = [ln.strip() for ln in (block or "").splitlines()]
    lines = [ln for ln in lines_all if _is_texty(ln)]
    if not lines:
        return ""
    for ln in lines:
        if any(k in _norm(ln) for k in KEYS):
            return ln
    for ln in lines:
        if len(ln) >= 20 and re.search(r"[\.:\-;]", ln):
            return ln
    return lines[0]

def _extract_steps(text: str, max_lines: int = 12) -> list:
    VERBS = ("verificar","checar","inspecionar","executar","limpar","ajustar",
             "substituir","alinhar","recolocar","testar","revisar","aguardar",
             "apertar","conferir","drenar","pressurizar","desobstruir")
    steps = []
    for raw in (text or "").splitlines():
        raw = raw.strip()
        if not raw or not _is_texty(raw):
            continue
        if _looks_like_table_line(raw):
            continue
        ln_norm = _norm(raw)
        is_bullet = raw.startswith(("-", "•", "·", "*")) or (
            len(raw) > 2 and raw[:2].isdigit() and raw[2:3] in ".-) "
        )
        has_verb = any(ln_norm.startswith(v) or f" {v} " in ln_norm for v in VERBS)
        if is_bullet or has_verb:
            clean = raw.lstrip("-•·* ").strip()
            if len(clean) >= 10:
                steps.append(clean)
        if len(steps) >= max_lines:
            break
    uniq, seen = [], set()
    for s in steps:
        k = _norm(s)
        if k and k not in seen:
            uniq.append(s)
            seen.add(k)
    return uniq

# ===========================
# Mapeamento por REGEX → seções do manual
# ===========================
MANUAL_CODES = {
    "3.18 Baixa pressão": {
        "keys": "baixa pressão low pressure pressao baixa fc006",
        "patterns": [r"\b3\.18\b.*baix[ao]\s+press", r"\blow[\s\-]?pressure\b", r"\bFC0?06\b"],
    },
    "3.28 Cabeçote requer limpeza ao desligar": {
        "keys": "cabeçote limpeza head clean shutdown",
        "patterns": [r"\b3\.28\b.*cabe(ç|c)ote.*limpez", r"\b(head|printhead).*(clean|cleaning)"],
    },
    "3.20 Sem Tempo de Voo (TOF)": {
        "keys": "sem tof no tof tempo de voo time of flight",
        "patterns": [r"\b3\.20\b.*(tempo\s+de\s+voo|time\s+of\s+flight|TOF)"],
    },
    "2.12 Viscosidade": {
        "keys": "viscosidade viscosity 2.12",
        "patterns": [r"\b2\.12\b.*viscos", r"\bviscosit"],
    },
    "2.03 Tempo de Voo": {
        "keys": "tempo de voo time of flight 2.03",
        "patterns": [r"\b2\.03\b.*(tempo\s+de\s+voo|time\s+of\s+flight)"],
    },
}

ALIASES_REGEX = [
    (re.compile(r"\bfalha (do|de) jato\b", re.I), ["3.18 Baixa pressão", "3.20 Sem Tempo de Voo (TOF)"]),
    (re.compile(r"\bcabe(ç|c)ote.*sujo\b", re.I), ["3.28 Cabeçote requer limpeza ao desligar"]),
    (re.compile(r"\bpress(ã|a)o baixa\b", re.I), ["3.18 Baixa pressão"]),
    (re.compile(r"\bviscosidad", re.I), ["2.12 Viscosidade"]),
    (re.compile(r"\btempo\s+de\s+voo\b|\bTOF\b", re.I), ["3.20 Sem Tempo de Voo (TOF)", "2.03 Tempo de Voo"]),
]

@st.cache_data(show_spinner=False)
def index_pdf_by_codes(pages):
    if not pages:
        return {}
    compiled = {
        code: [re.compile(pat, re.I) for pat in data["patterns"]]
        for code, data in MANUAL_CODES.items()
    }
    hits = {code: [] for code in MANUAL_CODES.keys()}
    for p in pages:
        txt = p["text"] or ""
        for code, regs in compiled.items():
            if any(r.search(txt) for r in regs):
                hits[code].append(p)
    return hits

PDF_CODE_HITS = index_pdf_by_codes(PDF_PAGES)

def _candidate_codes_from_term(term_norm: str) -> list:
    cands = set()
    for rgx, codes in ALIASES_REGEX:
        if rgx.search(term_norm):
            cands.update(codes)
    for code, meta in MANUAL_CODES.items():
        sim = _score_jaccard(_tokens(term_norm), _tokens(meta["keys"]))
        if sim >= 0.25:
            cands.add(code)
    mapped = ALIASES_NORM.get(term_norm, None)
    if mapped:
        for code, meta in MANUAL_CODES.items():
            if mapped in meta["keys"]:
                cands.add(code)
    return list(cands)

def kb_lookup_pdf_regex(term: str):
    if not PDF_PAGES or not PDF_CODE_HITS:
        return None
    tn = _norm(term)
    codes = _candidate_codes_from_term(tn)
    if not codes:
        return None
    ranked = []
    for code in codes:
        pages = PDF_CODE_HITS.get(code, [])
        if not pages:
            continue
        sim = _score_jaccard(_tokens(tn), _tokens(MANUAL_CODES[code]["keys"]))
        score = len(pages)*0.8 + sim*2.2
        ranked.append((score, code, pages))
    if not ranked:
        return None
    ranked.sort(key=lambda x: x[0], reverse=True)
    _, best_code, best_pages = ranked[0]

    use_pages = best_pages[:2]
    concl = ""
    all_steps = []
    fontes = []

    patterns = [re.compile(p, re.I) for p in MANUAL_CODES[best_code]["patterns"]]
    for p in use_pages:
        txt = p["text"] or ""
        if not txt.strip():
            continue
        mpos = None
        for rgx in patterns:
            m = rgx.search(txt)
            if m:
                mpos = m.start()
                break
        start = max(0, (mpos if mpos is not None else 0) - 400)
        window = txt[start:start+1500]
        window = "\n".join(ln for ln in window.splitlines() if _is_texty(ln))
        if not concl:
            concl = _first_informative_line(window)
        blocos = re.split(r"(?i)\b(procedimento|ação recomendada|acao recomendada|solução|solucao|diagnóstico|diagnostico|fluxograma)\b[:\-]?", window)
        alvo = window if len(blocos) < 3 else "".join(blocos[2:])
        steps_here = _extract_steps(alvo, max_lines=10)
        if len(steps_here) < 2:
            steps_here = _extract_steps(window, max_lines=10)
        all_steps.extend(steps_here)
        fontes.append(f"{p['source']} p.{p['page']}")

    steps = [s for s in all_steps if _is_texty(s) and not _looks_like_table_line(s)]
    if (not _is_texty(concl)) and len(steps) < 2:
        return None

    return {
        "conclusao": concl.strip() if _is_texty(concl) else "",
        "solucoes": steps[:10],
        "fonte": f"{best_code} — " + ", ".join(fontes),
    }

def kb_lookup_pdf_heuristic(term: str):
    if not PDF_PAGES:
        return None
    qn = _norm(term)
    qn = ALIASES_NORM.get(qn, qn)
    q_tokens = _tokens(qn)

    scored = []
    for p in PDF_PAGES:
        if not p["norm"]:
            continue
        token_hits = sum(p["norm"].count(t) for t in q_tokens)
        jac = _score_jaccard(q_tokens, p["tokens"])
        exact = 1 if qn in p["norm"] else 0
        score = token_hits*0.25 + jac*2.0 + exact*1.0
        if score > 0:
            scored.append((score, p))

    if not scored:
        return None

    scored.sort(key=lambda x: x[0], reverse=True)
    top_pages = [s[1] for s in scored[:2]]

    concl = ""
    all_steps = []
    fontes = []

    for p in top_pages:
        txt, norm = p["text"], p["norm"]
        pos = norm.find(qn)
        start = max(0, pos-400) if pos >= 0 else 0
        window_raw = txt[start:start+1400]
        window = "\n".join(ln for ln in window_raw.splitlines() if _is_texty(ln))
        if not concl:
            concl = _first_informative_line(window)
        steps_here = _extract_steps(window, max_lines=8)
        all_steps.extend(steps_here)
        fontes.append(f"{p['source']} p.{p['page']}")

    steps = [s for s in all_steps if _is_texty(s)]
    if (not _is_texty(concl)) and len(steps) < 2:
        return None

    return {
        "conclusao": concl.strip() if _is_texty(concl) else "",
        "solucoes": steps[:10],
        "fonte": ", ".join(fontes),
    }

# ===========================
# Lookup no CSV (user + fallback)
# ===========================
def kb_lookup_csv(term: str, cutoff_close=0.65, cutoff_jacc=0.35):
    if KB_ALL.empty:
        return None
    tn = _norm(term)
    target = ALIASES_NORM.get(tn, tn)
    hit = KB_ALL[KB_ALL["termo_norm"] == target]
    if not hit.empty:
        row = hit.iloc[0]
        sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
        return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": "manual CSV/alias"}

    cand = get_close_matches(tn, KB_ALL["termo_norm"].tolist(), n=1, cutoff=cutoff_close)
    if cand:
        row = KB_ALL.loc[KB_ALL["termo_norm"] == cand[0]].iloc[0]
        sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
        return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": "manual CSV/fuzzy"}

    tA = _tokens(tn)
    if tA:
        j = KB_ALL["tokens"].apply(lambda tB: _score_jaccard(tA, tB))
        idx = int(j.idxmax())
        if j.iloc[idx] >= cutoff_jacc:
            row = KB_ALL.iloc[idx]
            sols = [s.strip() for s in str(row["solucoes"]).split(";") if str(s).strip()]
            return {"conclusao": str(row["conclusao"]).strip(), "solucoes": sols, "fonte": f"manual CSV/jaccard {j.iloc[idx]:.2f}"}
    return None

def kb_lookup(term: str, prefer_pdf=True):
    """Orquestração: Regex→PDF → Heurística→PDF → CSV."""
    if prefer_pdf and PDF_PAGES and PDF_BACKEND is not None:
        hit = kb_lookup_pdf_regex(term)
        if not hit:
            hit = kb_lookup_pdf_heuristic(term)
        if hit and (hit["conclusao"] or hit["solucoes"]):
            return hit
    return kb_lookup_csv(term)

# ===========================
# Ações baseadas no diagnóstico + 5W2H
# ===========================
CONTEXT_FALLBACKS = {
    "surto de tensao": [
        "Instalar/disponibilizar DPS (classe II) no quadro da linha e no ponto da impressora {ASSET}",
        "Verificar aterramento (< 10 Ω) e equipotencialização do chassi/cabeçote da {ASSET}",
        "Colocar a impressora em UPS on-line (tempo de autonomia compatível com paradas)",
        "Inspecionar cabos de alimentação, borneiras e reaperto de conexões; procurar marcas de arco",
        "Checar fonte de alimentação/PSU e eletrônica associada após evento de pico; executar autotestes",
        "Revisar qualidade da rede (medição de surtos/harmônicos por 72h) e acionar manutenção elétrica"
    ],
    "baixa pressao": [
        "Checar nível de solvente/tinta e pressurizar o circuito",
        "Inspecionar mangueiras/engates por vazamento ou dobra",
        "Substituir filtro principal/linha se houver restrição de fluxo",
        "Testar transdutor de pressão e bomba (elétrica e mecânica)",
        "Verificar manifold/tubulação e limpeza do circuito"
    ],
    "viscosidade": [
        "Conferir make-up/solvente e executar rotina de diluição",
        "Verificar sensor de viscosidade/temperatura",
        "Ajustar temperatura ambiente e condições do processo",
        "Não adicionar solvente manualmente quando a impressora sinalizar medições corretas"
    ],
    "cabeçote requer limpeza": [
        "Executar procedimento de limpeza de cabeçote conforme manual",
        "Inspecionar bico/eletrodos e distância cabeça–produto",
        "Implementar rotina de limpeza ao desligar (padronização de turno)"
    ],
    "tof": [
        "Verificar estabilidade do jato e da modulação",
        "Ajustar posição/ângulo do cabeçote e checar TOF",
        "Corrigir condições de viscosidade/temperatura"
    ],
}

def _pick_fallback(defect_norm: str) -> list[str]:
    for key in CONTEXT_FALLBACKS.keys():
        if key in defect_norm:
            return CONTEXT_FALLBACKS[key]
    if "tempo de voo" in defect_norm or "tof" in defect_norm:
        return CONTEXT_FALLBACKS["tof"]
    if "cabeçote" in defect_norm and "limpeza" in defect_norm:
        return CONTEXT_FALLBACKS["cabeçote requer limpeza"]
    return []

def build_actions_from_diagnostic(top_row, col_a, col_b, kb_res,
                                  resp_padrao="Manutenção", prazo=7, custo="—",
                                  modo_5w2h=False):
    asset  = str(top_row[col_a])
    defect = str(top_row[col_b])
    occ    = int(top_row["QTD"])
    defect_norm = _norm(defect)

    steps = []
    fonte = ""
    if kb_res and kb_res.get("solucoes"):
        steps = list(kb_res["solucoes"])
        fonte = kb_res.get("fonte", "")
    if not steps:
        steps = _pick_fallback(defect_norm)
    steps = [s.replace("{ASSET}", asset) for s in steps]

    if not modo_5w2h:
        plano = []
        plano.append(f"**Contexto:** `{asset} x {defect}` — **{occ}** ocorrências (janela analisada).")
        if kb_res and kb_res.get("conclusao"):
            plano.append(f"**Conclusão (manual):** {kb_res['conclusao']}")
        if steps:
            plano.append("**Ações prioritárias:**")
            plano.extend([f"- {s}" for s in steps[:8]])
        else:
            plano.append("**Ações prioritárias:** (não encontradas no manual/fallback)")
        plano.append("**Checklist de execução (rápido):** responsável definido • prazo curto • evidência (foto/log) • reteste pós-ação • registrar ocorrência")
        if fonte:
            plano.append(f"*Fonte:* {fonte}")
        return "\n\n".join(plano)

    deadline = (datetime.today() + timedelta(days=prazo)).strftime("%Y-%m-%d")
    rows = []
    why = f"Reduzir recorrência de **{defect}** no ativo **{asset}** ({occ} ocorrências)"
    how_src = "Manual/CSV" if kb_res else "Boas práticas (fallback)"
    for s in steps[:10] if steps else ["—"]:
        rows.append({
            "What (o quê)": s,
            "Why (por quê)": why,
            "Where (onde)": asset,
            "Who (quem)": resp_padrao,
            "When (quando)": deadline,
            "How (como)": how_src,
            "How much (quanto)": custo,
        })
    df5w2h = pd.DataFrame(rows)
    if fonte:
        df5w2h.attrs["fonte"] = fonte
    return df5w2h

# ===========================
# Leitura da planilha do usuário
# ===========================
def read_any(file):
    if file.name.lower().endswith(".xlsx"):
        xls = pd.ExcelFile(file)
        sheet = st.selectbox("📑 Escolha a aba", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=sheet)
    else:
        try:
            df = pd.read_csv(file, sep=None, engine="python")
        except Exception:
            df = pd.read_csv(file)
    return df

uploaded_file = st.file_uploader("📂 Suba sua planilha (.xlsx ou .csv)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        df = read_any(uploaded_file)

        st.subheader("🔎 Pré-visualização")
        st.dataframe(df.head(), use_container_width=True)

        cols = df.columns.tolist()
        st.write("📋 Colunas detectadas:", cols)

        # ==============================
        # Relação entre duas colunas categóricas
        # ==============================
        st.subheader("📊 Relação entre duas colunas categóricas")
        col_a = st.selectbox("👉 Primeira coluna categórica", cols, key="cata")
        col_b = st.selectbox("👉 Segunda coluna categórica", cols, key="catb")

        coluna_manual = st.radio("Consultar manual usando qual coluna?",
                                 [col_b, col_a], index=0, horizontal=True)

        relacao, diag = None, None
        try:
            if col_a and col_b:
                tmp = df[[col_a, col_b]].copy()
                tmp[col_a] = tmp[col_a].astype(str).fillna("—")
                tmp[col_b] = tmp[col_b].astype(str).fillna("—")

                relacao = (
                    tmp.groupby([col_a, col_b], dropna=False)
                       .size()
                       .reset_index(name="QTD")
                )

                if not relacao.empty:
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

                    if MATPLOTLIB_OK:
                        st.subheader(f"🥧 Distribuição de {col_b}")
                        dist = df[col_b].value_counts(normalize=True) * 100
                        dist = dist.sort_values(ascending=False)
                        outros = dist[dist < 5].sum()
                        dist = dist[dist >= 5]
                        if outros > 0:
                            dist["Outros"] = outros
                        fig, ax = plt.subplots(figsize=(7, 6))
                        wedges, texts, autotexts = ax.pie(
                            dist, autopct="%1.1f%%", startangle=90,
                            counterclock=False, colors=plt.cm.tab20.colors
                        )
                        ax.legend(wedges, dist.index, title=f"{col_b}",
                                  loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        ax.set_title(f"Distribuição de {col_b}", fontsize=14)
                        st.pyplot(fig)

                    top = relacao.sort_values("QTD", ascending=False).iloc[0]
                    diag = (
                        f"⚠️ Diagnóstico Preventivo:\n\n"
                        f"- A combinação **{top[col_a]} x {top[col_b]}** apresentou **{int(top['QTD'])} ocorrências**.\n"
                        f"- Recomenda-se intensificar manutenção preventiva em **{top[col_a]}**, "
                        f"com foco em evitar novos casos de **{top[col_b]}**."
                    )
                    st.success(diag)

                    termo_para_consulta = str(top[coluna_manual])
                    kb_res = kb_lookup(termo_para_consulta, prefer_pdf=usar_pdf)

                    with st.expander("🧠 Conclusão & Plano (baseado no diagnóstico)", expanded=True):
                        plano = build_actions_from_diagnostic(
                            top, col_a, col_b, kb_res,
                            resp_padrao=responsavel_padrao,
                            prazo=prazo_dias,
                            custo=custo_padrao,
                            modo_5w2h=modo_5w2h
                        )
                        if modo_5w2h:
                            st.dataframe(plano, use_container_width=True)
                            if hasattr(plano, "attrs") and "fonte" in plano.attrs and plano.attrs["fonte"]:
                                st.caption(f"Fonte: {plano.attrs['fonte']}")
                            csv = plano.to_csv(index=False).encode("utf-8")
                            st.download_button("⬇️ Baixar 5W2H (CSV)", csv, file_name="plano_5w2h.csv")
                        else:
                            st.markdown(plano)

                        if (not modo_5w2h) and kb_res and kb_res.get("fonte"):
                            st.caption(f"Fonte: {kb_res['fonte']}")

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
