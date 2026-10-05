"""
THEME.PY — Tema visual LiderLimp (dark navy + laranja)
Aplicado via st.markdown com CSS injetado. Mantém as funções
públicas usadas pelo streamlit_app: aplicar_tema, render_header,
card_inicio, card_fim.
"""
from __future__ import annotations

import base64
from pathlib import Path

import streamlit as st

_CSS = """
<style>
:root{
  --bg:#060B1A; --bg2:#0A1230; --card:#0E1A38; --card2:#12214A;
  --border:#1E2C52; --text:#EAF0FF; --muted:#93A3C7;
  --orange:#FF7A1A; --orange2:#FF9540; --navy:#0B2A6B;
  --red:#FF4C4C; --green:#2FD57C; --blue:#4F8CFF;
}
.stApp{
  background:
    radial-gradient(1000px 480px at 85% -10%, rgba(255,122,26,.14), transparent 60%),
    radial-gradient(900px 500px at -10% -10%, rgba(11,42,107,.55), transparent 60%),
    linear-gradient(180deg, var(--bg) 0%, #081026 100%);
  color:var(--text);
}
#MainMenu{visibility:hidden} footer{visibility:hidden} header{visibility:hidden}
.block-container{max-width:1400px; padding-top:1.4rem; padding-bottom:2.5rem}

/* ── Sidebar ── */
[data-testid="stSidebar"]{background:rgba(7,12,30,.97); border-right:1px solid var(--border)}
[data-testid="stSidebar"] *{color:var(--text)}

/* ── Hero (cabeçalho principal) ── */
.hero{
  display:flex; align-items:center; gap:18px;
  background:linear-gradient(135deg, rgba(255,122,26,.10), rgba(11,42,107,.35));
  border:1px solid var(--border); border-radius:22px;
  padding:20px 26px; margin-bottom:18px;
  box-shadow:0 12px 34px rgba(0,0,0,.35);
}
.hero img{height:56px; width:auto; border-radius:12px}
.hero h1{margin:0; font-size:1.55rem; font-weight:800; color:var(--text)}
.hero p{margin:2px 0 0 0; color:var(--muted); font-size:.92rem}
.pill{
  display:inline-block; padding:5px 12px; border-radius:999px; font-size:.75rem;
  font-weight:700; letter-spacing:.04em; margin-bottom:8px;
  background:rgba(255,122,26,.16); color:#FFC89B; border:1px solid rgba(255,122,26,.35);
}

/* ── Cards ── */
.card{
  background:var(--card); border:1px solid var(--border); border-radius:18px;
  padding:18px 20px; margin-bottom:16px; box-shadow:0 10px 26px rgba(0,0,0,.25);
}
.card-title{font-size:1rem; font-weight:800; color:var(--text); margin-bottom:4px}
.card-sub{color:var(--muted); font-size:.86rem; margin-bottom:10px}

/* ── Legendas (cores do Excel) ── */
.legend{
  display:inline-flex; align-items:center; gap:8px; font-size:.82rem; color:var(--muted);
  background:rgba(255,76,76,.12); border:1px solid rgba(255,76,76,.35);
  padding:6px 12px; border-radius:10px; margin:4px 6px 10px 0;
}
.legend .box{width:14px; height:14px; border-radius:4px; background:var(--red)}
.legend-blue{background:rgba(79,140,255,.12); border-color:rgba(79,140,255,.35)}
.legend-blue .box{background:var(--blue)}
.legend-green{background:rgba(47,213,124,.12); border-color:rgba(47,213,124,.35)}
.legend-green .box{background:var(--green)}

/* ── Botões ── */
.stButton>button{
  border-radius:12px; border:1px solid var(--border); font-weight:700;
  background:linear-gradient(180deg, #16244D, #101B3D); color:var(--text);
  padding:.62rem 1rem;
}
.stButton>button:hover{border-color:var(--orange); color:#fff}
.stButton>button[kind="primary"]{
  background:linear-gradient(180deg, var(--orange), #E6610A); color:#1A0E02; border:none;
}
.stButton>button[kind="primary"]:hover{filter:brightness(1.07)}
div[data-testid="stDownloadButton"]>button{
  border-radius:12px; font-weight:800; color:#08102A; border:none;
  background:linear-gradient(180deg, var(--orange2), var(--orange));
}
div[data-testid="stDownloadButton"]>button:hover{filter:brightness(1.08)}

/* ── Métricas ── */
div[data-testid="stMetric"]{
  background:var(--card); border:1px solid var(--border); border-radius:16px;
  padding:12px 16px; box-shadow:0 8px 20px rgba(0,0,0,.22);
}
div[data-testid="stMetricLabel"]{color:var(--muted)}
div[data-testid="stMetricValue"]{color:var(--text); font-weight:800}

/* ── Inputs / uploaders / checkbox ── */
div[data-testid="stFileUploader"]{
  background:var(--card); border:1px dashed var(--border); border-radius:16px; padding:8px 12px;
}
div[data-testid="stFileUploader"]:hover{border-color:var(--orange)}
.stTextInput input, .stNumberInput input{
  background:var(--card2)!important; border:1px solid var(--border)!important;
  border-radius:12px!important; color:var(--text)!important;
}
div[data-testid="stCheckbox"] label{color:var(--text)}

/* ── Tabelas / abas / alertas ── */
div[data-testid="stDataFrame"]{border:1px solid var(--border); border-radius:14px; overflow:hidden}
.stTabs [data-baseweb="tab-list"]{gap:8px}
.stTabs [data-baseweb="tab"]{
  border-radius:12px; padding:9px 16px; background:rgba(255,255,255,.03); color:var(--muted);
  font-weight:700;
}
.stTabs [aria-selected="true"]{background:rgba(255,122,26,.18)!important; color:#FFD3AE!important}
.stAlert{border-radius:14px}
hr{border-color:var(--border)}
</style>
"""


def aplicar_tema() -> None:
    """Injeta o CSS do tema na página. Chamar uma vez no início do app."""
    st.markdown(_CSS, unsafe_allow_html=True)


def _logo_b64() -> str:
    """Logo em base64 (se existir em app/assets/logo.png)."""
    logo = Path(__file__).resolve().parent.parent / "assets" / "logo.png"
    if logo.exists():
        return base64.b64encode(logo.read_bytes()).decode()
    return ""


def render_header(*_args, **_kwargs) -> None:
    """Cabeçalho hero com logo + título do app."""
    img = _logo_b64()
    img_tag = f'<img src="data:image/png;base64,{img}" alt="LiderLimp">' if img else ""
    st.markdown(
        f"""
<div class="hero">
  {img_tag}
  <div>
    <div class="pill">LIDERLIMP · FERRAMENTA INTERNA</div>
    <h1>🟧 Importador de Admissões</h1>
    <p>Domínio → AppLider/EasyApp · Layout, Benefícios, CSV e verificação de duplicados</p>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def card_inicio(titulo: str = "", subtitulo: str = "", *_args, **_kwargs) -> None:
    """Abre um card visual (fechar com card_fim)."""
    sub = f'<div class="card-sub">{subtitulo}</div>' if subtitulo else ""
    st.markdown(
        f'<div class="card"><div class="card-title">{titulo}</div>{sub}',
        unsafe_allow_html=True,
    )


def card_fim(*_args, **_kwargs) -> None:
    """Fecha o card aberto por card_inicio."""
    st.markdown("</div>", unsafe_allow_html=True)
