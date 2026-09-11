"""
Tema visual (CSS injetado).
Inspirado nas cores do logo LiderLimp: azul-marinho + laranja.
"""
import base64
from pathlib import Path

import streamlit as st


def _logo_base64() -> str:
    p = Path(__file__).resolve().parent.parent / "assets" / "logo.png"
    if not p.exists():
        return ""
    return base64.b64encode(p.read_bytes()).decode("utf-8")


def aplicar_tema():
    css = """
    <style>
    :root {
      --primary: #FF6A00;
      --primary-soft: #FFA45C;
      --bg: #0E1B3D;
      --bg-2: #152352;
      --card: #1B2A60;
      --text: #E8ECF7;
      --muted: #9BA7CC;
      --good: #29CC8F;
      --bad: #FF5C7A;
      --warn: #FFB547;
    }

    /* fundo geral com gradiente sutil */
    .stApp {
      background:
        radial-gradient(1200px 600px at 90% -10%, rgba(255,106,0,0.12), transparent 60%),
        radial-gradient(1000px 500px at -10% 110%, rgba(35,80,200,0.18), transparent 60%),
        linear-gradient(180deg, #0E1B3D 0%, #0B1530 100%);
    }

    /* cabeçalhos */
    h1, h2, h3, h4 { color: var(--text) !important; letter-spacing: 0.2px; }
    .lider-title {
      display:flex; align-items:center; gap:14px; margin: 4px 0 6px 0;
    }
    .lider-title .bar {
      width: 8px; height: 38px; border-radius: 4px;
      background: linear-gradient(180deg, var(--primary), var(--primary-soft));
    }
    .lider-sub { color: var(--muted); margin-top:-6px; }

    /* cards */
    .lider-card {
      background: linear-gradient(180deg, rgba(255,255,255,0.04), rgba(255,255,255,0.015));
      border: 1px solid rgba(255,255,255,0.08);
      border-radius: 14px;
      padding: 18px 18px 14px 18px;
      margin-bottom: 14px;
      box-shadow: 0 12px 30px rgba(0,0,0,0.25);
    }
    .lider-card h3 { margin: 0 0 10px 0; font-size: 1.05rem; }

    /* botões primários */
    .stButton > button[kind="primary"], .stDownloadButton > button {
      background: linear-gradient(135deg, var(--primary) 0%, #FF8A3D 100%) !important;
      color: white !important;
      border: none !important;
      border-radius: 10px !important;
      font-weight: 600 !important;
      box-shadow: 0 8px 22px rgba(255,106,0,0.35) !important;
    }
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
      filter: brightness(1.08);
    }

    /* botões secundários */
    .stButton > button:not([kind="primary"]) {
      background: rgba(255,255,255,0.06) !important;
      color: var(--text) !important;
      border: 1px solid rgba(255,255,255,0.12) !important;
      border-radius: 10px !important;
    }

    /* uploader */
    section[data-testid="stFileUploaderDropzone"] {
      background: rgba(255,255,255,0.03);
      border: 1.5px dashed rgba(255,255,255,0.18) !important;
      border-radius: 12px !important;
    }

    /* métricas */
    div[data-testid="stMetric"] {
      background: rgba(255,255,255,0.04);
      border: 1px solid rgba(255,255,255,0.08);
      border-radius: 12px;
      padding: 14px 16px;
    }

    /* sidebar */
    section[data-testid="stSidebar"] {
      background: linear-gradient(180deg, #0A1431 0%, #0E1B3D 100%);
      border-right: 1px solid rgba(255,255,255,0.06);
    }

    /* dataframe */
    .stDataFrame { border-radius: 10px; overflow: hidden; }

    /* badges */
    .badge {
      display:inline-block; padding:3px 10px; border-radius:999px;
      font-size: 0.78rem; font-weight: 600;
    }
    .badge-warn { background: rgba(255,181,71,0.16); color: var(--warn); border:1px solid rgba(255,181,71,0.35); }
    .badge-bad  { background: rgba(255,92,122,0.16); color: var(--bad);  border:1px solid rgba(255,92,122,0.35); }
    .badge-good { background: rgba(41,204,143,0.15); color: var(--good); border:1px solid rgba(41,204,143,0.35); }

    /* tabs */
    button[data-baseweb="tab"] {
      color: var(--muted) !important;
      font-weight: 600 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
      color: var(--primary) !important;
      border-bottom-color: var(--primary) !important;
    }

    /* esconde menu padrão */
    #MainMenu, footer { visibility: hidden; }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


def render_header(subtitulo: str = "Importador de Admissões · Domínio → AppLider/EasyApp"):
    logo = _logo_base64()
    logo_html = (
        f'<img src="data:image/png;base64,{logo}" style="width:54px;height:54px;border-radius:12px;box-shadow:0 6px 18px rgba(0,0,0,.35)" />'
        if logo else ""
    )
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:16px;margin:6px 0 18px 0">
          {logo_html}
          <div>
            <div class="lider-title">
              <div class="bar"></div>
              <h2 style="margin:0">LiderLimp · Importador</h2>
            </div>
            <div class="lider-sub">{subtitulo}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def card_inicio(titulo: str):
    st.markdown(f'<div class="lider-card"><h3>{titulo}</h3>', unsafe_allow_html=True)


def card_fim():
    st.markdown("</div>", unsafe_allow_html=True)
