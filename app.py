# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════╗
║        IMPORTADOR DOMÍNIO → APPLIDER (EasyApp) — LÍDER LIMPE            ║
║                                                                          ║
║  Fluxo:                                                                  ║
║    🧹 Duplicados   → cruza Domínio x Sistema interno, remove quem já     ║
║                      existe e gera Arquivo Domínio LIMPO                 ║
║    1️⃣ Layout       → Domínio + Mapeamento → Importar Layout.xlsx         ║
║                      (campos faltantes ficam VERMELHOS)                  ║
║    2️⃣ Benefícios   → Layout + Relação Benefícios → Importar              ║
║                      Beneficios.xlsx (ID eSocial ou sequencial)          ║
║    3️⃣ CSV          → converte os .xlsx gerados em CSV preservando texto  ║
║    ⭐ Tudo          → executa o pipeline completo de uma vez             ║
║                                                                          ║
║  Downloads NÃO reprocessam nada: todos os arquivos ficam prontos em      ║
║  memória (session_state) — baixe quantos quiser, na ordem que quiser.    ║
╚══════════════════════════════════════════════════════════════════════════╝
"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from core.benefits_engine import processar_beneficios_df
from core.duplicates import comparar_colaboradores
from core.excel_io import read_excel_smart, xlsx_bytes_to_csv_bytes
from core.exporters import beneficios_xlsx_bytes, df_to_xlsx_bytes
from core.layout_engine import (
    CAMPOS_OBRIGATORIOS,
    ROTULOS_CAMPOS,
    carregar_todos_os_mapas,
    processar_layout,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="Importador Domínio → AppLider | Líder Limpe",
    page_icon="🧡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════
# TEMA DARK PREMIUM (cores do logo Líder Limpe: azul-marinho + laranja)
# ══════════════════════════════════════════════════════════════════════════
st.markdown(
    """
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
[data-testid="stSidebar"]{background:rgba(7,12,30,.97); border-right:1px solid var(--border)}
[data-testid="stSidebar"] *{color:var(--text)}

.hero{
  display:flex; align-items:center; gap:18px;
  background:linear-gradient(135deg, rgba(255,122,26,.10), rgba(11,42,107,.35));
  border:1px solid var(--border); border-radius:22px;
  padding:20px 26px; margin-bottom:18px;
  box-shadow:0 12px 34px rgba(0,0,0,.35);
}
.hero h1{margin:0; font-size:1.55rem; font-weight:800; color:var(--text)}
.hero p{margin:2px 0 0 0; color:var(--muted); font-size:.92rem}
.pill{
  display:inline-block; padding:5px 12px; border-radius:999px; font-size:.75rem;
  font-weight:700; letter-spacing:.04em; margin-bottom:8px;
  background:rgba(255,122,26,.16); color:#FFC89B; border:1px solid rgba(255,122,26,.35);
}
.card{
  background:var(--card); border:1px solid var(--border); border-radius:18px;
  padding:18px 20px; margin-bottom:16px; box-shadow:0 10px 26px rgba(0,0,0,.25);
}
.card-title{font-size:1rem; font-weight:800; color:var(--text); margin-bottom:4px}
.card-sub{color:var(--muted); font-size:.86rem; margin-bottom:10px}
.legend{
  display:inline-flex; align-items:center; gap:8px; font-size:.82rem; color:var(--muted);
  background:rgba(255,76,76,.12); border:1px solid rgba(255,76,76,.35);
  padding:6px 12px; border-radius:10px; margin:4px 0 10px 0;
}
.legend .box{width:14px; height:14px; border-radius:4px; background:var(--red)}
.legend-blue{background:rgba(79,140,255,.12); border-color:rgba(79,140,255,.35)}
.legend-blue .box{background:var(--blue)}

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
div[data-testid="stDataFrame"]{border:1px solid var(--border); border-radius:14px; overflow:hidden}
.stTabs [data-baseweb="tab-list"]{gap:8px}
.stTabs [data-baseweb="tab"]{
  border-radius:12px; padding:9px 16px; background:rgba(255,255,255,.03); color:var(--muted);
  font-weight:700;
}
.stTabs [aria-selected="true"]{background:rgba(255,122,26,.18)!important; color:#FFD3AE!important}
.stAlert{border-radius:14px}
</style>
""",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
ss = st.session_state
ss.setdefault("artifacts", {})       # nome -> {data, filename, mime}
ss.setdefault("ultimo_id", 0)        # último ID cadastrado no EasyApp
ss.setdefault("df_layout", None)     # Importar Layout gerado
ss.setdefault("df_dom_limpo", None)  # Domínio sem duplicados
ss.setdefault("df_dom_original", None)


def put_art(key, data, filename, mime):
    ss["artifacts"][key] = {"data": data, "filename": filename, "mime": mime}


def dl(key, label):
    """Botão de download que NUNCA reprocessa: lê bytes prontos da sessão."""
    art = ss["artifacts"].get(key)
    if art:
        st.download_button(
            label,
            art["data"],
            art["filename"],
            art["mime"],
            key=f"dl_{key}",
            use_container_width=True,
        )
        return True
    return False


def get_base_bytes(upload, default_name):
    if upload is not None:
        return upload.getvalue()
    p = DATA_DIR / default_name
    return p.read_bytes() if p.exists() else None


def style_missing(df):
    def hl(row):
        styles = [""] * len(row)
        for c in CAMPOS_OBRIGATORIOS:
            if c in df.columns:
                v = row[c]
                if pd.isna(v) or str(v).strip() in ("", "nan", "None", "NaN"):
                    styles[list(df.columns).index(c)] = (
                        "background-color:#7F1D1D;color:#FECACA;font-weight:700"
                    )
        return styles

    return df.style.apply(hl, axis=1)


def metricas(items):
    cols = st.columns(len(items))
    for c, (label, value) in zip(cols, items):
        c.metric(label, value)


# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    logo = BASE_DIR / "assets" / "logo.png"
    if logo.exists():
        st.image(str(logo), width=110)
    st.markdown("## ⚙️ Arquivos base")
    st.caption("Se não enviar, o app usa os arquivos padrão embutidos no repositório (pasta `data/`).")

    up_mapa = st.file_uploader("Mapeamento Sistema.xls", type=["xls", "xlsx"], key="up_mapa")
    up_modelo = st.file_uploader("Colunas Originais.xlsx", type=["xlsx"], key="up_modelo")
    up_benef = st.file_uploader("BENEFICIOS - RELAÇÃO.xlsx", type=["xlsx"], key="up_benef")

    for nome, up, arq in [
        ("Mapeamento", up_mapa, "Mapeamento Sistema.xls"),
        ("Modelo de colunas", up_modelo, "Colunas Originais.xlsx"),
        ("Relação de benefícios", up_benef, "BENEFICIOS - RELACAO.xlsx"),
    ]:
        if up is not None:
            st.caption(f"✅ {nome}: arquivo enviado")
        elif (DATA_DIR / arq).exists():
            st.caption(f"📦 {nome}: padrão embutido")
        else:
            st.caption(f"⚠️ {nome}: faltando")

    st.markdown("---")
    st.markdown("## 🔢 Último ID EasyApp")
    st.number_input(
        "Último ID cadastrado no sistema",
        min_value=0,
        step=1,
        key="ultimo_id",
        help="Usado para gerar os próximos IDs em sequência na planilha de benefícios.",
    )
    st.caption(f"Próximo ID será: **{ss['ultimo_id'] + 1}**")

    st.markdown("---")
    st.caption("Importador Domínio → AppLider • v2.0 Streamlit")

# ══════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════
logo_html = ""
if logo.exists():
    import base64

    b64 = base64.b64encode(logo.read_bytes()).decode()
    logo_html = f'<img src="data:image/png;base64,{b64}" style="width:74px;height:74px;object-fit:contain;border-radius:16px;background:rgba(255,255,255,.05);padding:6px;border:1px solid rgba(255,255,255,.08)"/>'

st.markdown(
    f"""
<div class="hero">
  {logo_html}
  <div>
    <span class="pill">LÍDER LIMPE • IMPORTAÇÃO DE ADMISSÕES</span>
    <h1>Importador Domínio → AppLider (EasyApp)</h1>
    <p>Duplicados • Importar Layout • Benefícios • CSV — tudo em um só lugar, sem refazer processamento para baixar arquivos.</p>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════
tab_dup, tab_layout, tab_benef, tab_csv, tab_all = st.tabs(
    [
        "🧹 Duplicados",
        "1️⃣ Importar Layout",
        "2️⃣ Benefícios",
        "3️⃣ Converter CSV",
        "⭐ Executar Tudo",
    ]
)

# ──────────────────────────────────────────────────────────────────────────
# ABA 0 — DUPLICADOS
# ──────────────────────────────────────────────────────────────────────────
with tab_dup:
    st.markdown(
        """
<div class="card">
  <div class="card-title">🧹 Verificação de colaboradores duplicados</div>
  <div class="card-sub">
    Envie o <b>Arquivo Domínio</b> (novas admissões) e a <b>planilha de colaboradores exportada do sistema interno</b>.
    O app identifica por <b>CPF</b> (ou nome) quem já está cadastrado, mostra a relação em vermelho e gera o
    <b>Arquivo Domínio LIMPO</b> (sem duplicados) para seguir na importação.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        up_dom_dup = st.file_uploader(
            "📄 Arquivo Domínio (.xls)", type=["xls", "xlsx"], key="dup_dom"
        )
    with c2:
        up_sis_dup = st.file_uploader(
            "🏢 Colaboradores do Sistema Interno (.xls)", type=["xls", "xlsx"], key="dup_sis"
        )

    if st.button("🔍 Verificar duplicados", type="primary", key="btn_dup"):
        if up_dom_dup is None or up_sis_dup is None:
            st.warning("Envie os dois arquivos para comparar.")
        else:
            with st.spinner("Comparando bases..."):
                try:
                    df_dom = read_excel_smart(up_dom_dup, dtype=str)
                    df_sis = read_excel_smart(up_sis_dup, dtype=str)
                    dupes, limpo, stats = comparar_colaboradores(df_dom, df_sis)

                    ss["df_dom_original"] = df_dom
                    ss["df_dom_limpo"] = limpo

                    put_art(
                        "dom_limpo_xlsx",
                        df_to_xlsx_bytes(limpo),
                        "Arquivo Dominio_LIMPO.xlsx",
                        XLSX_MIME,
                    )
                    put_art(
                        "dom_limpo_csv",
                        limpo.to_csv(index=False).encode("utf-8-sig"),
                        "Arquivo Dominio_LIMPO.csv",
                        "text/csv",
                    )
                    if not dupes.empty:
                        put_art(
                            "dupes_xlsx",
                            df_to_xlsx_bytes(dupes, red_all=True),
                            "Relacao_Duplicados.xlsx",
                            XLSX_MIME,
                        )
                    ss["dup_stats"] = stats
                    ss["df_dupes"] = dupes
                except Exception as e:
                    st.error(f"Erro ao comparar: {e}")

    if ss.get("dup_stats"):
        stats = ss["dup_stats"]
        metricas(
            [
                ("No Domínio", stats["total_dominio"]),
                ("Já cadastrados", stats["duplicados"]),
                ("Novos (irão importar)", stats["novos"]),
                ("Base do sistema", stats["total_sistema"]),
            ]
        )

        dupes = ss.get("df_dupes")
        if dupes is not None and not dupes.empty:
            st.markdown(
                '<div class="legend"><span class="box"></span> Estes colaboradores JÁ existem no sistema interno — foram REMOVIDOS do arquivo limpo</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(
                dupes.style.set_properties(
                    **{"background-color": "#7F1D1D", "color": "#FECACA"}
                ),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("✅ Nenhum duplicado encontrado — todos os colaboradores são novos.")

        d1, d2, d3 = st.columns(3)
        with d1:
            dl("dom_limpo_xlsx", "⬇️ Domínio LIMPO (.xlsx)")
        with d2:
            dl("dom_limpo_csv", "⬇️ Domínio LIMPO (.csv)")
        with d3:
            dl("dupes_xlsx", "⬇️ Relação de Duplicados (.xlsx)")

        if ss["df_dom_limpo"] is not None:
            st.info(
                "💡 O **Domínio LIMPO** já está em memória: na aba **1️⃣ Importar Layout** e no **⭐ Executar Tudo** "
                "você pode usá-lo direto, sem reenviar arquivo."
            )

# ──────────────────────────────────────────────────────────────────────────
# ABA 1 — IMPORTAR LAYOUT
# ──────────────────────────────────────────────────────────────────────────
with tab_layout:
    st.markdown(
        """
<div class="card">
  <div class="card-title">1️⃣ Geração do Importar Layout (AppLider)</div>
  <div class="card-sub">
    Domínio + Mapeamento do Sistema → <b>Importar Layout.xlsx</b>.
    Células sem <b>Posto de Serviço</b>, <b>Função</b>, <b>Escala</b> ou <b>Horário</b> ficam <b style="color:#FF8080">VERMELHAS</b> na planilha e no preview.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    usar_limpo = False
    if ss["df_dom_limpo"] is not None:
        usar_limpo = st.checkbox(
            f"🧹 Usar Domínio LIMPO da aba Duplicados ({len(ss['df_dom_limpo'])} registros)",
            value=True,
            key="chk_usar_limpo",
        )

    up_dom_layout = None
    if not usar_limpo:
        up_dom_layout = st.file_uploader(
            "📄 Arquivo Domínio (.xls)", type=["xls", "xlsx"], key="layout_dom"
        )

    if st.button("🚀 Gerar Importar Layout", type="primary", key="btn_layout"):
        base_mapa = get_base_bytes(up_mapa, "Mapeamento Sistema.xls")
        base_modelo = get_base_bytes(up_modelo, "Colunas Originais.xlsx")

        if base_mapa is None or base_modelo is None:
            st.error(
                "Faltam arquivos base: envie **Mapeamento Sistema.xls** e **Colunas Originais.xlsx** na barra lateral."
            )
        elif not usar_limpo and up_dom_layout is None:
            st.warning("Envie o Arquivo Domínio.")
        else:
            with st.spinner("Processando layout..."):
                try:
                    df_dom = (
                        ss["df_dom_limpo"].copy()
                        if usar_limpo
                        else read_excel_smart(up_dom_layout, dtype=str)
                    )
                    mapas = carregar_todos_os_mapas(io.BytesIO(base_mapa))
                    colunas = list(pd.read_excel(io.BytesIO(base_modelo), nrows=0).columns)

                    df_layout, pendencias = processar_layout(df_dom, mapas, colunas)
                    ss["df_layout"] = df_layout
                    ss["layout_pend"] = pendencias

                    xlsx = df_to_xlsx_bytes(df_layout, red_cells=pendencias)
                    put_art("layout_xlsx", xlsx, "Importar Layout.xlsx", XLSX_MIME)
                    put_art(
                        "layout_csv",
                        xlsx_bytes_to_csv_bytes(xlsx),
                        "Importar Layout.csv",
                        "text/csv",
                    )
                    st.success(f"✅ Layout gerado: {len(df_layout)} colaborador(es).")
                except Exception as e:
                    st.error(f"Erro ao gerar layout: {e}")

    if ss["df_layout"] is not None:
        df = ss["df_layout"]
        pend = ss.get("layout_pend", [])

        faltas = {c: 0 for c in CAMPOS_OBRIGATORIOS}
        for _, c in pend:
            faltas[c] = faltas.get(c, 0) + 1

        metricas(
            [
                ("Colaboradores", len(df)),
                ("Sem Posto", faltas.get("postotrabalho_id", 0)),
                ("Sem Função", faltas.get("funcao_id", 0)),
                ("Sem Escala", faltas.get("escalatrabalho_id", 0)),
                ("Sem Horário", faltas.get("horariotrabalho_id", 0)),
            ]
        )

        if pend:
            st.markdown(
                '<div class="legend"><span class="box"></span> Células vermelhas = informação faltando (Posto / Função / Escala / Horário). Corrija antes de importar.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.success("✅ Nenhuma pendência em Posto, Função, Escala ou Horário.")

        preview_cols = [
            c
            for c in [
                "nome",
                "matricula",
                "cpf",
                "nomepostotrabalho",
                "postotrabalho_id",
                "nomefuncao",
                "funcao_id",
                "tipoescala",
                "escalatrabalho_id",
                "nomehorariotrabalho",
                "horariotrabalho_id",
            ]
            if c in df.columns
        ]
        st.dataframe(
            style_missing(df[preview_cols].head(150)),
            use_container_width=True,
            hide_index=True,
        )

        d1, d2 = st.columns(2)
        with d1:
            dl("layout_xlsx", "⬇️ Importar Layout.xlsx (com vermelhos)")
        with d2:
            dl("layout_csv", "⬇️ Importar Layout.csv")

# ──────────────────────────────────────────────────────────────────────────
# ABA 2 — BENEFÍCIOS
# ──────────────────────────────────────────────────────────────────────────
with tab_benef:
    st.markdown(
        """
<div class="card">
  <div class="card-title">2️⃣ Geração do Importar Benefícios</div>
  <div class="card-sub">
    Usa o <b>Importar Layout</b> (desta sessão ou enviado) + a <b>Relação de Benefícios</b>.
    Você escolhe o <b>colaborador_id</b>: matrícula eSocial (padrão) ou <b>IDs em sequência</b> a partir do último ID do EasyApp.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    df_func = ss["df_layout"]
    if df_func is None:
        up_layout_benef = st.file_uploader(
            "📄 Importar Layout.xlsx (gerado na etapa 1)",
            type=["xlsx"],
            key="benef_layout",
        )
        if up_layout_benef is not None:
            df_func = pd.read_excel(up_layout_benef)
    else:
        st.caption(f"📎 Usando o Importar Layout desta sessão ({len(df_func)} colaboradores).")

    modo = st.radio(
        "Como preencher o colaborador_id?",
        ["Matrícula eSocial (comportamento original)", "IDs sequenciais a partir do último ID EasyApp"],
        key="modo_benef",
    )

    confirmado = True
    if modo.startswith("IDs sequenciais"):
        st.info(
            f"🔢 Último ID salvo: **{ss['ultimo_id']}** → o primeiro novo colaborador receberá **{ss['ultimo_id'] + 1}** "
            "(altere na barra lateral, se preciso)."
        )
        confirmado = st.checkbox(
            f"✅ Confirmo que **{ss['ultimo_id']}** é o último ID cadastrado no EasyApp",
            key="chk_confirma_id",
        )

    if st.button("🎁 Gerar Benefícios", type="primary", key="btn_benef"):
        base_benef = get_base_bytes(up_benef, "BENEFICIOS - RELACAO.xlsx")
        if df_func is None:
            st.warning("Gere o Importar Layout na etapa 1 (ou envie o arquivo acima).")
        elif base_benef is None:
            st.error("Falta a **Relação de Benefícios** (envie na barra lateral).")
        elif not confirmado:
            st.warning("Confirme o último ID antes de gerar em sequência.")
        else:
            with st.spinner("Calculando benefícios..."):
                try:
                    df_benef = pd.read_excel(io.BytesIO(base_benef), sheet_name="BENEFICIOS")
                    modo_key = "sequencial" if modo.startswith("IDs") else "esocial"
                    df_saida, mapa_ids, stats = processar_beneficios_df(
                        df_func, df_benef, modo=modo_key, ultimo_id=ss["ultimo_id"]
                    )

                    xlsx = beneficios_xlsx_bytes(df_saida)
                    put_art("benef_xlsx", xlsx, "Importar Beneficios.xlsx", XLSX_MIME)
                    put_art(
                        "benef_csv",
                        xlsx_bytes_to_csv_bytes(xlsx),
                        "Importar Beneficios.csv",
                        "text/csv",
                    )
                    ss["mapa_ids"] = mapa_ids
                    ss["benef_stats"] = stats

                    if modo_key == "sequencial" and mapa_ids is not None and len(mapa_ids):
                        ss["ultimo_id"] = int(mapa_ids["novo_id"].max())

                    st.success(
                        f"✅ Benefícios gerados: {stats['linhas']} linha(s) para {stats['colaboradores']} colaborador(es)."
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar benefícios: {e}")

    if ss.get("benef_stats"):
        stats = ss["benef_stats"]
        metricas(
            [
                ("Colaboradores", stats["colaboradores"]),
                ("Linhas de benefício", stats["linhas"]),
                ("Sem benefício", stats["sem_beneficio"]),
            ]
        )
        if stats["sem_beneficio"]:
            st.warning(
                f"⚠️ {stats['sem_beneficio']} linha(s) ficaram sem benefício (regra não encontrada)."
            )

        mapa_ids = ss.get("mapa_ids")
        if mapa_ids is not None and len(mapa_ids):
            st.markdown(
                '<div class="card-sub">🔢 <b>Mapa de IDs gerados</b> — novo ID × matrícula eSocial × nome:</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(mapa_ids, use_container_width=True, hide_index=True)

        st.markdown(
            '<div class="legend legend-blue"><span class="box"></span> No Excel: <b>azul</b> = colaborador com mais de um benefício &nbsp;•&nbsp; <b>vermelho</b> = colaborador com um único benefício (regra original)</div>',
            unsafe_allow_html=True,
        )

        d1, d2 = st.columns(2)
        with d1:
            dl("benef_xlsx", "⬇️ Importar Beneficios.xlsx")
        with d2:
            dl("benef_csv", "⬇️ Importar Beneficios.csv")

# ──────────────────────────────────────────────────────────────────────────
# ABA 3 — CONVERTER CSV
# ──────────────────────────────────────────────────────────────────────────
with tab_csv:
    st.markdown(
        """
<div class="card">
  <div class="card-title">3️⃣ Conversão para CSV</div>
  <div class="card-sub">
    Converte qualquer .xlsx em CSV <b>preservando zeros à esquerda</b> (CPF, PIS, matrícula, CEP) — mesma regra do conversor original.
    Os CSVs das etapas 1 e 2 já são gerados automaticamente e ficam prontos aqui para baixar quando quiser.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="card-sub">📦 CSVs já gerados nesta sessão:</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        ok1 = dl("layout_csv", "⬇️ Importar Layout.csv")
    with c2:
        ok2 = dl("benef_csv", "⬇️ Importar Beneficios.csv")
    if not ok1 and not ok2:
        st.caption("Nenhum CSV gerado ainda — rode as etapas 1/2 ou use o conversor avulso abaixo.")

    st.markdown("---")
    st.markdown('<div class="card-sub">🔄 Conversor avulso (qualquer .xlsx):</div>', unsafe_allow_html=True)
    up_any = st.file_uploader("Enviar .xlsx para converter", type=["xlsx"], key="any_xlsx")
    if up_any is not None and st.button("Converter para CSV", key="btn_any_csv"):
        with st.spinner("Convertendo..."):
            csv_bytes = xlsx_bytes_to_csv_bytes(up_any.getvalue())
            nome = Path(up_any.name).stem + ".csv"
            put_art("custom_csv", csv_bytes, nome, "text/csv")
        st.success(f"✅ {nome} pronto.")
    dl("custom_csv", "⬇️ Baixar CSV convertido")

# ──────────────────────────────────────────────────────────────────────────
# ABA 4 — EXECUTAR TUDO
# ──────────────────────────────────────────────────────────────────────────
with tab_all:
    st.markdown(
        """
<div class="card">
  <div class="card-title">⭐ Pipeline completo: Domínio → (sem duplicados) → Layout → Benefícios → CSVs</div>
  <div class="card-sub">
    Roda tudo de uma vez. Ao final, <b>todos os arquivos ficam prontos na Central de Downloads</b>:
    baixe um por um, na ordem que quiser, <b>sem refazer o processamento</b>.
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        up_dom_all = st.file_uploader(
            "📄 Arquivo Domínio (.xls)", type=["xls", "xlsx"], key="all_dom"
        )
    with c2:
        up_sis_all = st.file_uploader(
            "🏢 Sistema interno (opcional — remove duplicados)",
            type=["xls", "xlsx"],
            key="all_sis",
        )

    modo_all = st.radio(
        "colaborador_id nos benefícios:",
        ["Matrícula eSocial", "IDs sequenciais (último ID da barra lateral)"],
        key="modo_all",
        horizontal=True,
    )
    confirma_all = True
    if modo_all.startswith("IDs"):
        confirma_all = st.checkbox(
            f"✅ Confirmo que **{ss['ultimo_id']}** é o último ID do EasyApp",
            key="chk_confirma_all",
        )

    if st.button("▶️ EXECUTAR TUDO", type="primary", key="btn_all"):
        base_mapa = get_base_bytes(up_mapa, "Mapeamento Sistema.xls")
        base_modelo = get_base_bytes(up_modelo, "Colunas Originais.xlsx")
        base_benef = get_base_bytes(up_benef, "BENEFICIOS - RELACAO.xlsx")

        if up_dom_all is None:
            st.warning("Envie o Arquivo Domínio.")
        elif base_mapa is None or base_modelo is None or base_benef is None:
            st.error("Faltam arquivos base na barra lateral (Mapeamento / Modelo / Benefícios).")
        elif not confirma_all:
            st.warning("Confirme o último ID antes de rodar com IDs sequenciais.")
        else:
            with st.status("Executando pipeline completo...", expanded=True) as status:
                try:
                    # 0) Domínio
                    st.write("📄 Lendo Arquivo Domínio...")
                    df_dom = read_excel_smart(up_dom_all, dtype=str)

                    # 0.5) Duplicados (opcional)
                    if up_sis_all is not None:
                        st.write("🧹 Removendo duplicados com base no sistema interno...")
                        df_sis = read_excel_smart(up_sis_all, dtype=str)
                        dupes, limpo, stats_dup = comparar_colaboradores(df_dom, df_sis)
                        ss["df_dom_limpo"] = limpo
                        ss["dup_stats"] = stats_dup
                        ss["df_dupes"] = dupes
                        put_art(
                            "dom_limpo_xlsx",
                            df_to_xlsx_bytes(limpo),
                            "Arquivo Dominio_LIMPO.xlsx",
                            XLSX_MIME,
                        )
                        if not dupes.empty:
                            put_art(
                                "dupes_xlsx",
                                df_to_xlsx_bytes(dupes, red_all=True),
                                "Relacao_Duplicados.xlsx",
                                XLSX_MIME,
                            )
                        st.write(f"   → {stats_dup['duplicados']} duplicado(s) removido(s), {stats_dup['novos']} novo(s).")
                        df_dom = limpo

                    # 1) Layout
                    st.write("1️⃣ Gerando Importar Layout...")
                    mapas = carregar_todos_os_mapas(io.BytesIO(base_mapa))
                    colunas = list(pd.read_excel(io.BytesIO(base_modelo), nrows=0).columns)
                    df_layout, pendencias = processar_layout(df_dom, mapas, colunas)
                    ss["df_layout"] = df_layout
                    ss["layout_pend"] = pendencias
                    xlsx_layout = df_to_xlsx_bytes(df_layout, red_cells=pendencias)
                    put_art("layout_xlsx", xlsx_layout, "Importar Layout.xlsx", XLSX_MIME)
                    put_art(
                        "layout_csv",
                        xlsx_bytes_to_csv_bytes(xlsx_layout),
                        "Importar Layout.csv",
                        "text/csv",
                    )
                    st.write(f"   → {len(df_layout)} colaborador(es), {len(pendencias)} célula(s) em vermelho.")

                    # 2) Benefícios
                    st.write("2️⃣ Gerando Importar Benefícios...")
                    df_benef = pd.read_excel(io.BytesIO(base_benef), sheet_name="BENEFICIOS")
                    modo_key = "sequencial" if modo_all.startswith("IDs") else "esocial"
                    df_saida, mapa_ids, stats_b = processar_beneficios_df(
                        df_layout, df_benef, modo=modo_key, ultimo_id=ss["ultimo_id"]
                    )
                    xlsx_benef = beneficios_xlsx_bytes(df_saida)
                    put_art("benef_xlsx", xlsx_benef, "Importar Beneficios.xlsx", XLSX_MIME)
                    put_art(
                        "benef_csv",
                        xlsx_bytes_to_csv_bytes(xlsx_benef),
                        "Importar Beneficios.csv",
                        "text/csv",
                    )
                    ss["mapa_ids"] = mapa_ids
                    ss["benef_stats"] = stats_b
                    if modo_key == "sequencial" and mapa_ids is not None and len(mapa_ids):
                        ss["ultimo_id"] = int(mapa_ids["novo_id"].max())
                    st.write(f"   → {stats_b['linhas']} linha(s) de benefício.")

                    # 3) CSVs já gerados acima
                    st.write("3️⃣ CSVs gerados (texto preservado).")
                    status.update(label="✅ Pipeline concluído!", state="complete")
                except Exception as e:
                    status.update(label="❌ Falha no pipeline", state="error")
                    st.error(f"Erro: {e}")

    # Central de downloads — tudo que existe na sessão
    arts = ss["artifacts"]
    if arts:
        st.markdown("---")
        st.markdown(
            '<div class="card-title">📦 Central de Downloads</div>'
            '<div class="card-sub">Clique à vontade — os arquivos já estão prontos, baixar NÃO refaz o processamento.</div>',
            unsafe_allow_html=True,
        )
        grid = [
            ("dom_limpo_xlsx", "⬇️ Domínio LIMPO (.xlsx)"),
            ("dupes_xlsx", "⬇️ Relação Duplicados (.xlsx)"),
            ("layout_xlsx", "⬇️ Importar Layout.xlsx"),
            ("layout_csv", "⬇️ Importar Layout.csv"),
            ("benef_xlsx", "⬇️ Importar Beneficios.xlsx"),
            ("benef_csv", "⬇️ Importar Beneficios.csv"),
        ]
        cols = st.columns(3)
        for i, (k, label) in enumerate(grid):
            with cols[i % 3]:
                dl(k, label)

st.markdown(
    "<p style='text-align:center;color:#5C6B93;font-size:.8rem;margin-top:24px'>"
    "Importador Domínio → AppLider • Líder Limpe • todas as regras do sistema original preservadas"
    "</p>",
    unsafe_allow_html=True,
)
