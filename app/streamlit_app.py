"""
╔══════════════════════════════════════════════════════════════════════════╗
║  STREAMLIT_APP.PY — LiderLimp Importador (v2 — set/2026)                 ║
╠══════════════════════════════════════════════════════════════════════════╣
║  Abas:                                                                   ║
║    1️⃣  Importar Layout  — Domínio → planilha AppLider                   ║
║    2️⃣  Benefícios       — motor de regras de VA/CB/CF/VT                ║
║    3️⃣  Converter CSV    — xlsx → csv preservando texto                  ║
║    ⭐  Executar Tudo    — fluxo completo de uma vez                     ║
║    ❓  Ajuda                                                             ║
╠══════════════════════════════════════════════════════════════════════════╣
║  NOVIDADES v2:                                                           ║
║  • "Colunas Originais" NÃO é mais necessária: a ordem das colunas        ║
║    do layout é fixa no código (importar_dados.LAYOUT_COLUNAS).           ║
║  • Mapeamento e Benefícios ficam SALVOS na pasta `data/` depois do       ║
║    primeiro upload — não precisa reenviar a cada sessão.                 ║
║    Para trocar, marque "Enviar nova planilha" na barra lateral.          ║
╚══════════════════════════════════════════════════════════════════════════╝
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from core.theme import aplicar_tema, render_header, card_inicio, card_fim
from core.importar_dados import (
    processar_layout, aplicar_destaque_vermelho, CAMPOS_CRITICOS_VERMELHO,
)
from core.beneficios import (
    processar_beneficios, aplicar_cores_colaborador_id,
)
from core.state import get_ultimo_id, set_ultimo_id
from core.utils import xlsx_bytes_para_csv_bytes, dataframe_para_xlsx_bytes
from core.config import carregar_regras
from core.duplicates import comparar_colaboradores
from core.gsheets import carregar_beneficios_com_fallback, carregar_mapeamento_com_fallback


# ══════════════════════════════════════════════════════════════════════════
# ARQUIVOS DE CONFIGURAÇÃO (pasta data/)
# ══════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

ARQ_MAPEAMENTO_PADRAO = DATA_DIR / "Mapeamento Sistema.xls"
ARQ_BENEFICIOS_PADRAO = DATA_DIR / "BENEFICIOS - RELACAO.xlsx"


def _salvar_config(uploaded, destino: Path) -> None:
    """Grava a planilha enviada na pasta data/ (persiste entre sessões)."""
    destino.write_bytes(uploaded.read())


# ══════════════════════════════════════════════════════════════════════════
# SETUP DA PÁGINA
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="LiderLimp · Importador",
    page_icon=str(BASE_DIR / "app" / "assets" / "logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_tema()
render_header()


# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR — arquivos de configuração (persistem em data/)
# ══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.subheader("🗂️ Arquivos de configuração")
    st.caption(
        "Ficam **salvos** na pasta `data/` depois do primeiro envio — "
        "você só precisa enviar de novo quando quiser **atualizar**."
    )

    # ── Mapeamento Sistema — Google Sheets (padrão) ou upload/local ──
    st.markdown("**Mapeamento Sistema**")
    usar_gsheets_map = st.checkbox(
        "☁️ Mapeamento via Google Sheets",
        value=True, key="chk_gsheets_map",
        help="Lê o MAPEAMENTO SISTEMA direto do Google Sheets. "
             "Se a conexão falhar, usa automaticamente o arquivo local/upload."
    )
    f_map = None
    if not usar_gsheets_map:
        if ARQ_MAPEAMENTO_PADRAO.exists():
            st.caption(f"💾 Local: {ARQ_MAPEAMENTO_PADRAO.name}")
        f_map = st.file_uploader(
            "Mapeamento Sistema (.xls)", type=["xls"], key="up_mapeamento")

    st.divider()

    # ── Benefícios Relação — Google Sheets (padrão) ou upload manual ──
    st.markdown("**Benefícios Relação**")
    usar_gsheets = st.checkbox(
        "☁️ Usar Google Sheets (sempre atualizado)",
        value=True, key="chk_gsheets",
        help="Lê a planilha BENEFICIOS - RELAÇÃO direto do Google Sheets. "
             "Se a conexão falhar, o app usa automaticamente o arquivo local "
             "ou o upload abaixo."
    )
    f_benef = None
    if not usar_gsheets:
        if ARQ_BENEFICIOS_PADRAO.exists():
            st.caption(f"💾 Local: {ARQ_BENEFICIOS_PADRAO.name}")
        f_benef = st.file_uploader(
            "BENEFICIOS - RELACAO (.xlsx)", type=["xlsx"], key="up_benef")

    st.divider()
    st.caption(
        "ℹ️ A planilha **Colunas Originais** não é mais usada: "
        "a ordem das colunas do layout agora é fixa no código."
    )


def _ler_config(uploaded, padrao_path: Path) -> bytes | None:
    """
    Prioridade: (1) arquivo recém-enviado → já salva em data/ e usa;
    (2) arquivo salvo em data/.
    """
    if uploaded is not None:
        dados = uploaded.read()
        padrao_path.write_bytes(dados)
        return dados
    if padrao_path.exists():
        return padrao_path.read_bytes()
    return None


# ══════════════════════════════════════════════════════════════════════════
# CARGA DE CONFIGURAÇÕES (Google Sheets → upload → arquivo local)
# ══════════════════════════════════════════════════════════════════════════
def _carregar_beneficios_e_regras():
    """
    Devolve (df_benef, regras, origem).
    Estratégia: Google Sheets (se habilitado) → upload da sessão → data/.
    """
    upload_bytes = f_benef.read() if f_benef is not None else None
    df_benef, df_listas, origem = carregar_beneficios_com_fallback(
        arquivo_local=ARQ_BENEFICIOS_PADRAO,
        beneficios_relacao_bytes=upload_bytes,
    )
    regras = carregar_regras(df_listas)
    return df_benef, regras, origem


def _carregar_mapeamento():
    """
    Devolve (fonte_mapeamento, origem).
    Estratégia: Google Sheets (se habilitado) → upload da sessão → data/.
    `fonte_mapeamento` pode ser dict de DataFrames (Google), bytes ou Path.
    """
    upload_bytes = f_map.read() if f_map is not None else None
    fonte, origem = carregar_mapeamento_com_fallback(
        arquivo_local=ARQ_MAPEAMENTO_PADRAO,
        mapeamento_bytes=upload_bytes,
    )
    return fonte, origem


# ══════════════════════════════════════════════════════════════════════════
# ABAS PRINCIPAIS
# ══════════════════════════════════════════════════════════════════════════
tab_layout, tab_benef, tab_csv, tab_tudo, tab_dup, tab_help = st.tabs(
    ["1️⃣  Importar Layout", "2️⃣  Benefícios", "3️⃣  Converter para CSV",
     "⭐  Executar Tudo", "❓ Ajuda", "🧹  Duplicados"]
)


# ══════════════════════════════════════════════════════════════════════════
# ABA 1 — IMPORTAR LAYOUT
# ══════════════════════════════════════════════════════════════════════════
with tab_layout:
    card_inicio("1) Gerar planilha de importação do AppLider (Importar Layout)")
    st.write(
        "Envie o **Arquivo Domínio (.xls)** exportado do sistema Domínio. "
        "O app cruza com os mapeamentos e gera a planilha de importação do AppLider. "
        "Campos críticos faltantes (**posto**, **função**, **escala**) ficam **em vermelho**."
    )

    f_dominio = st.file_uploader(
        "📥  Arquivo Domínio (.xls)", type=["xls"], key="up_dominio_t1"
    )

    if "dominio_limpo_xlsx" in st.session_state:
        if st.checkbox("🧹 Usar o Domínio LIMPO da aba Duplicados (sem duplicados)", key="chk_limpo_tab1"):
            f_dominio = io.BytesIO(st.session_state["dominio_limpo_xlsx"])
            f_dominio.name = "dominio_limpo.xlsx"

    if st.button("🚀  Gerar Importar Layout", type="primary", key="btn_layout"):
        if not f_dominio:
            st.error("Envie o Arquivo Domínio.")
        else:
            try:
                map_fonte, origem_map = _carregar_mapeamento()
                _, regras_layout, origem_benef = _carregar_beneficios_e_regras()
            except Exception as e:
                st.error(f"Falha ao carregar configurações: {e}")
                st.stop()
            st.caption(f"Mapeamento: {origem_map} · Benefícios: {origem_benef}")
            with st.spinner("Processando…"):
                df, problemas = processar_layout(
                    f_dominio.read(), map_fonte, regras=regras_layout
                )
                xlsx_bytes = dataframe_para_xlsx_bytes(df)
                xlsx_bytes = aplicar_destaque_vermelho(xlsx_bytes, problemas)

                st.session_state["layout_df"]   = df
                st.session_state["layout_xlsx"] = xlsx_bytes
                st.session_state["layout_problemas"] = problemas

                col1, col2, col3 = st.columns(3)
                col1.metric("Colaboradores", len(df))
                col2.metric("Com pendência", len(problemas))
                col3.metric("Campos críticos monitorados", len(CAMPOS_CRITICOS_VERMELHO))

                if problemas:
                    st.warning(
                        "Algumas linhas têm campos críticos vazios "
                        "(marcados em vermelho na planilha)."
                    )
                    st.dataframe(pd.DataFrame(problemas), use_container_width=True)
                else:
                    st.success("Nenhuma pendência encontrada! 🎉")

    if "layout_xlsx" in st.session_state:
        st.download_button(
            "⬇️  Baixar Importar Layout.xlsx",
            data=st.session_state["layout_xlsx"],
            file_name="Importar Layout.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.dataframe(st.session_state["layout_df"].head(20), use_container_width=True)

    card_fim()


# ══════════════════════════════════════════════════════════════════════════
# ABA 2 — BENEFÍCIOS
# ══════════════════════════════════════════════════════════════════════════
with tab_benef:
    card_inicio("2) Gerar planilha de Benefícios (com ID sequencial opcional)")

    st.write(
        "Você pode usar:\n\n"
        "- a planilha **Importar Layout** que acabou de gerar (em memória), **ou**\n"
        "- uma planilha **Importar Layout.xlsx** já existente que você suba abaixo."
    )

    fonte = st.radio(
        "Fonte da planilha de funcionários:",
        ["Usar a Importar Layout gerada agora", "Subir uma nova"],
        horizontal=True, key="fonte_benef"
    )

    df_func = None
    if fonte == "Usar a Importar Layout gerada agora":
        if "layout_df" in st.session_state:
            df_func = st.session_state["layout_df"].copy()
            st.info(f"Usando {len(df_func)} colaboradores da aba 1.")
        else:
            st.warning("Você ainda não gerou a Importar Layout. Gere na aba 1 ou suba uma.")
    else:
        up = st.file_uploader(
            "Importar Layout.xlsx", type=["xlsx"], key="up_layout_t2")
        if up:
            df_func = pd.read_excel(up, dtype=str)
            st.info(f"Planilha carregada com {len(df_func)} linhas.")

    st.markdown("##### ID sequencial do EasyApp")
    ultimo_salvo = get_ultimo_id()
    col_a, col_b = st.columns([2, 1])
    with col_a:
        usar_seq = st.checkbox(
            "Gerar `colaborador_id` em sequência a partir do último ID do EasyApp",
            value=True if ultimo_salvo else False,
        )
    with col_b:
        ultimo_id_input = st.number_input(
            "Último ID cadastrado no EasyApp",
            min_value=0, step=1,
            value=int(ultimo_salvo) if ultimo_salvo else 0,
        )

    confirma = False
    if usar_seq:
        confirma = st.checkbox(
            f"✅ Confirmo que o último ID cadastrado é **{ultimo_id_input}**",
            value=False,
        )

    if st.button("🎁  Gerar planilha de Benefícios", type="primary", key="btn_benef"):
        if df_func is None or len(df_func) == 0:
            st.error("Não há colaboradores para processar.")
        elif usar_seq and not confirma:
            st.error("Confirme o último ID antes de gerar.")
        else:
            try:
                df_benef_cfg, regras_benef, origem_benef = _carregar_beneficios_e_regras()
            except Exception as e:
                st.error(f"Falha ao carregar benefícios: {e}")
                st.stop()
            st.caption(f"Fonte dos benefícios: {origem_benef} · Listas: {regras_benef.origem}")
            with st.spinner("Gerando benefícios…"):
                df_saida, erros = processar_beneficios(
                    df_func,
                    df_benef_pronto=df_benef_cfg,
                    regras=regras_benef,
                    ultimo_id_easyapp=ultimo_id_input if usar_seq else None,
                    usar_id_sequencial=usar_seq,
                )
                xlsx_bytes = dataframe_para_xlsx_bytes(df_saida)
                xlsx_bytes = aplicar_cores_colaborador_id(xlsx_bytes)

                # salva último ID atualizado
                if usar_seq:
                    novo_ultimo = int(ultimo_id_input) + len(df_func)
                    set_ultimo_id(novo_ultimo)
                    st.success(f"Último ID atualizado para **{novo_ultimo}** "
                               f"(salvo em `state.json`).")

                st.session_state["benef_df"] = df_saida
                st.session_state["benef_xlsx"] = xlsx_bytes

                col1, col2, col3 = st.columns(3)
                col1.metric("Linhas geradas", len(df_saida))
                col2.metric("Sem benefício", sum(df_saida["beneficio_id"].isna()))
                col3.metric("Colaboradores", df_saida["parceiro_id"].nunique())

                if erros:
                    st.warning("Colaboradores sem benefício identificado:")
                    st.dataframe(pd.DataFrame(erros), use_container_width=True)

    if "benef_xlsx" in st.session_state:
        st.download_button(
            "⬇️  Baixar Importar Beneficios.xlsx",
            data=st.session_state["benef_xlsx"],
            file_name="Importar Beneficios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.dataframe(st.session_state["benef_df"].head(20), use_container_width=True)

    card_fim()


# ══════════════════════════════════════════════════════════════════════════
# ABA 3 — CONVERTER PARA CSV
# ══════════════════════════════════════════════════════════════════════════
with tab_csv:
    card_inicio("3) Converter planilhas .xlsx para CSV (preservando texto)")
    st.write(
        "Converta para CSV qualquer arquivo `.xlsx` (mantém zeros à esquerda, datas, etc)."
    )

    fonte_csv = st.radio(
        "Origem do arquivo:",
        ["Usar arquivos gerados nesta sessão", "Subir um .xlsx"],
        horizontal=True, key="fonte_csv"
    )

    if fonte_csv == "Usar arquivos gerados nesta sessão":
        c1, c2 = st.columns(2)
        with c1:
            if "layout_xlsx" in st.session_state:
                csv_bytes = xlsx_bytes_para_csv_bytes(st.session_state["layout_xlsx"])
                st.download_button(
                    "⬇️  Importar Layout.csv",
                    data=csv_bytes,
                    file_name="Importar Layout.csv",
                    mime="text/csv", type="primary", use_container_width=True,
                )
            else:
                st.caption("Gere a Importar Layout (aba 1) para liberar o CSV.")
        with c2:
            if "benef_xlsx" in st.session_state:
                csv_bytes = xlsx_bytes_para_csv_bytes(st.session_state["benef_xlsx"])
                st.download_button(
                    "⬇️  Importar Beneficios.csv",
                    data=csv_bytes,
                    file_name="Importar Beneficios.csv",
                    mime="text/csv", type="primary", use_container_width=True,
                )
            else:
                st.caption("Gere a planilha de benefícios (aba 2) para liberar o CSV.")
    else:
        up = st.file_uploader("Arquivo .xlsx", type=["xlsx"], key="up_csv")
        if up:
            csv_bytes = xlsx_bytes_para_csv_bytes(up.read())
            st.download_button(
                "⬇️  Baixar CSV", data=csv_bytes,
                file_name=Path(up.name).with_suffix(".csv").name,
                mime="text/csv", type="primary", use_container_width=True,
            )

    card_fim()


# ══════════════════════════════════════════════════════════════════════════
# ABA 4 — EXECUTAR TUDO
# ══════════════════════════════════════════════════════════════════════════
with tab_tudo:
    # Downloads persistentes: ficam disponíveis mesmo após clicar/baixar
    if st.session_state.get("tudo_pronto"):
        st.markdown("#### 📥 Última geração (continua disponível)")
        p1, p2 = st.columns(2)
        p1.download_button("⬇️ Importar Layout.xlsx", data=st.session_state["layout_xlsx"],
                           file_name="Importar Layout.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_tudo_layout")
        p2.download_button("⬇️ Importar Beneficios.xlsx", data=st.session_state["benef_xlsx"],
                           file_name="Importar Beneficios.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_tudo_benef")
        st.caption("Os CSVs estão na aba 3. Para gerar novamente, use o botão abaixo.")
        st.divider()

    card_inicio("⭐ Executar todo o fluxo de uma vez")
    st.write("Sobe o Arquivo Domínio, confirma o último ID e o app gera **todos os arquivos**.")

    f_dom_all = st.file_uploader(
        "📥  Arquivo Domínio (.xls)", type=["xls"], key="up_dominio_all"
    )

    ultimo_salvo2 = get_ultimo_id()
    c1, c2 = st.columns([2, 1])
    with c1:
        usar_seq2 = st.checkbox(
            "Gerar colaborador_id sequencial (recomendado)",
            value=True, key="seq_tudo"
        )
    with c2:
        ultimo_id_all = st.number_input(
            "Último ID EasyApp", min_value=0, step=1,
            value=int(ultimo_salvo2) if ultimo_salvo2 else 0,
            key="ult_tudo"
        )
    confirma2 = st.checkbox(
        f"✅ Confirmo que o último ID é **{ultimo_id_all}**",
        value=False, key="conf_tudo",
    ) if usar_seq2 else True

    if "dominio_limpo_xlsx" in st.session_state:
        if st.checkbox("🧹 Usar o Domínio LIMPO da aba Duplicados (sem duplicados)", key="chk_limpo_tudo"):
            f_dom_all = io.BytesIO(st.session_state["dominio_limpo_xlsx"])
            f_dom_all.name = "dominio_limpo.xlsx"

    if st.button("⭐  EXECUTAR TUDO", type="primary", key="btn_tudo"):
        if not f_dom_all:
            st.error("Envie o Arquivo Domínio.")
        elif usar_seq2 and not confirma2:
            st.error("Confirme o último ID.")
        else:
            try:
                map_fonte, origem_map = _carregar_mapeamento()
                df_benef_cfg, regras_all, origem_benef = _carregar_beneficios_e_regras()
            except Exception as e:
                st.error(f"Falha ao carregar configurações: {e}")
                st.stop()
            st.caption(f"Mapeamento: {origem_map} · Benefícios: {origem_benef} · Listas: {regras_all.origem}")
            with st.spinner("Executando o fluxo completo…"):
                dom_bytes = f_dom_all.read()
                df, problemas = processar_layout(dom_bytes, map_fonte, regras=regras_all)
                layout_xlsx = dataframe_para_xlsx_bytes(df)
                layout_xlsx = aplicar_destaque_vermelho(layout_xlsx, problemas)

                df_benef, erros = processar_beneficios(
                    df, df_benef_pronto=df_benef_cfg, regras=regras_all,
                    ultimo_id_easyapp=ultimo_id_all if usar_seq2 else None,
                    usar_id_sequencial=usar_seq2,
                )
                benef_xlsx = dataframe_para_xlsx_bytes(df_benef)
                benef_xlsx = aplicar_cores_colaborador_id(benef_xlsx)

                csv_layout = xlsx_bytes_para_csv_bytes(layout_xlsx)
                csv_benef  = xlsx_bytes_para_csv_bytes(benef_xlsx)

            if usar_seq2:
                set_ultimo_id(int(ultimo_id_all) + len(df))

            st.session_state["layout_df"] = df
            st.session_state["layout_xlsx"] = layout_xlsx
            st.session_state["benef_df"] = df_benef
            st.session_state["benef_xlsx"] = benef_xlsx
            st.session_state["tudo_pronto"] = True

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Colaboradores", len(df))
            c2.metric("Pendências críticas", len(problemas))
            c3.metric("Linhas Benefícios", len(df_benef))
            c4.metric("Sem benefício", sum(df_benef["beneficio_id"].isna()))

            st.success("Pronto! Baixe os 4 arquivos abaixo:")
            colA, colB = st.columns(2)
            colA.download_button(
                "⬇️  Importar Layout.xlsx", data=layout_xlsx,
                file_name="Importar Layout.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
            colB.download_button(
                "⬇️  Importar Beneficios.xlsx", data=benef_xlsx,
                file_name="Importar Beneficios.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
            colC, colD = st.columns(2)
            colC.download_button(
                "⬇️  Importar Layout.csv", data=csv_layout,
                file_name="Importar Layout.csv",
                mime="text/csv", use_container_width=True,
            )
            colD.download_button(
                "⬇️  Importar Beneficios.csv", data=csv_benef,
                file_name="Importar Beneficios.csv",
                mime="text/csv", use_container_width=True,
            )

            if problemas:
                st.warning("Linhas com campos críticos faltando:")
                st.dataframe(pd.DataFrame(problemas), use_container_width=True)
            if erros:
                st.warning("Colaboradores sem benefício identificado:")
                st.dataframe(pd.DataFrame(erros), use_container_width=True)

    card_fim()


# ══════════════════════════════════════════════════════════════════════════
# ABA 5 — AJUDA
# ══════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════
# ABA: VERIFICAÇÃO DE DUPLICADOS (Domínio × Sistema interno)
# ══════════════════════════════════════════════════════════════════════════
with tab_dup:
    st.subheader("🧹 Verificação de colaboradores duplicados")
    st.markdown(
        "Cruze o **Arquivo Domínio** com a exportação de colaboradores do "
        "sistema interno (AppLider/EasyApp) para descobrir quem **já está "
        "cadastrado** — por CPF (principal) ou Nome. O arquivo **limpo** fica "
        "salvo na sessão para usar nas abas 1 e 4."
    )
    col_a, col_b = st.columns(2)
    with col_a:
        f_dom_dup = st.file_uploader("Arquivo Domínio", type=["xls", "xlsx"], key="up_dup_dom")
    with col_b:
        f_sis_dup = st.file_uploader("Colaboradores do sistema interno", type=["xls", "xlsx"], key="up_dup_sis")

    def _ler_plan_dup(up):
        dados = up.read()
        if up.name.lower().endswith(".xlsx"):
            return pd.read_excel(io.BytesIO(dados), engine="openpyxl")
        return pd.read_excel(io.BytesIO(dados), engine="xlrd")

    if st.button("🔍 Verificar duplicados", type="primary", key="btn_dup"):
        if f_dom_dup is None or f_sis_dup is None:
            st.error("Envie os dois arquivos.")
        else:
            try:
                df_dup, df_limpo, stats = comparar_colaboradores(
                    _ler_plan_dup(f_dom_dup), _ler_plan_dup(f_sis_dup)
                )
                st.session_state["dup_relatorio"] = df_dup
                st.session_state["dup_stats"] = stats
                st.session_state["dominio_limpo_xlsx"] = dataframe_para_xlsx_bytes(df_limpo)
            except Exception as e:
                st.error("Falha na verificação: " + str(e))

    # Resultados persistem na sessão (não somem após download)
    if "dup_stats" in st.session_state:
        stats = st.session_state["dup_stats"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("No Domínio", stats["total_dominio"])
        c2.metric("Já cadastrados", stats["duplicados"])
        c3.metric("Novos (importar)", stats["novos"])
        c4.metric("No sistema", stats["total_sistema"])
        rel = st.session_state["dup_relatorio"]
        if not rel.empty:
            st.warning("⚠️ Estes colaboradores JÁ existem no sistema:")
            st.dataframe(rel, use_container_width=True)
        else:
            st.success("✅ Nenhum duplicado encontrado — todos são novos.")
        st.download_button(
            "⬇️ Baixar Arquivo Domínio LIMPO (.xlsx)",
            data=st.session_state["dominio_limpo_xlsx"],
            file_name="Arquivo Dominio - SEM DUPLICADOS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_dup_limpo",
        )

with tab_help:
    card_inicio("Como usar")
    st.markdown("""
- **Aba 1**: gere a planilha **Importar Layout.xlsx** a partir do Arquivo Domínio. Campos críticos faltantes (**posto**, **função**, **escala**) aparecem em vermelho.
- **Aba 2**: gere a planilha **Importar Beneficios.xlsx** (layout novo: `id` · `parceiro_id` · `beneficio_id`) com os benefícios certos para cada colaborador, com a opção de **gerar o ID em sequência a partir do último ID do EasyApp**.
- **Aba 3**: converte qualquer `.xlsx` para `.csv` preservando zeros à esquerda, datas e textos.
- **Aba 4 (Executar Tudo)**: faz todo o fluxo de uma vez.
- **Sidebar**: os arquivos de configuração (**Mapeamento** e **Benefícios Relação**) ficam **salvos** na pasta `data/` — envie uma vez e só reenvie quando quiser atualizar (marque "Enviar nova planilha").
- **Colunas Originais**: não é mais necessária — a ordem das colunas do layout é fixa no código.
- **Segurança**: o acesso é restrito (login). Configure os usuários no `.streamlit/secrets.toml`.

---

### 🎁 Regras de benefícios (resumo)

**Benefício principal (VA)** — a primeira regra que bater vence:

| # | Condição | Benefício |
|---|----------|-----------|
| P0 | Menor de 18 anos | *(sem VA)* |
| P1 | Função VIGILANTE | 56 · VA Vigilante VSP |
| P2 | Função MOTORISTA | 40 · VA Motorista |
| P3 | Função SUPERVISOR | 63 · VA ADM 700 |
| P4 | Posto "VALE" | 136 · VA Vale |
| P5 | Posto VPORTS / MULTILIFT | 54 · VA VPORTS |
| P6 | Escala 12x36 | 46 (TB1) / 47 (TB2) |
| P7 | Escala 5x2/6x1 · 6h+ | 44 (TB1) / 45 (TB2) |
| P8 | Escala 5x2/6x1 · 4h | 42 (TB1) / 43 (TB2) |

**TB2** = postos com GUIDONI, CESAN, TVV, PORTMAC, MULTILIFT ou BRITANIA.

**Extras (acumulam)**: 53 Petrolina · 51 Cozinha Amb. Vale · 41 Motorista ·
55 Coffee VPORTS/MULTILIFT · 48 Guarda-vida · 137 Desjejum Vale · 136 VA Vale.

**VT (Grande Vitória-ES)**: 58 (posto CETURB) ou 59 (demais postos) —
somente para quem mora em Vitória, Vila Velha, Serra, Cariacica ou Viana.
""")
    card_fim()
