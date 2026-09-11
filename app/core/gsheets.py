"""
╔══════════════════════════════════════════════════════════════════════════╗
║  GSHEETS.PY — Leitura da planilha de BENEFÍCIOS direto do Google Sheets  ║
╠══════════════════════════════════════════════════════════════════════════╣
║  POR QUÊ?                                                                ║
║  A planilha "BENEFICIOS - RELAÇÃO" agora vive no Google Sheets.          ║
║  Assim o setor pode atualizar valores/listas SEM precisar fazer          ║
║  deploy nem commit no GitHub: editou a planilha → o app já usa.          ║
║                                                                          ║
║  COMO O APP LÊ (ordem de tentativa):                                     ║
║   1) Conta de serviço (st.secrets["gcp_service_account"]) — robusto,     ║
║      funciona com planilha PRIVADA (compartilhada com a conta).          ║
║   2) Exportação CSV pública — funciona quando a planilha está com        ║
║      "Qualquer pessoa com o link · Leitor".                              ║
║   3) Fallback: arquivo local `data/BENEFICIOS - RELACAO.xlsx` —          ║
║      garante que o app NUNCA para por falta de internet.                 ║
║                                                                          ║
║  CONFIGURAÇÃO (Streamlit → Settings → Secrets):                          ║
║      [gsheets]                                                           ║
║      beneficios_sheet_id = "1DtYsnKwglTA2Xos7oTbb3uXkd8OQFgE-GXs-K-4JQsk"║
║                                                                          ║
║      # opcional (modo privado/robusto):                                  ║
║      [gcp_service_account]                                               ║
║      type = "service_account"                                            ║
║      project_id = "..."                                                  ║
║      private_key_id = "..."                                              ║
║      private_key = "-----BEGIN PRIVATE KEY-----\n...\n"                  ║
║      client_email = "...@....iam.gserviceaccount.com"                    ║
║      client_id = "..."                                                   ║
║      token_uri = "https://oauth2.googleapis.com/token"                   ║
╚══════════════════════════════════════════════════════════════════════════╝
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd

# ── Identificadores das planilhas oficiais no Google Sheets ─────────────
# (podem ser sobrescritos via st.secrets["gsheets"])
DEFAULT_SHEET_ID = "149DL4ip25O8IM3w7LS08tXQJfxdXoSdSQkrI2kMhV2c"      # BENEFICIOS v2
GID_BENEFICIOS = "663612116"      # aba BENEFICIOS
GID_LISTAS = "677614552"          # aba LISTAS

ABA_BENEFICIOS = "BENEFICIOS"
ABA_LISTAS = "LISTAS"

# Planilha do MAPEAMENTO SISTEMA no Google Sheets (12 abas)
DEFAULT_MAPEAMENTO_SHEET_ID = "1Rn3pYB9uHQysMMfNGgHccJPALDMqf9f4sOYk7MyZaxc"
MAPEAMENTO_ABAS = [
    "CIDADES", "HORARIOS", "ESCALAS", "FUNCOES", "POSTOS", "EMPRESAS",
    "ESCOLARIDADE", "RACA", "ESTADOCIVIL", "DESLIGAMENTO", "SINDICATOS",
    "TIPOSALARIO",
]


# ══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ══════════════════════════════════════════════════════════════════════════
def sheet_id_configurado() -> str | None:
    """ID da planilha de BENEFÍCIOS: secrets > padrão do código."""
    try:
        import streamlit as st
        sid = st.secrets.get("gsheets", {}).get("beneficios_sheet_id")
        if sid:
            return str(sid).strip()
    except Exception:
        pass
    return DEFAULT_SHEET_ID


def mapeamento_sheet_id_configurado() -> str | None:
    """ID da planilha de MAPEAMENTO: secrets > padrão do código."""
    try:
        import streamlit as st
        sid = st.secrets.get("gsheets", {}).get("mapeamento_sheet_id")
        if sid:
            return str(sid).strip()
    except Exception:
        pass
    return DEFAULT_MAPEAMENTO_SHEET_ID


def google_sheets_habilitado() -> bool:
    """
    True se o app deve tentar ler do Google Sheets.
    Pode ser desligado via secrets: [gsheets] enabled = false
    """
    try:
        import streamlit as st
        return bool(st.secrets.get("gsheets", {}).get("enabled", True))
    except Exception:
        return True


# ══════════════════════════════════════════════════════════════════════════
# MODO 1 — Conta de serviço (planilha privada, mais robusto)
# ══════════════════════════════════════════════════════════════════════════
def _ler_com_conta_servico(sheet_id: str, aba: str) -> pd.DataFrame:
    import streamlit as st
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    resp = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=f"{aba}!A1:Z1000")
        .execute()
    )
    valores = resp.get("values", [])
    if not valores:
        raise RuntimeError(f"Aba {aba} vazia no Google Sheets.")
    header, linhas = valores[0], valores[1:]
    # completa linhas mais curtas que o cabeçalho
    linhas = [l + [""] * (len(header) - len(l)) for l in linhas]
    return pd.DataFrame(linhas, columns=header)


# ══════════════════════════════════════════════════════════════════════════
# MODO 2 — Exportação CSV pública (planilha "qualquer pessoa com o link")
# ══════════════════════════════════════════════════════════════════════════
def _ler_csv_publico(sheet_id: str, gid: str) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(url, dtype=str, keep_default_na=False)
    if df.empty:
        raise RuntimeError("CSV público do Google Sheets veio vazio.")
    return df


# ══════════════════════════════════════════════════════════════════════════
# API PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════
def carregar_beneficios_google() -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lê as abas BENEFICIOS e LISTAS do Google Sheets.
    Retorna (df_beneficios, df_listas).
    Levanta exceção se nenhum modo funcionar.
    """
    sheet_id = sheet_id_configurado()
    if not sheet_id:
        raise RuntimeError("Google Sheets não configurado (sem sheet_id).")

    # 1) tenta conta de serviço (se houver credenciais nos secrets)
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            df_b = _ler_com_conta_servico(sheet_id, ABA_BENEFICIOS)
            df_l = _ler_com_conta_servico(sheet_id, ABA_LISTAS)
            return df_b, df_l
    except Exception:
        pass  # cai para o modo público

    # 2) tenta CSV público
    df_b = _ler_csv_publico(sheet_id, GID_BENEFICIOS)
    df_l = _ler_csv_publico(sheet_id, GID_LISTAS)
    return df_b, df_l


def carregar_beneficios_com_fallback(
    arquivo_local: Path,
    beneficios_relacao_bytes: bytes | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    """
    Estratégia à prova de falhas:
      1º) Google Sheets (se habilitado);
      2º) bytes enviados por upload nesta sessão;
      3º) arquivo local em data/.

    Retorna (df_beneficios, df_listas, origem) — origem é um texto
    explicando de onde vieram os dados (para mostrar na tela).
    """
    # 1) Google Sheets
    if google_sheets_habilitado():
        try:
            df_b, df_l = carregar_beneficios_google()
            return df_b, df_l, "☁️ Google Sheets (sempre atualizado)"
        except Exception:
            pass  # cai para os fallbacks locais

    # 2) upload da sessão
    if beneficios_relacao_bytes is not None:
        df_b = pd.read_excel(io.BytesIO(beneficios_relacao_bytes))
        return df_b, _listas_vazias(), "📤 Upload desta sessão"

    # 3) arquivo local
    if arquivo_local.exists():
        df_b = pd.read_excel(arquivo_local)
        return df_b, _listas_vazias(), f"💾 Arquivo local ({arquivo_local.name})"

    raise RuntimeError(
        "Nenhuma fonte de benefícios disponível "
        "(Google Sheets falhou e não há arquivo local)."
    )


def _listas_vazias() -> pd.DataFrame:
    return pd.DataFrame(columns=["LISTA", "VALOR", "DESCRICAO", "EXTRA"])


# ══════════════════════════════════════════════════════════════════════════
# MAPEAMENTO SISTEMA via Google Sheets
# ══════════════════════════════════════════════════════════════════════════
def carregar_mapeamento_google() -> dict[str, pd.DataFrame]:
    """
    Lê TODAS as abas do MAPEAMENTO SISTEMA no Google Sheets.
    Retorna {nome_da_aba: DataFrame}.
    """
    sheet_id = mapeamento_sheet_id_configurado()
    if not sheet_id:
        raise RuntimeError("Planilha de Mapeamento não configurada.")

    abas: dict[str, pd.DataFrame] = {}

    # 1) conta de serviço (planilha privada)
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            for aba in MAPEAMENTO_ABAS:
                abas[aba] = _ler_com_conta_servico(sheet_id, aba)
            return abas
    except Exception:
        pass  # cai para o modo público

    # 2) CSV público por aba — usa o endpoint gviz (não precisa do gid)
    for aba in MAPEAMENTO_ABAS:
        url = (
            f"https://docs.google.com/spreadsheets/d/{sheet_id}"
            f"/gviz/tq?tqx=out:csv&sheet={aba}"
        )
        df = pd.read_csv(url, dtype=str, keep_default_na=False)
        if df.empty:
            raise RuntimeError(f"Aba {aba} do Mapeamento veio vazia.")
        abas[aba] = df
    return abas


def carregar_mapeamento_com_fallback(
    arquivo_local: Path,
    mapeamento_bytes: bytes | None = None,
) -> tuple[object, str]:
    """
    Estratégia à prova de falhas para o MAPEAMENTO:
      1º) Google Sheets;  2º) upload da sessão;  3º) arquivo local data/.

    Retorna (fonte, origem) — `fonte` pode ser dict de DataFrames (Google),
    bytes (upload) ou Path (arquivo local). `origem` descreve de onde veio.
    """
    if google_sheets_habilitado():
        try:
            abas = carregar_mapeamento_google()
            return abas, "☁️ Google Sheets (sempre atualizado)"
        except Exception:
            pass  # fallback local

    if mapeamento_bytes is not None:
        return mapeamento_bytes, "📤 Upload desta sessão"

    if arquivo_local.exists():
        return arquivo_local, f"💾 Arquivo local ({arquivo_local.name})"

    raise RuntimeError(
        "Nenhuma fonte de Mapeamento disponível "
        "(Google Sheets falhou e não há arquivo local)."
    )
