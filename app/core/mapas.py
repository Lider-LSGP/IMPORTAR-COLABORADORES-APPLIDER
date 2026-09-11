"""
╔══════════════════════════════════════════════════════════════════════════╗
║  MAPAS.PY — Carga dos mapas "de/para" do `Mapeamento Sistema.xls`        ║
╠══════════════════════════════════════════════════════════════════════════╣
║  Lê as 12 abas da planilha de mapeamento e transforma cada uma em um     ║
║  dicionário Python:                                                      ║
║      MAPA_<ABA>_ID    →  texto Domínio (normalizado)  →  ID no EasyApp   ║
║      MAPA_<ABA>_NOME  →  texto Domínio (normalizado)  →  Nome no EasyApp ║
║                                                                          ║
║  Abas com tratamento especial:                                           ║
║    • CIDADES  → também gera MAPA_CIDADE_PARA_MUNICIPIO                   ║
║    • HORARIOS → chave é o PADRÃO da jornada (ex.: "05X02_07:30_11:00")   ║
╚══════════════════════════════════════════════════════════════════════════╝
"""
from __future__ import annotations

import io
import pandas as pd

from .utils import normalizar_texto, extrair_padrao_horario


# ══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DAS ABAS
#   chave   = nome da aba no .xls
#   valor   = coluna usada como CHAVE de busca (o texto que vem do Domínio)
# ══════════════════════════════════════════════════════════════════════════
CONFIG_ABAS = {
    "POSTOS":        "Nome_Dominio",
    "CIDADES":       "Nome_Dominio",
    "FUNCOES":       "Nome_Dominio",
    "HORARIOS":      "Nome_EasyApp",
    "ESCALAS":       "TipoEscala",
    "RACA":          "Nome_EasyApp",
    "ESCOLARIDADE":  "Nome_Dominio",
    "ESTADOCIVIL":   "Nome_Dominio",
    "EMPRESAS":      "ID_Dominio",
    "DESLIGAMENTO":  "Nome_Dominio",
    "SINDICATOS":    "Nome_Dominio",
    "TIPOSALARIO":   "Nome_Dominio",
}


# ══════════════════════════════════════════════════════════════════════════
# FUNÇÃO PRINCIPAL — carrega TODAS as abas de uma vez
# ══════════════════════════════════════════════════════════════════════════
def carregar_todos_os_mapas(arquivo_mapeamento) -> dict:
    """
    `arquivo_mapeamento` pode ser:
      • path ou bytes do arquivo .xls (modo clássico);
      • dict {nome_da_aba: DataFrame} vindo do Google Sheets (gsheets.py).
    """
    abas_df: dict = {}
    if isinstance(arquivo_mapeamento, dict):
        # Fonte Google Sheets: já recebemos os DataFrames prontos
        abas_df = arquivo_mapeamento
        xls = None
    elif isinstance(arquivo_mapeamento, (bytes, bytearray)):
        xls = pd.ExcelFile(io.BytesIO(arquivo_mapeamento), engine="xlrd")
    else:
        xls = pd.ExcelFile(arquivo_mapeamento, engine="xlrd")

    mapas: dict = {}

    for aba, chave_busca in CONFIG_ABAS.items():
        try:
            if xls is not None:
                df = pd.read_excel(xls, aba)
            else:
                df = abas_df[aba].copy()
                # Google Sheets devolve células vazias como "" — converte
                # para NaN para que o dropna funcione igual ao .xls
                df = df.replace("", pd.NA)
                # Google Sheets lê tudo como texto; IDs precisam virar número
                for col in df.columns:
                    if str(col).upper().startswith("ID_") or str(col) == "CodIbge":
                        df[col] = pd.to_numeric(df[col], errors="coerce")
            if aba == "CIDADES":
                df = df.dropna(subset=["Nome_Dominio", "Nome_EasyApp"], how="all")
            else:
                df = df.dropna(subset=[chave_busca])

            # ──────────────────────────────────────────────────────────────
            # CIDADES — aceita chave pelo nome Domínio E pelo nome EasyApp;
            # também monta o de/para Cidade → Município (IBGE).
            # ──────────────────────────────────────────────────────────────
            if aba == "CIDADES":
                df["chave_dom"] = df["Nome_Dominio"].apply(normalizar_texto)
                if "Nome_EasyApp" in df.columns:
                    df["chave_easy"] = df["Nome_EasyApp"].apply(normalizar_texto)
                else:
                    df["chave_easy"] = df["chave_dom"]
                mapa_id, mapa_nome = {}, {}
                for _, row in df.iterrows():
                    key_dom = row["chave_dom"]
                    key_easy = row["chave_easy"]
                    id_easy = row.get("ID_EasyApp", None)
                    nome_easy = row.get("Nome_EasyApp", None)
                    if pd.notna(key_dom) and id_easy is not None:
                        mapa_id[key_dom] = id_easy
                        mapa_nome[key_dom] = nome_easy
                    if pd.notna(key_easy) and id_easy is not None:
                        mapa_id[key_easy] = id_easy
                        mapa_nome[key_easy] = nome_easy
                mapas["MAPA_CIDADES_ID"] = mapa_id
                mapas["MAPA_CIDADES_NOME"] = mapa_nome
                if "ID_EasyApp" in df.columns and "ID_EasyApp_Municipio" in df.columns:
                    mapas["MAPA_CIDADE_PARA_MUNICIPIO"] = (
                        df.set_index("ID_EasyApp")["ID_EasyApp_Municipio"].to_dict()
                    )
                else:
                    mapas["MAPA_CIDADE_PARA_MUNICIPIO"] = {}

            # ──────────────────────────────────────────────────────────────
            # HORARIOS — a chave é o padrão extraído da jornada
            # (escala + horários), ex.: "12X36_19:00_23:00_00:00_07:00"
            # ──────────────────────────────────────────────────────────────
            elif aba == "HORARIOS":
                df["chave"] = df["Nome_EasyApp"].apply(extrair_padrao_horario)
                df = df.dropna(subset=["chave"])
                mapas["MAPA_HORARIOS_ID"] = df.set_index("chave")["ID_EasyApp"].to_dict()
                mapas["MAPA_HORARIOS_NOME"] = df.set_index("chave")["Nome_EasyApp"].to_dict()

            # ──────────────────────────────────────────────────────────────
            # DEMAIS ABAS — chave = texto normalizado da coluna configurada
            # ──────────────────────────────────────────────────────────────
            else:
                df["chave"] = df[chave_busca].apply(normalizar_texto)
                mapas[f"MAPA_{aba}_ID"] = df.set_index("chave")["ID_EasyApp"].to_dict()
                mapas[f"MAPA_{aba}_NOME"] = df.set_index("chave")["Nome_EasyApp"].to_dict()

        except Exception as e:
            print(f"⚠️ Falha ao carregar aba [{aba}]: {e}")

    return mapas
