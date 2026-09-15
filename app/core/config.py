"""
╔══════════════════════════════════════════════════════════════════════════╗
║  CONFIG.PY — Listas de regras configuráveis (aba LISTAS do Google Sheets)║
╠══════════════════════════════════════════════════════════════════════════╣
║  As listas que o setor pode querer ajustar com o tempo (postos TB2,      ║
║  cidades da Grande Vitória, postos com benefício próprio etc.) ficam     ║
║  na aba LISTAS da planilha do Google Sheets.                             ║
║                                                                          ║
║  • Se a aba LISTAS estiver disponível → os valores vêm DELA.             ║
║  • Se não estiver (fallback local/planilha antiga) → valem os PADRÕES    ║
║    definidos aqui no código (PADRAO_*), que refletem as regras           ║
║    validadas com o setor em set/2026.                                    ║
║                                                                          ║
║  Formato da aba LISTAS:                                                  ║
║      LISTA | VALOR | DESCRICAO | EXTRA                                   ║
║      POSTOS_TB2 | GUIDONI | ...                                          ║
║      EMPRESA_PARA_POSTO_ADM | 1 | VSP ... | ADM - VSP VIGILANCIA         ║
╚══════════════════════════════════════════════════════════════════════════╝
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List

import pandas as pd

from .utils import normalizar_texto


# ══════════════════════════════════════════════════════════════════════════
# VALORES PADRÃO (usados quando a aba LISTAS não está disponível)
# ══════════════════════════════════════════════════════════════════════════

PADRAO_POSTOS_TB2 = ["GUIDONI", "CESAN", "TVV", "PORTMAC", "MULTILIFT", "BRITANIA", "BRASITALIA"]
PADRAO_POSTOS_VA_VPORTS = ["VPORTS", "MULTILIFT"]
PADRAO_POSTOS_COFFEE_VPORTS = ["VPORTS", "MULTILIFT"]
PADRAO_POSTOS_VT_CETURB = ["CETURB"]
PADRAO_CIDADES_GRANDE_VITORIA = ["VITORIA", "VILA VELHA", "SERRA", "CARIACICA", "VIANA"]

PADRAO_EMPRESA_PARA_POSTO_ADM = {
    "1": "ADM - VSP VIGILANCIA",     # → EasyApp: ADM - VSP (id 257)
    "2": "BASE - ADMINISTRATIVO",    # → EasyApp: ADM - ATIVA (id 501)
    "3": "ADM - LIDER MULTISSERVICOS",
    "4": "ADM - LIDER LIMPE",
    "5": "ADM - B2WE",
}


# ══════════════════════════════════════════════════════════════════════════
# ESTRUTURA DE REGRAS
# ══════════════════════════════════════════════════════════════════════════
@dataclass
class RegrasConfig:
    """Conjunto de listas usadas pelas regras de benefícios e de importação."""
    postos_tb2: List[str] = field(default_factory=lambda: list(PADRAO_POSTOS_TB2))
    postos_va_vports: List[str] = field(default_factory=lambda: list(PADRAO_POSTOS_VA_VPORTS))
    postos_coffee_vports: List[str] = field(default_factory=lambda: list(PADRAO_POSTOS_COFFEE_VPORTS))
    postos_vt_ceturb: List[str] = field(default_factory=lambda: list(PADRAO_POSTOS_VT_CETURB))
    cidades_grande_vitoria: List[str] = field(default_factory=lambda: list(PADRAO_CIDADES_GRANDE_VITORIA))
    empresa_para_posto_adm: Dict[str, str] = field(
        default_factory=lambda: dict(PADRAO_EMPRESA_PARA_POSTO_ADM)
    )
    origem: str = "padrão do código"


# ══════════════════════════════════════════════════════════════════════════
# LEITURA DA ABA LISTAS
# ══════════════════════════════════════════════════════════════════════════
def carregar_regras(df_listas: pd.DataFrame | None) -> RegrasConfig:
    """
    Monta o RegrasConfig a partir do DataFrame da aba LISTAS.
    Se o DataFrame for None/vazio (fallback local), devolve os padrões.
    """
    cfg = RegrasConfig()

    if df_listas is None or df_listas.empty or "LISTA" not in df_listas.columns:
        return cfg

    df = df_listas.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "EXTRA" not in df.columns:
        df["EXTRA"] = ""

    df["LISTA"] = df["LISTA"].apply(normalizar_texto)
    df = df[df["LISTA"].notna() & (df["LISTA"] != "")]

    def _valores(nome_lista: str) -> List[str]:
        itens = df[df["LISTA"] == nome_lista]["VALOR"].dropna().tolist()
        return [normalizar_texto(v) for v in itens if str(v).strip() != ""]

    # Listas simples — só sobrescrevem o padrão se houver algo na planilha
    listas_simples = {
        "POSTOS_TB2": "postos_tb2",
        "POSTOS_VA_VPORTS": "postos_va_vports",
        "POSTOS_COFFEE_VPORTS": "postos_coffee_vports",
        "POSTOS_VT_CETURB": "postos_vt_ceturb",
        "CIDADES_GRANDE_VITORIA": "cidades_grande_vitoria",
    }
    for nome_lista, atributo in listas_simples.items():
        valores = _valores(nome_lista)
        if valores:
            setattr(cfg, atributo, valores)

    # Mapa empresa → posto ADM (coluna EXTRA guarda o posto)
    linhas_emp = df[df["LISTA"] == "EMPRESA_PARA_POSTO_ADM"]
    if not linhas_emp.empty:
        mapa_emp: Dict[str, str] = {}
        for _, r in linhas_emp.iterrows():
            chave = str(r.get("VALOR", "")).strip()
            if chave.endswith(".0"):
                chave = chave[:-2]
            posto = str(r.get("EXTRA", "")).strip()
            if chave and posto:
                mapa_emp[chave] = normalizar_texto(posto)
        if mapa_emp:
            cfg.empresa_para_posto_adm = mapa_emp

    cfg.origem = "☁️ aba LISTAS do Google Sheets"
    return cfg
