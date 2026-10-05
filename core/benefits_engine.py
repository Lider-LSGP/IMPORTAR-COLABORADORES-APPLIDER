# -*- coding: utf-8 -*-
"""
Motor de geração do "Importar Beneficios.xlsx".

Port 1:1 das regras do script original BeneficiosImportar.py, com o recurso
adicional de gerar o colaborador_id em sequência a partir do último ID
cadastrado no EasyApp/AppLider (opcional).
"""
import re

import pandas as pd


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper().replace("Ç", "C")


def calcular_carga_horaria(horario_str):
    try:
        texto = str(horario_str)
        horarios = re.findall(r"\d{2}:\d{2}", texto)
        total_horas = 0
        for i in range(0, len(horarios), 2):
            if i + 1 < len(horarios):
                h1, m1 = map(int, horarios[i].split(":"))
                h2, m2 = map(int, horarios[i + 1].split(":"))
                entrada = h1 + m1 / 60
                saida = h2 + m2 / 60
                total_horas += saida - entrada
        return total_horas
    except Exception:
        return 0


def tratar_base_beneficios(df):
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    df.columns = [str(col).strip().lower() for col in df.columns]

    colunas_esperadas = ["beneficio_id", "beneficio_nome", "sigla", "valor_unico", "qtd_dias", "valor_total"]
    for c in colunas_esperadas:
        if c not in df.columns:
            df[c] = None

    df["sigla"] = df["sigla"].astype(str).str.strip().str.upper()
    df["beneficio_nome"] = df["beneficio_nome"].astype(str).str.strip().str.upper()

    for c in ["valor_unico", "qtd_dias", "valor_total"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["valor_total"] = df["valor_total"].round(2)
    df["valor_unico"] = df["valor_unico"].round(2)
    return df


# ---------------------------------------------------------------------------
# Regras de negócio (idênticas ao original)
# ---------------------------------------------------------------------------
def encontrar_beneficios(row, df_benef):
    funcao = normalizar_texto(row.get("nomefuncao"))
    escala = normalizar_texto(row.get("tipoescala"))
    posto = normalizar_texto(row.get("nomepostotrabalho"))
    horario = row.get("nomehorariotrabalho", "")
    carga = calcular_carga_horaria(horario)

    df = df_benef[df_benef["sigla"] == "VA"].copy()

    if any(x in funcao for x in ["VIGILANTE", "VIGILANCIA"]):
        return df_benef[df_benef["beneficio_id"] == 56]

    if "VALE" in posto:
        return df[df["beneficio_id"] == 136]

    if "12X36" in escala:
        filtro_tipo = "12X36"
    elif any(x in escala for x in ["5X2", "6X1"]):
        filtro_tipo = "6H" if carga >= 6 else "4H"
    else:
        return pd.DataFrame()

    if any(x in posto for x in ["GUIDONI", "CESAN", "VPORTS", "BRITANIA"]):
        filtro_tb = "TB2"
    else:
        filtro_tb = "TB1"

    return df[
        df["beneficio_nome"].str.contains(filtro_tipo, case=False, na=False)
        & df["beneficio_nome"].str.contains(filtro_tb, case=False, na=False)
    ]


def encontrar_extras(row, df_benef):
    funcao = normalizar_texto(row.get("nomefuncao"))
    posto = normalizar_texto(row.get("nomepostotrabalho"))
    extras_ids = set()

    if "PETROLINA" in posto:
        extras_ids.add(53)
    if "AMBIENTAL VALE" in posto:
        if "COZINHEIRO" in funcao or "AJUDANTE DE COZINHA" in funcao:
            extras_ids.add(51)
    if "MOTORISTA" in funcao:
        extras_ids.add(41)
    if "VPORTS" in posto:
        extras_ids.add(55)
    if "GUARDA VIDA" in funcao:
        extras_ids.add(48)
    if "AMBIENTAL VALE" in posto and "LAVADOR DE VEICULOS PESADO" in funcao:
        extras_ids.add(137)
    if "AMBIENTAL VALE" in posto or "INSTITUTO AMBIENTAL VALE" in posto:
        extras_ids.add(136)

    if not extras_ids:
        return pd.DataFrame()
    return df_benef[df_benef["beneficio_id"].isin(extras_ids)].copy()


# ---------------------------------------------------------------------------
# Processamento principal
# ---------------------------------------------------------------------------
def processar_beneficios_df(df_func, df_benef, modo="esocial", ultimo_id=None):
    """
    modo = 'esocial'    -> colaborador_id = matrícula (Cód eSocial)  [comportamento original]
    modo = 'sequencial' -> colaborador_id = último ID EasyApp + 1, +2, ... por matrícula única

    Retorna (df_saida, mapa_ids, stats)
      - mapa_ids: DataFrame novo_id x matricula x nome (apenas no modo sequencial)
      - stats: dict com totais
    """
    df_benef = tratar_base_beneficios(df_benef)
    linhas = []

    for _, row in df_func.iterrows():
        beneficios_principais = encontrar_beneficios(row, df_benef)
        beneficios_extras = encontrar_extras(row, df_benef)

        beneficios_encontrados = pd.concat(
            [beneficios_principais, beneficios_extras], ignore_index=True
        ).drop_duplicates(subset=["beneficio_id"])

        if beneficios_encontrados.empty:
            nova = row.copy()
            nova["beneficio_id"] = None
            nova["beneficio_nome"] = None
            nova["qtd_dias"] = None
            nova["valor_unico"] = None
            nova["valor_total"] = None
            nova["sigla"] = None
            linhas.append(nova)
        else:
            for _, b in beneficios_encontrados.iterrows():
                if pd.isna(b["beneficio_id"]):
                    continue
                nova = row.copy()
                beneficio_id = int(b["beneficio_id"])
                nova["beneficio_id"] = beneficio_id
                nova["beneficio_nome"] = b["beneficio_nome"]
                nova["sigla"] = b["sigla"]

                if beneficio_id == 136:
                    nova["qtd_dias"] = 1
                    nova["valor_unico"] = 1106.70
                    nova["valor_total"] = 1106.70
                elif beneficio_id == 137:
                    nova["qtd_dias"] = 1
                    nova["valor_unico"] = 154.49
                    nova["valor_total"] = 154.49
                else:
                    nova["qtd_dias"] = b["qtd_dias"]
                    nova["valor_unico"] = b["valor_unico"]
                    nova["valor_total"] = b["valor_total"]
                linhas.append(nova)

    df_final = pd.DataFrame(linhas)
    if df_final.empty:
        raise ValueError("Nenhum colaborador para processar benefícios.")
    df_final["erro_beneficio"] = df_final["beneficio_id"].isna()

    # ---- colaborador_id ----
    mapa_ids = None
    if modo == "sequencial":
        mats = df_final["matricula"].astype(str).str.strip()
        uniques = list(dict.fromkeys(mats.tolist()))  # ordem de aparecimento
        base = int(ultimo_id or 0)
        sequencia = {m: base + i + 1 for i, m in enumerate(uniques)}
        colaborador_ids = mats.map(sequencia)

        nomes = (
            df_final.assign(_mat=mats)
            .drop_duplicates(subset=["_mat"])[["_mat", "nome"]]
            .rename(columns={"_mat": "matricula"})
        )
        mapa_ids = nomes.copy()
        mapa_ids["novo_id"] = mapa_ids["matricula"].map(sequencia)
        mapa_ids = mapa_ids[["novo_id", "matricula", "nome"]].reset_index(drop=True)
    else:
        colaborador_ids = df_final["matricula"]

    df_saida = pd.DataFrame()
    df_saida["coluna1"] = ""
    df_saida["colaborador_id"] = colaborador_ids.values
    df_saida["beneficio_id"] = df_final["beneficio_id"].values
    df_saida["qtd_dias"] = df_final["qtd_dias"].values
    df_saida["valor_unico"] = df_final["valor_unico"].values
    df_saida["valor_total"] = df_final["valor_total"].values
    for i in range(2, 8):
        df_saida[f"coluna{i}"] = ""
    df_saida["sigla"] = df_final["sigla"].values
    df_saida["beneficio_nome"] = df_final["beneficio_nome"].values

    stats = {
        "colaboradores": int(df_final["matricula"].nunique()),
        "linhas": int(len(df_saida)),
        "sem_beneficio": int(df_final["erro_beneficio"].sum()),
    }
    return df_saida, mapa_ids, stats
