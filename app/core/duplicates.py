# -*- coding: utf-8 -*-
"""
Módulo de detecção de colaboradores DUPLICADOS.

Cruza o Arquivo Domínio (admissões novas) com a planilha de colaboradores
exportada do sistema interno (AppLider/EasyApp) e identifica quem JÁ está
cadastrado — por CPF (principal) ou por Nome (fallback).
"""
import re
import unicodedata

import pandas as pd


def so_digitos(valor):
    return re.sub(r"\D", "", str(valor if valor is not None else ""))


def norm_cpf(valor):
    d = so_digitos(valor)
    if d.endswith("0") and len(d) > 11 and "." not in str(valor):
        pass
    return d.zfill(11) if d else ""


def norm_nome(valor):
    t = str(valor if valor is not None else "").strip().upper()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t)


def _norm_col(c):
    n = norm_nome(c).replace(":", "").replace("/", " ").replace(".", " ")
    return re.sub(r"\s+", " ", n).strip()


def _achar_coluna(df, alvos):
    """Encontra coluna por nome normalizado (exato ou contido)."""
    norm = {_norm_col(c): c for c in df.columns}
    for alvo in alvos:
        alvo_n = _norm_col(alvo)
        if alvo_n in norm:
            return norm[alvo_n]
    for alvo in alvos:
        alvo_n = _norm_col(alvo)
        for n, original in norm.items():
            if alvo_n and alvo_n in n:
                return original
    return None


def comparar_colaboradores(df_dom, df_sis):
    """
    Retorna (df_duplicados, df_dominio_limpo, stats).

    df_duplicados: relação visual de quem já existe no sistema interno.
    df_dominio_limpo: Arquivo Domínio SEM os duplicados (pronto p/ importação).
    """
    col_cpf_dom = _achar_coluna(df_dom, ["CPF"])
    col_nome_dom = _achar_coluna(df_dom, ["Nome"])
    col_esocial = _achar_coluna(df_dom, ["Cód eSocial", "Cod eSocial", "eSocial"])

    col_cpf_sis = _achar_coluna(df_sis, ["CPF", "CPF:"])
    col_nome_sis = _achar_coluna(df_sis, ["Nome"])
    col_id_sis = _achar_coluna(df_sis, ["Id:", "Id", "ID"])
    col_mat_sis = _achar_coluna(df_sis, ["Matricula", "Matrícula"])

    if not col_cpf_dom or not col_cpf_sis:
        raise ValueError(
            "Não encontrei a coluna de CPF em um dos arquivos. "
            f"Colunas Domínio: {list(df_dom.columns)[:12]} | "
            f"Colunas Sistema: {list(df_sis.columns)[:12]}"
        )

    # Índices do sistema interno
    sis_por_cpf, sis_por_nome = {}, {}
    for _, r in df_sis.iterrows():
        cpf = norm_cpf(r.get(col_cpf_sis))
        nome = norm_nome(r.get(col_nome_sis)) if col_nome_sis else ""
        info = {
            "id": r.get(col_id_sis, "") if col_id_sis else "",
            "matricula": r.get(col_mat_sis, "") if col_mat_sis else "",
            "nome": r.get(col_nome_sis, "") if col_nome_sis else "",
        }
        if cpf and cpf not in sis_por_cpf:
            sis_por_cpf[cpf] = info
        if nome and nome not in sis_por_nome:
            sis_por_nome[nome] = info

    registros = []
    mascara_duplicado = []

    for _, r in df_dom.iterrows():
        cpf = norm_cpf(r.get(col_cpf_dom))
        nome = norm_nome(r.get(col_nome_dom)) if col_nome_dom else ""

        match, origem = None, None
        if cpf and cpf in sis_por_cpf:
            match, origem = sis_por_cpf[cpf], "CPF"
        elif nome and nome in sis_por_nome:
            match, origem = sis_por_nome[nome], "Nome"

        mascara_duplicado.append(match is not None)

        if match is not None:
            registros.append(
                {
                    "Nome (Domínio)": r.get(col_nome_dom, "") if col_nome_dom else "",
                    "CPF": cpf,
                    "Matrícula eSocial (Domínio)": r.get(col_esocial, "") if col_esocial else "",
                    "ID no Sistema": match["id"],
                    "Matrícula no Sistema": match["matricula"],
                    "Nome no Sistema": match["nome"],
                    "Encontrado por": origem,
                }
            )

    df_duplicados = pd.DataFrame(registros)
    df_limpo = df_dom.loc[[not m for m in mascara_duplicado]].reset_index(drop=True)

    stats = {
        "total_dominio": int(len(df_dom)),
        "duplicados": int(sum(mascara_duplicado)),
        "novos": int(len(df_limpo)),
        "total_sistema": int(len(df_sis)),
    }
    return df_duplicados, df_limpo, stats
