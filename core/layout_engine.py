# -*- coding: utf-8 -*-
"""
Motor de geração do "Importar Layout.xlsx" (AppLider/EasyApp).

Port 1:1 das regras do script original ImportarDados.py, adaptado para
trabalhar com DataFrames em memória (sem depender de caminhos em disco).
"""
import calendar
import re
import unicodedata
from datetime import datetime

import pandas as pd

# Campos que ficam VERMELHOS na planilha quando não forem encontrados/mapeados
CAMPOS_OBRIGATORIOS = [
    "postotrabalho_id",
    "funcao_id",
    "escalatrabalho_id",
    "horariotrabalho_id",
]

ROTULOS_CAMPOS = {
    "postotrabalho_id": "Posto de Serviço",
    "funcao_id": "Função",
    "escalatrabalho_id": "Escala de Trabalho",
    "horariotrabalho_id": "Horário de Trabalho",
}

IDS_CARGOS_SUPERVISOR = [11, 12, 13, 14, 42]


# ---------------------------------------------------------------------------
# Funções de tratamento (idênticas ao original)
# ---------------------------------------------------------------------------
def limpar_documento(valor, tamanho=None):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return ""
    s = str(valor).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if tamanho:
        return s.zfill(tamanho)
    return s


def limpar_email(email):
    if pd.isna(email) or str(email).strip() == "":
        return ""
    email = str(email).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFKD", email) if not unicodedata.combining(c)
    )


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def identificar_tipo_escala(jornada, data_adm):
    if pd.isna(jornada):
        return calcular_variacao_escala(data_adm)
    j = str(jornada).upper().strip()
    if j in ["", "NAN", "NONE"]:
        return calcular_variacao_escala(data_adm)
    j_norm = re.sub(r"[^A-Z0-9]", "", j)
    if "12X36" in j_norm:
        return calcular_variacao_escala(data_adm)
    if "5X2" in j_norm or "5X02" in j_norm or "05X02" in j_norm:
        return "5x2"
    if "6X1" in j_norm or "6X01" in j_norm or "06X01" in j_norm:
        return "6x1"
    return ""


def calcular_variacao_escala(data_adm_str):
    """12x36 P/I conforme paridade do dia de admissão e inversão em meses 31/29."""
    if pd.isna(data_adm_str) or str(data_adm_str).strip() == "":
        return "12x36"
    try:
        data_adm = pd.to_datetime(data_adm_str, dayfirst=True)
        hoje = datetime.now()
        paridade = "P" if data_adm.day % 2 == 0 else "I"
        ano, mes = data_adm.year, data_adm.month
        while (ano < hoje.year) or (ano == hoje.year and mes < hoje.month):
            ultimo_dia_mes = calendar.monthrange(ano, mes)[1]
            if ultimo_dia_mes in [31, 29]:
                paridade = "I" if paridade == "P" else "P"
            mes += 1
            if mes > 12:
                mes = 1
                ano += 1
        return f"12x36{paridade}"
    except Exception:
        return "12x36"


def formatar_sexo(valor):
    v = str(valor).strip().upper()
    return "F" if v.startswith("F") else "M" if v.startswith("M") else ""


def formatar_data(valor):
    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%Y-%m-%d")
    except Exception:
        return None


def checar_supervisor(valor):
    try:
        if pd.isna(valor):
            return "Não"
        if int(float(valor)) in IDS_CARGOS_SUPERVISOR:
            return "Sim"
    except Exception:
        pass
    return "Não"


def escolher_escolaridade(row):
    sit = str(row.get("Situação", "")).strip().lower()
    grau = str(row.get("Grau instrução", "")).strip()
    if sit in ["", "nan", "none"] or "trabalhando" in sit:
        return grau
    return row.get("Situação", "")


def checar_ativo(valor):
    v = str(valor).strip().lower()
    if "trabalhando" in v:
        return "Sim"
    if v in ["nan", "", "none", "nat", "01/01/1900", "00/00/0000"]:
        return "Sim"
    return "Não"


def extrair_mes(valor):
    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%m")
    except Exception:
        return None


def extrair_ano(valor):
    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%Y")
    except Exception:
        return None


def extrair_mes_ano(valor):
    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%m/%Y")
    except Exception:
        return None


def formatar_salario(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "nan", "NaN"]:
        return "0.00"
    v = str(valor).strip().replace(".", "").replace(",", "")
    if len(v) <= 4:
        return "0.00"
    v_cortado = v[:-4]
    parte_inteira = v_cortado[:-2]
    centavos = v_cortado[-2:]
    return f"{parte_inteira}.{centavos}"


def extrair_padrao_horario(texto):
    if pd.isna(texto):
        return None
    texto = str(texto).upper()
    escala = re.search(r"(\d{1,2})\s*X\s*(\d{1,2})", texto)
    if not escala:
        return None
    esc1 = escala.group(1).zfill(2)
    esc2 = escala.group(2).zfill(2)
    escala_formatada = f"{esc1}X{esc2}"

    texto_limpo = texto.replace("H", ":").replace("-", " ").replace("/", " ").replace("AS", " ")
    possiveis = re.findall(r"\d{1,2}:?\d{0,2}", texto_limpo)

    horarios_formatados = []
    for h in possiveis:
        if ":" in h:
            partes = h.split(":")
            if len(partes) == 2:
                hora = partes[0].zfill(2)
                minuto = partes[1].zfill(2) if partes[1] else "00"
                horarios_formatados.append(f"{hora}:{minuto}")
        else:
            if len(h) == 4:
                horarios_formatados.append(f"{h[:2]}:{h[2:]}")
            elif len(h) == 3:
                horarios_formatados.append(f"0{h[0]}:{h[1:]}")
            elif len(h) <= 2:
                horarios_formatados.append(f"{h.zfill(2)}:00")

    horarios_validos = []
    for h in horarios_formatados:
        hora, minuto = h.split(":")
        if int(hora) <= 23 and int(minuto) <= 59:
            horarios_validos.append(h)

    if not horarios_validos:
        return None
    return escala_formatada + "_" + "_".join(horarios_validos)


def definir_tipo_pcd(row):
    possui = str(row.get("Possui deficiência", "")).strip().upper()
    if possui != "SIM":
        return ""
    fisica = str(row.get("Deficiência física", "")).strip().upper() == "SIM"
    visual = str(row.get("Deficiência visual", "")).strip().upper() == "SIM"
    auditiva = str(row.get("Deficiência auditiva", "")).strip().upper() == "SIM"
    intelectual = str(row.get("Deficiência intelectual", "")).strip().upper() == "SIM"
    mental = str(row.get("Deficiência mental", "")).strip().upper() == "SIM"

    tipos = []
    if fisica:
        tipos.append("Física")
    if visual:
        tipos.append("Visual")
    if auditiva:
        tipos.append("Auditiva")
    if intelectual:
        tipos.append("Intelectual")
    if mental:
        tipos.append("Psicossocial ou por Saúde Mental")

    if len(tipos) > 1:
        return "Múltipla"
    if len(tipos) == 1:
        return tipos[0]
    return ""


# ---------------------------------------------------------------------------
# Carregamento dos mapas (Mapeamento Sistema.xls)
# ---------------------------------------------------------------------------
def carregar_todos_os_mapas(arquivo):
    """arquivo: path ou BytesIO do 'Mapeamento Sistema.xls'."""
    xls = pd.ExcelFile(arquivo)
    mapas = {}

    config_abas = {
        "POSTOS": "Nome_Dominio",
        "CIDADES": "Nome_Dominio",
        "FUNCOES": "Nome_Dominio",
        "HORARIOS": "Nome_EasyApp",
        "ESCALAS": "TipoEscala",
        "RACA": "Nome_EasyApp",
        "ESCOLARIDADE": "Nome_Dominio",
        "ESTADOCIVIL": "Nome_Dominio",
        "EMPRESAS": "ID_Dominio",
        "DESLIGAMENTO": "Nome_Dominio",
        "SINDICATOS": "Nome_Dominio",
        "TIPOSALARIO": "Nome_Dominio",
    }

    for aba, chave_busca in config_abas.items():
        try:
            df = pd.read_excel(xls, aba)
            if aba == "CIDADES":
                df = df.dropna(subset=["Nome_Dominio", "Nome_EasyApp"], how="all")
            else:
                df = df.dropna(subset=[chave_busca])

            if aba == "CIDADES":
                df["chave_dom"] = df["Nome_Dominio"].apply(normalizar_texto)
                if "Nome_EasyApp" in df.columns:
                    df["chave_easy"] = df["Nome_EasyApp"].apply(normalizar_texto)
                else:
                    df["chave_easy"] = df["chave_dom"]

                mapa_id, mapa_nome = {}, {}
                for _, row in df.iterrows():
                    key_dom, key_easy = row["chave_dom"], row["chave_easy"]
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
                    mapas["MAPA_CIDADE_PARA_MUNICIPIO"] = df.set_index("ID_EasyApp")[
                        "ID_EasyApp_Municipio"
                    ].to_dict()
                else:
                    mapas["MAPA_CIDADE_PARA_MUNICIPIO"] = {}

            elif aba == "HORARIOS":
                df["chave"] = df["Nome_EasyApp"].apply(extrair_padrao_horario)
                df = df.dropna(subset=["chave"], how="any")
                mapas["MAPA_HORARIOS_ID"] = df.set_index("chave")["ID_EasyApp"].to_dict()
                mapas["MAPA_HORARIOS_NOME"] = df.set_index("chave")["Nome_EasyApp"].to_dict()

            else:
                df["chave"] = df[chave_busca].apply(normalizar_texto)
                mapas[f"MAPA_{aba}_ID"] = df.set_index("chave")["ID_EasyApp"].to_dict()
                mapas[f"MAPA_{aba}_NOME"] = df.set_index("chave")["Nome_EasyApp"].to_dict()

        except Exception:
            # aba ausente não derruba o processo (comportamento do original)
            continue

    return mapas


# ---------------------------------------------------------------------------
# Motor principal
# ---------------------------------------------------------------------------
def processar_layout(arq_dom, mapa, colunas_modelo):
    """
    Gera o DataFrame do Importar Layout a partir do DataFrame da Domínio.
    Retorna (df_layout, pendencias) onde pendencias = [(indice_linha, coluna), ...]
    das células obrigatórias que ficaram vazias (para pintar de vermelho).
    """
    arq_lider = pd.DataFrame(index=arq_dom.index)
    agora = datetime.now()

    def col(nome):
        if nome in arq_dom.columns:
            return arq_dom[nome]
        return pd.Series([""] * len(arq_dom), index=arq_dom.index)

    def clean(nome):
        return col(nome).astype(str).str.upper().str.strip().replace({"NAN": "", "NONE": ""})

    def mapa_de(chave):
        return mapa.get(chave) or {}

    # ---- Identificação e Empresa ----
    arq_lider["empresa_id"] = clean("Cód Emp").map(mapa_de("MAPA_EMPRESAS_ID"))
    arq_lider["nomeempresa"] = clean("Cód Emp").map(mapa_de("MAPA_EMPRESAS_NOME"))
    arq_lider["sexo"] = col("Sexo").apply(formatar_sexo)
    arq_lider["matricula"] = col("Cód eSocial")
    arq_lider["nome"] = col("Nome")
    arq_lider["cpf"] = col("CPF").apply(lambda x: limpar_documento(x, 11))

    # ---- Dados Pessoais ----
    arq_lider["racaetinia_id"] = clean("Raça/Cor").map(mapa_de("MAPA_RACA_ID"))
    escolaridade_escolhida = arq_dom.apply(escolher_escolaridade, axis=1)
    escolaridade_norm = escolaridade_escolhida.apply(normalizar_texto)
    arq_lider["escolaridade_id"] = escolaridade_norm.map(mapa_de("MAPA_ESCOLARIDADE_ID"))
    arq_lider["estadocivil_id"] = col("Estado Civil").apply(normalizar_texto).map(
        mapa_de("MAPA_ESTADOCIVIL_ID")
    )
    arq_lider["telefone"] = col("Telefone").apply(limpar_documento)
    arq_lider["celular"] = col("Celular").apply(limpar_documento)
    arq_lider["email"] = col("Email").apply(limpar_email)
    arq_lider["nomepai"] = col("Nome Pai")
    arq_lider["nomemae"] = col("Nome Mãe")

    # ---- Documentos ----
    arq_lider["ci"] = col("RG")
    arq_lider["ciemissor"] = col("Orgão RG")
    arq_lider["pis"] = col("PIS").apply(limpar_documento)
    arq_lider["tituloeleitor"] = col("Titulo").apply(limpar_documento)
    arq_lider["tituloeleitorsecao"] = col("Seção").apply(limpar_documento)
    arq_lider["tituloeleitorzona"] = col("Zona").apply(limpar_documento)
    arq_lider["certificadomilitar"] = col("Reservista").apply(limpar_documento)
    arq_lider["cnh"] = col("CNH").apply(limpar_documento)
    arq_lider["cnhcategoria"] = col("Categoria CNH").astype(str).str.strip().str.upper()
    arq_lider["ctpsnumero"] = col("CTPS").apply(limpar_documento)
    arq_lider["ctpsserie"] = col("Serie CTPS").apply(limpar_documento)
    arq_lider["ctpsuf"] = col("UF CTPS")
    arq_lider["tipopcd"] = arq_dom.apply(definir_tipo_pcd, axis=1)

    # ---- Endereço e Localização ----
    cid_normalizada = col("Cidade").apply(normalizar_texto)
    arq_lider["cidade_id"] = cid_normalizada.map(mapa_de("MAPA_CIDADES_ID"))
    arq_lider["municipios_id"] = arq_lider["cidade_id"].map(mapa_de("MAPA_CIDADE_PARA_MUNICIPIO"))
    arq_lider["cidadeparceiro"] = cid_normalizada.map(mapa_de("MAPA_CIDADES_NOME"))
    arq_lider["cep"] = col("Cep").apply(limpar_documento)
    arq_lider["rua"] = col("Endereço")
    arq_lider["numero"] = col("Numero").apply(limpar_documento)
    arq_lider["bairro"] = col("Bairro")
    arq_lider["uf"] = col("UF End")
    arq_lider["complemento"] = col("Complemento")
    arq_lider["cidadetrabalho"] = cid_normalizada.map(mapa_de("MAPA_CIDADES_NOME"))

    # ---- Cargo/Função, Posto e Sindicato ----
    posto_norm = col("Descrição Ccusto").apply(normalizar_texto)
    funcao_norm = col("Descrição cargo").apply(normalizar_texto)
    sind_norm = col("Sindicato").apply(normalizar_texto)
    arq_lider["postotrabalho_id"] = posto_norm.map(mapa_de("MAPA_POSTOS_ID"))
    arq_lider["nomepostotrabalho"] = posto_norm.map(mapa_de("MAPA_POSTOS_NOME"))
    arq_lider["funcao_id"] = funcao_norm.map(mapa_de("MAPA_FUNCOES_ID"))
    arq_lider["nomefuncao"] = funcao_norm.map(mapa_de("MAPA_FUNCOES_NOME"))
    arq_lider["sindicato_id"] = sind_norm.map(mapa_de("MAPA_SINDICATOS_ID"))
    arq_lider["nomesindicato"] = sind_norm.map(mapa_de("MAPA_SINDICATOS_NOME"))
    arq_lider["supervisor"] = arq_lider["funcao_id"].apply(checar_supervisor)

    # ---- Escala e Horário (12x36 P/I) ----
    arq_lider["tipoescala"] = arq_dom.apply(
        lambda r: identificar_tipo_escala(r.get("Jornada"), r.get("Admissão")), axis=1
    )
    escala_norm = arq_lider["tipoescala"].apply(normalizar_texto)
    arq_lider["escalatrabalho_id"] = escala_norm.map(mapa_de("MAPA_ESCALAS_ID"))
    arq_lider["nomeescala"] = "GRADE DA ESCALA " + arq_lider["tipoescala"].fillna("").astype(str)

    jornada_key = col("Jornada").apply(extrair_padrao_horario)
    arq_lider["horariotrabalho_id"] = jornada_key.map(mapa_de("MAPA_HORARIOS_ID"))
    arq_lider["nomehorariotrabalho"] = jornada_key.map(mapa_de("MAPA_HORARIOS_NOME"))

    # ---- Datas e Competência ----
    arq_lider["mes"] = col("Admissão").apply(extrair_mes)
    arq_lider["ano"] = col("Admissão").apply(extrair_ano)
    arq_lider["mes_ano"] = col("Admissão").apply(extrair_mes_ano)
    arq_lider["dataadmissao"] = col("Admissão").apply(formatar_data)
    arq_lider["cidataemissao"] = col("Data EX").apply(formatar_data)
    arq_lider["cnhdataemissao"] = col("Expedição CNH").apply(formatar_data)
    arq_lider["cnhdatavalidade"] = col("Vencimento").apply(formatar_data)
    arq_lider["datanascimento"] = col("Data nascimento").apply(formatar_data)
    arq_lider["datavenci1diasexp"] = col("Fim Determinado").apply(formatar_data)
    arq_lider["datavenci2diasexp"] = col("Fim Prorrogação").apply(formatar_data)

    # ---- Salário e Tipo de Salário ----
    arq_lider["salariobase"] = col("Salário").apply(formatar_salario)
    arq_lider["tiposalario"] = col("Categoria").apply(normalizar_texto).map(
        mapa_de("MAPA_TIPOSALARIO_NOME")
    )

    # ---- Situação / Desligamento ----
    arq_lider["ativo"] = col("Data Demissão").apply(checar_ativo)
    arq_lider["motivodesligamento"] = clean("Motivo Demissão").map(
        mapa_de("MAPA_DESLIGAMENTO_NOME")
    )
    arq_lider["datademissao"] = col("Data Demissão").apply(formatar_data)

    # ---- Valores fixos (regras do original) ----
    arq_lider["tipoadmissao"] = "CLT"
    arq_lider["tipocontrolepessoa"] = "Funcionário"
    arq_lider["tipofuncionario"] = "Funcionário"
    arq_lider["ctpsdigital"] = "Não"
    arq_lider["nacionalidade_id"] = "1"
    arq_lider["paises_id"] = "1"
    arq_lider["planosaude"] = "Sim"
    arq_lider["planoodonto"] = "Sim"
    arq_lider["possuisegurovida"] = "Sim"
    arq_lider["tipirecebimento"] = "Banco"
    arq_lider["periodo1diasexp"] = "45"
    arq_lider["periodo2diasexp"] = "45"
    arq_lider["tipoendereco"] = "Rua"
    arq_lider["nivelcoberturaposto"] = "Alto"
    arq_lider["databasesindicato"] = f"{agora.year}-01-01"
    arq_lider["datacadastro"] = agora.strftime("%Y-%m-%d")

    # Garante exatamente as colunas/ordem do modelo original
    arq_lider = arq_lider.reindex(columns=colunas_modelo)

    # Pendências (células que devem ficar vermelhas)
    pendencias = []
    pos = {idx: i for i, idx in enumerate(arq_lider.index)}
    for c in CAMPOS_OBRIGATORIOS:
        if c in arq_lider.columns:
            serie = arq_lider[c]
            vazias = serie.isna() | serie.astype(str).str.strip().isin(["", "nan", "None", "NaN"])
            for idx in serie.index[vazias]:
                pendencias.append((pos[idx], c))

    return arq_lider, pendencias
