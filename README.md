# 🧡 Importador Domínio → AppLider (EasyApp) — Líder Limpe

Aplicativo **Streamlit** que automatiza a importação de admissões da **Domínio** para o sistema interno **AppLider/EasyApp**, com visual dark premium nas cores da marca (azul-marinho + laranja).

---

## ✨ Funcionalidades

| Módulo | O que faz |
|---|---|
| 🧹 **Duplicados** | Cruza o Arquivo Domínio com a planilha de colaboradores exportada do sistema interno. Mostra **em vermelho** quem já está cadastrado (por CPF, com fallback por nome) e gera o **Arquivo Domínio LIMPO** (sem duplicados) em `.xlsx` e `.csv`. |
| 1️⃣ **Importar Layout** | Domínio + Mapeamento → `Importar Layout.xlsx` (106 colunas, mesmas do modelo original). Células **sem Posto de Serviço, Função, Escala ou Horário ficam VERMELHAS** na planilha e no preview. |
| 2️⃣ **Benefícios** | Layout + Relação de Benefícios → `Importar Beneficios.xlsx`. `colaborador_id` = matrícula eSocial **ou** **IDs sequenciais** a partir do último ID do EasyApp (com confirmação e mapa visual novo ID × matrícula × nome). |
| 3️⃣ **Converter CSV** | Converte qualquer `.xlsx` em CSV **preservando zeros à esquerda** (CPF, PIS, matrícula, CEP). Os CSVs das etapas 1 e 2 já ficam prontos automaticamente. |
| ⭐ **Executar Tudo** | Pipeline completo: Domínio → (remove duplicados, opcional) → Layout → Benefícios → CSVs, com **Central de Downloads** no final. |

> 📥 **Downloads não reprocessam nada.** Todos os arquivos gerados ficam em memória (session_state). Você pode baixar vários arquivos, em qualquer ordem, sem que a página "refaça" o processamento.

---

## 🚀 Como rodar localmente

```bash
git clone <url-do-repo>
cd app_importador_streamlit
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy no Streamlit Community Cloud

1. Suba este repositório no GitHub (todos os arquivos, incluindo `data/` e `assets/`).
2. Em [share.streamlit.io](https://share.streamlit.io): **New app** → selecione o repositório → arquivo `app.py` → **Deploy**.

## 📁 Estrutura

```
app_importador_streamlit/
├── app.py                      # Aplicação Streamlit (UI + orquestração)
├── requirements.txt
├── .streamlit/config.toml      # Tema dark (laranja/azul-marinho)
├── assets/logo.png             # Logo Líder Limpe
├── core/
│   ├── excel_io.py             # Leitura .xls/.xlsx/HTML + CSV preservando texto
│   ├── layout_engine.py        # Regras do Importar Layout (port 1:1 do original)
│   ├── benefits_engine.py      # Regras de benefícios + IDs sequenciais
│   ├── duplicates.py           # Detecção de duplicados (CPF/nome)
│   └── exporters.py            # XLSX com células vermelhas/azuis
└── data/                       # Arquivos base padrão (podem ser substituídos na barra lateral)
    ├── Mapeamento Sistema.xls
    ├── Colunas Originais.xlsx
    └── BENEFICIOS - RELACAO.xlsx
```

---

## 📋 Regras de negócio preservadas (Importar Layout)

- **Escala 12x36 P/I**: paridade pelo dia da admissão; inverte em meses terminados em 31/29.
- **5x2 / 6x1**: detecção por padrão na jornada (`5X2`, `05X02`, `6X1`...).
- **Salário**: remove 4 dígitos finais e posiciona decimais (`1682700000` → `1682.70`).
- **Documentos**: CPF com 11 dígitos (zeros à esquerda), PIS/CTPS/título sem `.0`.
- **Datas**: formato `YYYY-MM-DD`; mês/ano e `MM/AAAA` extraídos da admissão.
- **Ativo**: "Sim" se demissão vazia ou contendo "trabalhando".
- **Supervisor**: funções com ID 11, 12, 13, 14 ou 42.
- **PCD**: Física/Visual/Auditiva/Intelectual/Psicossocial, "Múltipla" quando >1.
- **Escolaridade**: usa Situação, exceto quando vazia/"trabalhando" → Grau de instrução.
- **Mapeamentos**: empresas, postos, funções, cidades (+município), horários (padrão `NNXNN_HH:MM_...`), escalas, raça, estado civil, escolaridade, sindicatos, desligamento, tipo de salário.
- **Valores fixos**: CLT, Funcionário, plano saúde/odonto/seguro = Sim, 45+45 dias experiência, tipo endereço Rua, cobertura Alto, `databasesindicato` = 01/01 do ano corrente.

## 🎁 Regras de benefícios preservadas

- Vigilante → benefício 56; posto "VALE" → 136.
- 12x36 / 5x2-6x1 (≥6h → 6H, senão 4H), TB2 (Guidoni/Cesan/VPorts/Britânia) vs TB1.
- Extras: Petrolina (53), Ambiental Vale + Cozinheiro/Ajudante (51), Motorista (41), VPorts (55), Guarda-Vidas (48), Lavador Veículos Pesado Ambiental Vale (137), Ambiental/Instituto Ambiental Vale (136).
- Overrides: 136 = R$ 1.106,70; 137 = R$ 154,49 (qtd 1).
- Excel final: **azul** = colaborador com mais de um benefício; **vermelho** = benefício único.

## 🔢 IDs sequenciais (novo)

Na aba Benefícios (ou Executar Tudo), escolha **"IDs sequenciais"**: informe o último ID cadastrado no EasyApp na barra lateral, confirme, e o app gera `último+1, +2, ...` na ordem das matrículas — com tabela de conferência novo ID × matrícula eSocial × nome. O último ID é atualizado automaticamente após a geração.

## ⚠️ Observações

- Arquivos `.xls` exportados como HTML pela Domínio/EasyApp são lidos automaticamente.
- Os arquivos base podem ser substituídos a qualquer momento pela barra lateral (sem editar código).
- O "último ID" fica salvo apenas durante a sessão do app.
