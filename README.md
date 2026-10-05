# 🟧 LiderLimp · Importador de Admissões (Streamlit) — v4

App web interno para importação de admissões: lê a exportação do **Domínio**,
gera a planilha de importação do **AppLider / EasyApp**, gera a planilha de
**benefícios** (VA/CB/CF/VT) com as regras de negócio do RH e converte tudo em **CSV**.

**Sem login** — app aberto para uso interno da equipe.

---

## 🆕 Novidades da v4 (set/2026)

- **Login removido** — o app abre direto, sem tela de acesso.
- **Cores do Excel de benefícios revisadas**:
  - 🔵 **Azul** → colaborador com **mais de um** benefício VA/CB/CF;
  - 🔴 **Vermelho** → colaborador com **apenas um** VA/CB/CF;
  - 🟢 **Verde** → linhas de **VT** (vale-transporte), sempre.
  - Antes um colaborador com VA + VT aparecia azul (como se tivesse 2 VAs);
    agora fica vermelho (1 VA) e a linha do VT fica verde.
- **Mapeamento Sistema também via Google Sheets** ☁️ — além da planilha de
  benefícios, o mapeamento (12 abas) também pode ser lido do Google.

---

## ☁️ COMO IMPLEMENTAR O GOOGLE SHEETS (passo a passo)

### 1) As duas planilhas já estão no seu Google Drive

| Planilha | Link |
|---|---|
| 🎁 BENEFICIOS - RELAÇÃO (LiderLimp) | https://docs.google.com/spreadsheets/d/149DL4ip25O8IM3w7LS08tXQJfxdXoSdSQkrI2kMhV2c/edit |
| 🗺️ Mapeamento Sistema (LiderLimp) | https://docs.google.com/spreadsheets/d/1Rn3pYB9uHQysMMfNGgHccJPALDMqf9f4sOYk7MyZaxc/edit |

### 2) Compartilhe as DUAS planilhas (1 clique em cada)

Em cada planilha: **Compartilhar → "Qualquer pessoa com o link" → Leitor**.
Isso permite o app ler os dados. (Quem tiver o link só *lê*; ninguém edita.)

### 3) Configure os Secrets no Streamlit Cloud

No painel do app → **Settings → Secrets**, cole:

```toml
[gsheets]
beneficios_sheet_id  = "149DL4ip25O8IM3w7LS08tXQJfxdXoSdSQkrI2kMhV2c"
mapeamento_sheet_id  = "1Rn3pYB9uHQysMMfNGgHccJPALDMqf9f4sOYk7MyZaxc"
```

> Mesmo sem esse passo, o app usa os IDs padrão já embutidos no código —
> os secrets servem para você poder trocar de planilha no futuro.

### 4) Suba o código v4 no GitHub e faça o deploy

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### 5) Pronto!

Na barra lateral do app, deixe ligado:
- ☁️ **Mapeamento via Google Sheets**
- ☁️ **Usar Google Sheets (Benefícios)**

Se a internet/Google falhar, o app cai automaticamente para o arquivo local
em `data/` (ou para um upload manual) — **ele nunca para por falta de conexão**.

### 🔒 Modo robusto (opcional — planilhas privadas)

Se preferir NÃO deixar as planilhas públicas por link, crie uma **conta de
serviço** no Google Cloud, compartilhe as duas planilhas com o e-mail dela e
adicione nos secrets:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "....iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```

---

## Recursos das v2/v3 (mantidos)

- **Layout fixo no código** (106 colunas) — *Colunas Originais* aposentada.
- **`beneficioemconta`** (Sim/Não pela cidade) · `cpf` → `cnpj_cpf`.
- **Posto ADM por empresa** (VSP, Ativa, Multisserviços, Lider Limpe, B2WE).
- **Menor aprendiz** (<18 anos → função 76, sem VA; extras e VT mantidos).
- **Benefícios v2** — VT 58 (CETURB) / 59 (Grande Vitória-ES), VA motorista (40),
  VA supervisor (63), VA/coffee VPORTS+MULTILIFT (54/55), TB2 ampliado,
  horários noturnos corrigidos.
- **Aba LISTAS editável** (postos TB2, cidades GV, empresa→ADM etc.) —
  editou na planilha, o app já usa.

## 🗂️ Arquivos de configuração

| Fonte | Obrigatório | Observação |
|---|---|---|
| Google Sheets *BENEFICIOS - RELAÇÃO* | ✅ (padrão) | lido online a cada execução |
| Google Sheets *Mapeamento Sistema* | ✅ (padrão) | lido online a cada execução |
| `data/Mapeamento Sistema.xls` | fallback | usado se o Google estiver fora |
| `data/BENEFICIOS - RELACAO.xlsx` | fallback | usado se o Google estiver fora |
