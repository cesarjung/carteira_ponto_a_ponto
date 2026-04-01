# Pipeline GPM

Automação em Python executada via **GitHub Actions** para gerar e atualizar arquivos CSV no Google Drive a partir de dados do Google Sheets relacionados ao fluxo GPM.

O pipeline executa 3 etapas em sequência:

1. Gera o CSV de obras/projetos
2. Gera o CSV de atividades por projeto
3. Gera o CSV de materiais por projeto

Tudo isso com autenticação por **Service Account do Google**, podendo rodar tanto no **GitHub Actions** quanto localmente.

---

## Objetivo

Este repositório centraliza a geração de arquivos modelo utilizados no processo de importação e consolidação do GPM, garantindo que os arquivos:

- sejam gerados automaticamente em horários programados
- sejam atualizados diretamente no Google Drive
- não criem arquivos duplicados desnecessários
- mantenham compatibilidade com Excel por meio de `utf-8-sig`
- possam ser executados localmente ou no GitHub

---

## Fluxo do Pipeline

A automação segue esta ordem:

### 1. `import_gpm.py`
Gera o arquivo:

- `modelo_importar_obras_ponto.csv`

Esse script:
- lê dados da **Carteira**
- cruza com **BD_Obras_GPM**
- consulta **ATIVIDADES_POR_PONTO_BASE**
- aplica regras de exclusão
- normaliza chaves de configuração
- monta o CSV em memória
- envia diretamente para a pasta do Google Drive

---

### 2. `import_gpm_atividades.py`
Gera o arquivo:

- `modelo_imp_ativ_obra_lote.csv`

Esse script:
- usa como base a lista de projetos do arquivo `modelo_importar_obras_ponto`
- busca as planilhas de origem listadas em `BD_Config!A3:A`
- consulta a aba `ATIVIDADES_POR_PONTO`
- extrai:
  - **Projeto**
  - **Atividade**
  - **Quantidade**
- sobrescreve o arquivo no Google Drive, sem criar cópias duplicadas

> O script aguarda 5 minutos após a geração do primeiro CSV para garantir que o arquivo base já esteja disponível no Drive antes da leitura.

---

### 3. `import_gpm_materiais.py`
Gera o arquivo:

- `modelo_materiais_lote.csv`

Esse script:
- usa como base a lista de projetos do arquivo `modelo_importar_obras_ponto`
- busca as planilhas de origem listadas em `BD_Config!A3:A`
- consulta a aba `MATERIAIS_POR_PONTO`
- extrai:
  - **Projeto**
  - **Material**
  - **Quantidade**
- sobrescreve o arquivo no Google Drive, sem criar cópias duplicadas

---

## Agendamento

O workflow está configurado para rodar automaticamente nos seguintes horários:

- **07:00 BRT**
- **12:00 BRT**
- **15:00 BRT**

No GitHub Actions isso foi configurado em UTC como:

- `0 10 * * *`
- `0 15 * * *`
- `0 18 * * *`

Também é possível executar manualmente via:

- `workflow_dispatch`

---

## Estrutura esperada do projeto

```bash
.
├── .github/
│   └── workflows/
│       └── pipeline_gpm.yml
├── import_gpm.py
├── import_gpm_atividades.py
├── import_gpm_materiais.py
├── requirements.txt
├── credenciais.json   # opcional para execução local
└── README.md
