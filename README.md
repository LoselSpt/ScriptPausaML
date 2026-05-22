# 🛒 Excluir Anúncios Mercado Livre — Automação

Script de automação para exclusão em massa de anúncios no Mercado Livre Seller Center, usando **Playwright** via conexão CDP ao Chrome do usuário.

## 📋 Funcionalidades

- Lê lista de SKUs de uma planilha Excel (`anuncios_excluir_v1.xlsx`)
- Para cada SKU, navega diretamente para a URL filtrada no Seller Center
- Clica no menu de ações (⋮) de cada anúncio
- Clica em **"Excluir"** no dropdown
- Confirma no modal de confirmação
- Repete até não restar anúncios daquele SKU
- Gera log detalhado da execução

## 🔧 Pré-requisitos

- Python 3.10+
- Google Chrome instalado
- Playwright instalado: `pip install playwright && playwright install chromium`
- openpyxl: `pip install openpyxl`

## 🚀 Como usar

### 1. Abrir o Chrome com depuração remota

Execute o arquivo `abrir_chrome_debug.bat` (duplo clique). Isso abrirá o Chrome com a porta de depuração 9222 ativa.

> **Importante:** Faça login no Mercado Livre Seller Center antes de rodar o script.

### 2. Preparar a planilha

Crie o arquivo `anuncios_excluir_v1.xlsx` com uma coluna chamada `Sku` contendo os SKUs a excluir.

### 3. Executar o script

```bash
python excluir_anuncios_ml.py
```

O script vai:
1. Conectar ao Chrome já aberto via CDP
2. Processar cada SKU da planilha
3. Excluir todos os anúncios encontrados
4. Exibir progresso no terminal

## 📁 Arquivos

| Arquivo | Descrição |
|---|---|
| `excluir_anuncios_ml.py` | Script principal de automação |
| `abrir_chrome_debug.bat` | Abre o Chrome com porta de depuração ativa |
| `orquestrar_exclusoes.py` | Orquestrador alternativo |

## ⚠️ Observações

- O Mercado Livre aplica um **rate limit temporário** após exclusões seguidas. O script pode não conseguir excluir todos os anúncios de um SKU na primeira passada — execute novamente para limpar os restantes.
- Os perfis do Chrome (`chrome_debug_profile/`) não são versionados por conter dados de sessão pessoais.
- As planilhas com dados de SKUs também não são versionadas por conter informações sensíveis.

## 📄 Licença

Uso interno / privado.
