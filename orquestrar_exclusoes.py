import os
import sys
import subprocess
import re
import time
from datetime import datetime

# Configurações de diretórios e caminhos
SCRIPT_DIR = r"C:\Users\joaos\OneDrive\Documentos\ScriptPausa"
SCRIPT_PATH = os.path.join(SCRIPT_DIR, "excluir_anuncios_ml.py")
DRIVE_DIR = r"G:\Meu Drive\LoselNotebook"
MAX_RODADAS = 8

# Garante suporte completo a caracteres UTF-8 no console do Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

def log_local(mensagem: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {mensagem}", flush=True)

def executar_rodada(numero_rodada: int) -> tuple[bool, int, int, int, str]:
    """
    Executa o script de exclusão, captura o output em tempo real e analisa os resultados.
    Retorna (sucesso, total_excluidos, total_erros, total_timeouts, relatorio_md)
    """
    log_local(f"Iniciando Rodada {numero_rodada}...")
    
    # Inicia o processo python de exclusão
    process = subprocess.Popen(
        [sys.executable, SCRIPT_PATH],
        cwd=SCRIPT_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    stdout_lines = []
    
    # Lê a saída em tempo real e imprime no console local
    while True:
        line = process.stdout.readline()
        if not line and process.poll() is not None:
            break
        if line:
            sys.stdout.write(line)
            sys.stdout.flush()
            stdout_lines.append(line.strip())

    process.wait()
    exit_code = process.returncode
    log_local(f"Rodada {numero_rodada} finalizada com código de saída: {exit_code}")

    full_stdout = "\n".join(stdout_lines)
    
    # Variáveis de métricas da rodada
    total_skus = 0
    skus_com_exclusao = {}
    skus_com_erro = {}
    skus_com_timeout = set()
    total_excluidos = 0

    # Padrões Regex para análise do log
    regex_skus_planilha = re.compile(r"\[PLANILHA\] (\d+) SKUs unicos encontrados")
    regex_resumo_sku = re.compile(r"\[RESUMO\] SKU (.+?): (\d+) anuncio\(s\) excluido\(s\)")
    regex_erro_sku = re.compile(r"\[SKU (.+?)\] ERRO: (.+)")
    regex_timeout_sku = re.compile(r"\[SKU (.+?)\] TIMEOUT")

    # Analisa linha por linha
    for line in stdout_lines:
        # Detecta quantidade inicial de SKUs
        match_planilha = regex_skus_planilha.search(line)
        if match_planilha:
            total_skus = int(match_planilha.group(1))
            continue
            
        # Detecta exclusões
        match_resumo = regex_resumo_sku.search(line)
        if match_resumo:
            sku = match_resumo.group(1).strip()
            qtd = int(match_resumo.group(2))
            if qtd > 0:
                skus_com_exclusao[sku] = qtd
                total_excluidos += qtd
            continue

        # Detecta erros
        match_erro = regex_erro_sku.search(line)
        if match_erro:
            sku = match_erro.group(1).strip()
            erro_msg = match_erro.group(2).strip()
            skus_com_erro[sku] = erro_msg
            continue

        # Detecta timeouts
        match_timeout = regex_timeout_sku.search(line)
        if match_timeout:
            sku = match_timeout.group(1).strip()
            skus_com_timeout.add(sku)
            continue

    total_erros = len(skus_com_erro)
    total_timeouts = len(skus_com_timeout)
    
    # Define o status da rodada
    status_rodada = "SUCESSO COMPLETO" if (total_excluidos == 0 and total_erros == 0 and total_timeouts == 0) else "RODANDO RE-TENTATIVAS"
    if exit_code != 0:
        status_rodada = "FALHA (CÓDIGO DE ERRO)"

    # Gera o conteúdo formatado em Markdown
    timestamp_fim = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    md_lines = [
        f"# Relatório de Execução - Rodada {numero_rodada}",
        f"**Data/Hora de Conclusão:** {timestamp_fim}  ",
        f"**Status da Rodada:** `{status_rodada}`",
        "",
        "## Resumo das Métricas",
        "| Métrica | Valor |",
        "| :--- | :--- |",
        f"| **SKUs Totais Carregados** | {total_skus} |",
        f"| **Total de Anúncios Excluídos** | {total_excluidos} |",
        f"| **SKUs com Timeout** | {total_timeouts} |",
        f"| **SKUs com Outros Erros** | {total_erros} |",
        f"| **Código de Saída do Script** | {exit_code} |",
        ""
    ]

    # Adiciona detalhes de SKUs com Exclusões
    if skus_com_exclusao:
        md_lines.append("## Detalhamento de Exclusões por SKU")
        md_lines.append("| SKU | Anúncios Excluídos |")
        md_lines.append("| :--- | :--- |")
        for sku, qtd in skus_com_exclusao.items():
            md_lines.append(f"| `{sku}` | {qtd} |")
        md_lines.append("")

    # Adiciona detalhes de SKUs com Problemas (Erros ou Timeouts)
    if skus_com_erro or skus_com_timeout:
        md_lines.append("## Alertas e Problemas Identificados")
        md_lines.append("| SKU | Tipo de Ocorrência | Detalhe / Mensagem |")
        md_lines.append("| :--- | :--- | :--- |")
        for sku in skus_com_timeout:
            md_lines.append(f"| `{sku}` | `TIMEOUT` | A página demorou a responder durante o fluxo de exclusão. |")
        for sku, erro_msg in skus_com_erro.items():
            md_lines.append(f"| `{sku}` | `ERRO` | {erro_msg} |")
        md_lines.append("")

    # Adiciona o Log bruto da rodada em aba retrátil
    md_lines.append("## Log Técnico da Rodada")
    md_lines.append("<details>")
    md_lines.append("<summary>Clique aqui para visualizar o Log de Saída completo</summary>")
    md_lines.append("")
    md_lines.append("```text")
    md_lines.append(full_stdout)
    md_lines.append("```")
    md_lines.append("</details>")
    md_lines.append("")

    relatorio_md = "\n".join(md_lines)
    return (exit_code == 0, total_excluidos, total_erros, total_timeouts, relatorio_md)

def main():
    log_local("=== INICIANDO INTEGRAÇÃO DO ORQUESTRADOR ===")
    
    if not os.path.exists(DRIVE_DIR):
        log_local(f"ERRO: A pasta do Google Drive '{DRIVE_DIR}' não foi encontrada.")
        sys.exit(1)

    rodada = 1
    total_acumulado_excluidos = 0

    while rodada <= MAX_RODADAS:
        log_local(f"\n--- RODADA {rodada} DE {MAX_RODADAS} ---")
        
        # Executa
        sucesso, excluidos, erros, timeouts, relatorio = executar_rodada(rodada)
        total_acumulado_excluidos += excluidos

        # Salva o arquivo markdown diretamente no Google Drive
        timestamp_slug = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Relatorio_Rodada_{rodada}_{timestamp_slug}.md"
        caminho_drive = os.path.join(DRIVE_DIR, nome_arquivo)
        
        try:
            with open(caminho_drive, "w", encoding="utf-8") as f:
                f.write(relatorio)
            log_local(f"Relatório salvo com sucesso em: {caminho_drive}")
        except Exception as e:
            log_local(f"ERRO ao salvar relatório no Google Drive: {e}")

        # Decisão de Loop
        # Paramos se a rodada atual for 100% limpa (0 anúncios excluídos E 0 erros E 0 timeouts)
        if excluidos == 0 and erros == 0 and timeouts == 0:
            log_local("SUCESSO: Rodada 100% limpa atingida (0 anúncios encontrados/restantes em todos os SKUs).")
            
            # Gera relatório consolidado final
            nome_final = f"Relatorio_Final_CONCLUIDO_{timestamp_slug}.md"
            caminho_final = os.path.join(DRIVE_DIR, nome_final)
            conteudo_final = f"""# Relatório Consolidado de Conclusão

A automação foi concluída com sucesso absoluto! Todos os anúncios pendentes de todos os SKUs foram completamente removidos da base.

- **Total de Rodadas Realizadas:** {rodada}
- **Total Acumulado de Anúncios Excluídos:** {total_acumulado_excluidos}
- **Data/Hora de Encerramento:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

Todos os SKUs possuem agora 0 anúncios ativos.
"""
            try:
                with open(caminho_final, "w", encoding="utf-8") as f:
                    f.write(conteudo_final)
                log_local(f"Relatório consolidado de encerramento salvo em: {caminho_final}")
            except Exception as e:
                log_local(f"Erro ao salvar relatório final consolidado: {e}")
            break
        
        # Se restaram anúncios ou ocorreram erros/timeouts, faz nova tentativa
        log_local(f"Ainda há pendências (Excluídos nesta rodada: {excluidos}, Erros: {erros}, Timeouts: {timeouts}).")
        
        if rodada == MAX_RODADAS:
            log_local("AVISO: Limite máximo de rodadas atingido. Encerrando orquestração.")
            break
            
        log_local("Aguardando 15 segundos para estabilização da página do Mercado Livre e do Chrome antes da próxima rodada...")
        time.sleep(15)
        rodada += 1

    log_local("\n=== ORQUESTRADOR FINALIZADO ===")

if __name__ == "__main__":
    main()
