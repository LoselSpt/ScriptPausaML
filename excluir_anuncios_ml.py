import os
import sys
import time
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# Garante suporte completo a caracteres UTF-8 no console do Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

EXCEL_PATH = "anuncios_excluir_v1.xlsx"
if not os.path.exists(EXCEL_PATH):
    if os.path.exists("anuncios_excluir.xlsx"):
        EXCEL_PATH = "anuncios_excluir.xlsx"
    else:
        EXCEL_PATH = "skus.xlsx"

AUTH_PATH = "auth.json"
LISTA_URL = "https://www.mercadolivre.com.br/anuncios/lista"

# == Le todos os SKUs da planilha =============================================
def ler_skus(path: str) -> list[str]:
    print(f"[INFO] Lendo planilha: {path}")
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = -1
    for idx, header in enumerate(headers):
        if header.lower() == "sku":
            col = idx
            break
    if col == -1:
        raise ValueError(f"Coluna SKU nao encontrada. Headers: {headers}")
    skus = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > col:
            val = row[col]
            if val:
                skus.append(str(val).strip())
    unique_skus = []
    for s in skus:
        if s not in unique_skus:
            unique_skus.append(s)
    print(f"[PLANILHA] {len(unique_skus)} SKUs unicos encontrados.")
    return unique_skus

# == Exclui todos os anuncios do SKU atual ====================================
def excluir_por_sku(page, sku: str) -> int:
    print(f"\n[SKU {sku}] Buscando anuncios...")

    # Navega diretamente para a URL com o SKU no parâmetro search
    # (mais confiavel que preencher o campo via CDP, que nao aciona o filtro corretamente)
    import urllib.parse
    url_busca = f"https://www.mercadolivre.com.br/anuncios?page=1&sort=DEFAULT&search={urllib.parse.quote(sku)}"
    page.goto(url_busca)
    page.wait_for_load_state("domcontentloaded")
    time.sleep(2.5)  # aguarda a listagem renderizar

    excluidos = 0

    while True:
        botoes = page.get_by_role("button", name="Ações secundárias").all()
        if not botoes:
            botoes = page.get_by_role("button", name="Acciones secundarias").all()
        if not botoes:
            botoes = page.locator("button.sll-list-row-secondary-actions-trigger").all()

        if not botoes:
            print(f"[SKU {sku}] Nenhum anuncio restante.")
            break

        print(f"[SKU {sku}] {len(botoes)} anuncio(s) encontrado(s). Excluindo o primeiro...")

        try:
            botoes[0].scroll_into_view_if_needed()
            botoes[0].click()
            time.sleep(1.2)  # Aguarda o popper abrir

            # Confirma que o popper abriu
            popper = page.locator('[data-testid="popper"]')
            popper.wait_for(state="visible", timeout=6000)

            # Clica via JavaScript no botão Excluir dentro do popper
            # (pode estar com atributo disabled no DOM mas visível — JS click ignora isso)
            clicou = page.evaluate("""() => {
                const popper = document.querySelector('[data-testid="popper"]');
                if (!popper) return false;
                const btns = Array.from(popper.querySelectorAll('button.andes-list__item-action'));
                const excluir = btns.find(b => b.textContent.trim().includes('Excluir') || b.textContent.trim().includes('Eliminar'));
                if (excluir) { excluir.click(); return true; }
                return false;
            }""")

            if not clicou:
                print(f"[SKU {sku}] Botao Excluir nao encontrado no popper.")
                page.keyboard.press("Escape")
                break

            print(f"[SKU {sku}] Clicou em Excluir. Aguardando modal...")
            time.sleep(2.0)  # aguarda o modal de confirmacao aparecer

            # Clica no botão de confirmação via JavaScript
            # (busca o botão primário/loud visível fora do popper = modal de confirmação)
            confirmou = page.evaluate("""() => {
                // Busca botão de confirmação em modais reais (não o popper, não o nav de header)
                // O modal de exclusão usa andes-button--loud ou andes-button--primary
                const candidatos = document.querySelectorAll(
                    '.andes-modal button.andes-button--loud, ' +
                    '.andes-modal button.andes-button--primary, ' +
                    '[data-testid="modal"] button.andes-button--loud, ' +
                    '[data-testid="modal"] button.andes-button--primary'
                );
                for (const btn of candidatos) {
                    if (btn.offsetParent !== null) { // visível
                        btn.click();
                        return 'loud/primary: ' + btn.textContent.trim();
                    }
                }
                // Fallback: qualquer botão visível com texto de confirmação que NÃO seja do popper
                const popper = document.querySelector('[data-testid="popper"]');
                const todos = document.querySelectorAll('button');
                for (const btn of todos) {
                    if (popper && popper.contains(btn)) continue; // ignora botões do popper
                    if (!btn.offsetParent) continue; // ignora hidden
                    const txt = btn.textContent.trim();
                    if (txt === 'Excluir' || txt === 'Confirmar' || txt === 'Eliminar' || txt === 'Sim') {
                        btn.click();
                        return 'fallback: ' + txt;
                    }
                }
                return null;
            }""")

            if confirmou:
                print(f"[SKU {sku}] Confirmacao clicada ({confirmou}).")
            else:
                print(f"[SKU {sku}] Modal de confirmacao nao encontrado - pulando.")
                page.keyboard.press("Escape")
                break

            time.sleep(2.5)  # aguarda o anuncio ser removido

            excluidos += 1
            print(f"[SKU {sku}] OK - Anuncio {excluidos} excluido.")

            # Recarrega a pagina com o filtro do SKU para atualizar a listagem
            page.goto(url_busca)
            page.wait_for_load_state("domcontentloaded")
            time.sleep(2.0)

        except PlaywrightTimeout:
            print(f"[SKU {sku}] TIMEOUT - pulando este anuncio.")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            break
        except Exception as e:
            print(f"[SKU {sku}] ERRO: {e} - pulando.")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            break

    return excluidos

# == Main ======================================================================
def main():
    skus = ler_skus(EXCEL_PATH)
    log_linhas = []

    with sync_playwright() as pw:
        # Conecta diretamente ao Chrome pessoal do usuario via CDP (porta 9222)
        print("\n[CDP] Conectando ao seu Chrome pessoal na porta 9222...")
        browser = pw.chromium.connect_over_cdp("http://localhost:9222")
        print("[CDP] Conectado com sucesso!")

        # Pega o contexto e a pagina ativa (a aba do ML que o usuario ja esta)
        context = browser.contexts[0]
        
        # Procura a aba que esta no ML
        page = None
        for p in context.pages:
            if "mercadolivre.com.br" in p.url:
                page = p
                print(f"[CDP] Usando aba ativa: {p.url}")
                break
        
        if not page:
            # Cria nova aba se nao encontrou nenhuma do ML
            page = context.new_page()
            print("[CDP] Abrindo pagina de anuncios...")
            page.goto(LISTA_URL)
        elif "anuncios" not in page.url:
            # Esta no ML mas nao na pagina de anuncios
            print("[CDP] Navegando para a pagina de anuncios...")
            page.goto(LISTA_URL)

        # Aguarda pagina de anuncios carregar
        print("[INFO] Aguardando pagina de anuncios carregar...")
        page.wait_for_load_state("networkidle", timeout=15000)
        print("[INFO] Pagina carregada! Iniciando exclusoes...\n")

        total = 0
        for i, sku in enumerate(skus, 1):
            print(f"\n=========================================")
            print(f" SKU {i}/{len(skus)}: {sku}")
            print(f"=========================================")
            n = excluir_por_sku(page, sku)
            total += n
            linha = f"SKU {sku}: {n} anuncio(s) excluido(s)"
            log_linhas.append(linha)
            print(f"[RESUMO] {linha}")

        # Nao fecha o browser pois e o Chrome pessoal do usuario!
        print("\n[FIM] Automacao concluida. Chrome nao foi fechado.")

    # Grava log
    with open("log_exclusoes.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(log_linhas))
        f.write(f"\n\nTOTAL EXCLUIDO: {total} anuncio(s)")

    print(f"[FIM] Total excluido: {total} anuncio(s). Log salvo em log_exclusoes.txt")

if __name__ == "__main__":
    main()
