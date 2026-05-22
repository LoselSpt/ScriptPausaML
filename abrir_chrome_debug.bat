@echo off
echo Fechando Chrome...
taskkill /F /IM chrome.exe 2>nul
timeout /t 2 /nobreak >nul

echo Abrindo Chrome com depuracao remota na porta 9222...
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\joaos\OneDrive\Documentos\ScriptPausa\chrome_debug_profile" --no-first-run --no-default-browser-check "https://www.mercadolivre.com.br/anuncios/lista"
