@echo off
title Trans Pinho - Painel de Escalas (escala.phaelz.com)
color 0b

echo =========================================================
echo             INICIANDO PAINEL DE ESCALAS
echo =========================================================
echo.

:: 1. Start the Node.js Server in the background
echo [+] Iniciando o servidor Node.js na porta 3000...
start "Servidor Node.js" /min node server.js

:: Wait a moment for Node to boot
timeout /t 3 /nobreak >nul

:: 2. Start the Cloudflare Tunnel
echo [+] Estabelecendo conexao segura para escala.phaelz.com...
echo [+] O painel estara disponivel em: https://escala.phaelz.com
echo.
echo Pressione Ctrl+C para encerrar tudo.
echo =========================================================
echo.

"C:\Users\opera\AppData\Local\Microsoft\WinGet\Packages\Cloudflare.cloudflared_Microsoft.Winget.Source_8wekyb3d8bbwe\cloudflared.exe" tunnel run --url http://localhost:3000 escala-whats

pause
