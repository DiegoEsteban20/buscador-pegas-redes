@echo off
echo Buscando ofertas de trabajo de redes y conectividad... (2-4 minutos)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BuscadorPegas.ps1" -Forzar -SiempreAbrir
