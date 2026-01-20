@echo off
echo Redémarrage du serveur avec support HEIC...
echo.
echo Appuyez sur Ctrl+C pour arrêter le serveur en cours
echo puis lancez ce script à nouveau.
echo.
pause
echo.
echo Démarrage du serveur...
uvicorn app:app --reload --host 127.0.0.1 --port 8000
