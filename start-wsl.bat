@echo off
echo Starting ContractAI via WSL2...
wsl -e bash -c "cd /mnt/e/CODE/aic2 && node server.js"
pause
