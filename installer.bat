@echo off
color 3
echo Installing required Packs
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo Installer done bye!
timeout /t 2
exit