@echo off
REM Activate conda environment and run the script

CALL "User\miniconda3\Scripts\activate.bat" activate base
python "bi_phonereport_update.py"
pause
