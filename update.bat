@echo off
REM Activate conda environment and run the script

CALL "C:\Users\PAMC-NB-Alpha\miniconda3\Scripts\activate.bat" activate base
python "C:\pam_mis\BI_dashboard_data\bi_phonereport_update.py"
pause
