@echo off
echo Dang tien hanh dong goi BOM Processor...

:: Cai dat pyinstaller neu chua co
pip install pyinstaller customtkinter pandas openpyxl

:: Dong goi bang pyinstaller (dung onedir de ho tro luu file JSON)
pyinstaller --noconfirm --onedir --windowed --icon "icon.ico" --add-data "attrition_rules.json;." --add-data "component_dictionary.json;." --add-data "_active_dictionary.json;." --add-data "_active_rules.json;." --collect-all "customtkinter" "main.py" --name "BOM-Processor"

echo.
echo =======================================================
echo DONG GOI HOAN TAT!
echo Ung dung cua ban nam trong thu muc: dist\BOM-Processor
echo Ban chi can copy hoac nen file ZIP thu muc "BOM-Processor" nay va gui cho nguoi khac.
echo =======================================================
pause
