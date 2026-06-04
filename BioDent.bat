@echo off
for /f "delims=" %%p in ('python -c "import sys,os; pw=os.path.join(os.path.dirname(sys.executable),'pythonw.exe'); print(pw if os.path.exists(pw) else sys.executable)"') do set PYEXE=%%p
start "" "%PYEXE%" "%~dp0BioDent_Main.py"
