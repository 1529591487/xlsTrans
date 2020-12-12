rmdir build /S /Q 
rmdir dist /S /Q 
::pyinstaller --noconsole --onefile  -p "ui" -p "..\\Public" DocxHandle.py
pyinstaller --onefile  -p "ui" -p "..\\Public" xlsTrans.py