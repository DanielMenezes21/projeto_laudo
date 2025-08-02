# Execute no PowerShell como usuário normal
cd "C:\Users\Usuario\Desktop\Daniel_Menezes\projeto_laudo"
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
$env:TEMP = "C:\Temp"
.\venv\Scripts\Activate.ps1
python -m PyInstaller @(
    '--noconfirm', '--clean', '--onedir',
    '--icon', 'models/icon/DAgro_Logo_03.ico',
    '--name', 'DAgroDocuments',
    '--workpath', 'C:/Temp/py_build',
    '--distpath', 'dist',
    '--add-data', 'app;app',
    '--add-data', 'venv/Lib/site-packages/kivy/data;data',
    '--exclude-module', 'pkgutil',
    'main.py'
)