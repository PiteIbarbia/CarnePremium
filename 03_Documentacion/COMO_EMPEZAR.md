# Cómo empezar

En PowerShell:

```powershell
cd "C:\Users\Pite\Documents\Trabajo\Carne Premium-Vacío\CarnePremiumERP"
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\pip.exe install -r .\04_Desarrollo\requirements.txt
.\.venv\Scripts\python.exe .\04_Desarrollo\Scripts\builder.py
.\.venv\Scripts\python.exe .\04_Desarrollo\Pruebas\test_mvp.py
```

Después abrí:
`01_ERP\CarnePremiumERP.xlsx`

No borres ni renombres las hojas del proyecto.
