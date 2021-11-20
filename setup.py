
from cx_Freeze import setup, Executable

files = ['Libro1.xlsx','icon.png','app.ui','FASE1A_1B_conDDA.xlsx']

target = Executable(
      script='main.py',
      base='Win32GUI'
)


setup(name = "MyApp",
      version = "1.2",
      description = "My APP status",
      options = {'build_exe':{'include_files':files}},
      executables = [target]

      )