import sys
from cx_Freeze import setup, Executable

# Arquivo Python que você deseja compilar
target_file = 'layout.py'

# Configurações de compilação
build_options = {
    'packages': ['PySimpleGUI', 'pandas', 'openpyxl'],
    'includes': ['tkinter', 'tkinter.ttk', 'tkinter.simpledialog'],
}

# Criação do executável
executables = [Executable(target_file, base='Win32GUI', target_name ='AppContabilMultiplos.exe')]

# Configuração de setup
setup(
    name='AppContabilMultiplos',
    version='1.0',
    description='Descrição do seu programa',
    options={'build_exe': build_options},
    executables=executables
)