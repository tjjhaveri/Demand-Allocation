import PyInstaller.__main__
import os

PyInstaller.__main__.run([
    'GUI.py',
    '--onefile',
    #'--collect-data pulp',
    '--icon=./tesla.ico',
])

'''
import pulp from their github(pip install git+https://github.com/coin-or/pulp.git)
will resolve issue on AttributeError: 'NoneType' object has no attribute 'actualSolve'
terminal command:
    'pyinstaller --icon tesla.icns --collect-data pulp --onefile --name 'Allocation Tool' --noconsole GUI.py'
    'pyinstaller -i 'Allocation tool_2.ico' --collect-data pulp --onefile --name 'Single Layer Allocation Tool' --noconsole GUI.py'
'''
