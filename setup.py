from cx_Freeze import setup, Executable
import sys

packages = ['win32com.server', 'win32com.client', 'shlex', 'os', 'lxml']
includes = ['lxml', 'lxml.etree', 'lxml._elementpath', 'lxml.ElementInclude']
excludes = []
path = []

setup(name='SiemensPie',
      version='0.06',
      description='Tool for compile Siemens SIPROTEC relay configuration (.xml and .xrio files) to readable Excel document',
      executables=[
          Executable(
              script='sp.py',
              icon='pict.ico'
          )
      ],
      options={
          "build_exe": {
              "includes": includes,
              "excludes": excludes,
              "packages": packages,
              "path": path,
              "compressed": True,
              "optimize": 2
          }
      }
      )
