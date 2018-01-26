from cx_Freeze import setup, Executable
import sys

packages = ['shlex', 'os', 'lxml']
includes = ['lxml', 'lxml.etree', 'lxml._elementpath', 'lxml.ElementInclude']
includefiles = ['doc.ico']
excludes = []
path = []

setup(name='SiemensPie',
      version='0.06',
      description='Tool for compile Siemens SIPROTEC relay configuration (.xml and .xrio files) to readable Excel document',
      executables=[
          Executable(
              script='sp.py',
              icon='exe.ico'
          )
      ],
      options={
          "build_exe": {
              "includes": includes,
              "include_files": includefiles,
              "excludes": excludes,
              "packages": packages,
              "path": path
          }
      }
      )
