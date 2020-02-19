from cx_Freeze import setup, Executable
from time import gmtime, strftime
import re
import os
import sys

packages = ['shlex', 'os', 'lxml']
includes = ['lxml', 'lxml.etree', 'lxml._elementpath', 'lxml.ElementInclude']
includefiles = ['doc.ico']
excludes = []
path = []

# set compilation time as a version
sp = os.path.join(os.path.dirname(os.path.realpath(__file__)), "sp.py")
with open(sp, 'r') as file :
  filedata = file.read()

# Replace the target string
filedata = re.sub(r"(__version__ = )('\d{4}-\d{2}-\d{2}')", r"\1'"+strftime("%Y-%m-%d")+"'", filedata)

# Write the file out again
with open(sp, 'w') as file:
  file.write(filedata)

setup(name='SiemensPie',
      version='2.0.2',
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
