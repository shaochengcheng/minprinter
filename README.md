This is a tool to help print invoices from
China Mobile, China Unicom and China Telcom.


Features:
  - do statstics by searching text from pdf file (using pdftotext)
  - convert pdf into jpg to avoid signed PDF that forbid merging (using pdf2image)
  - place two images on one a4-size pdf page to print
  - a windows exe to run standalone (using Pyinstaller)

Prerequirement:
  - python 3.7 tested
  - python packages see requirements.txt
  - poppler-0.68.0 tested, you can use poppler under this repo OR:
    - download bins from http://blog.alivate.com.au/poppler-windows/
    - download poppler-data from https://poppler.freedesktop.org/
    - put the poppler-data under POPPLER_HOME/share folder, please check the structure of poppler-0.68.0 under this repo

Installation:
  - python setup.py install

To use Pyinstaller:
  - pyinstaller to_exe/minprint.spec (as bundle folder)
  - pyinstaller to_exe/minprint-ALLINONE.spec (as onefile)
  - Or, check [Pyinstaller documents](https://pythonhosted.org/PyInstaller/)

