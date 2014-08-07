rtf2docx
========
Converts rtf to docx

...Captures text in an rtf file including headers and footers, and converts to a docx
...Run with -h or -help flag for more information on how to run

Code is based on docx.py which is downloaded from here:
    https://github.com/mikemaccana/python-docx
and pyth 0.6.0 which is downloaded from here:
    https://pypi.python.org/pypi/pyth/

Installation requires
  apt-get install libxml2-dev libxslt1-dev python-dev
    (this may be required on linux, shouldn't be necessary on later
     model osx systems)

  sudo pip install lxml
  sudo pip install Pillow (used to be PIL)
