SAP and Python (SAP GUI Scripting, PyRFC)
=========================================

...

Hints 
-----
1. Extracting Python 2.7 from MSI Windows Installer 

   ::

     msiexec /a python-2.7.14.msi /qb TARGETDIR=C:\Python27

2. Installing pip, setuptools, pywin32

   ::

     C:\Python27\python.exe -m ensurepip --default-pip
     C:\Python27\python.exe -m pip install --upgrade pip setuptools
     C:\Python27\python.exe -m pip install --upgrade pywin32

3. Installing PyRFC for Python 2.7 (win32)

   ::

     C:\Python27\Scripts\easy_install.exe https://github.com/SAP/PyRFC/raw/master/dist/pyrfc-1.9.3-py2.7-win32.egg

4. Setting PATH variable for Python and sapnwrfc.dll (Windows 7 non-administrative users)

   ::

     SETX PATH C:\Python27;C:\Python27\Scripts;C:\Windows\SysWOW64;

Related
-------

- SAP GUI Scripting with Python

  - https://blogs.sap.com/2017/09/19/how-to-use-sap-gui-scripting-inside-python-programming-language/
  - https://sappython.blogspot.com/
  - https://www.stschnell.de/
  - https://stackoverflow.com/questions/26764978/using-win32com-with-multithreading
  - https://mail.python.org/pipermail/python-win32/2008-June/007788.html

- PyRFC

  - http://wbarczynski.pl/calling-bapis-with-python-and-pyrfc/
  - https://blogs.sap.com/2017/05/15/sap-gui-for-windows-7.50-offers-sapnwrfc-library-version-7.49-pl-0/
  - https://blogs.sap.com/2016/02/21/how-to-use-actual-sap-netweaver-rfc-library-with-python-call-abap-report/
