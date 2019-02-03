# SAP and Python (SAP GUI Scripting, PyRFC)

This is a collection of my notes and code snippets.  
I work with SAP on Windows 10 with SAP GUI for Windows 7.50.  
On my work account I don't have administrative rights and access to SAP NetWeaver RFC SDK.

If you have questions, please contact me:
* Email: {github account}@gmail (dot) com
* LinkedIn: https://www.linkedin.com/in/aleksander-czapski-437b1b162/

## Table of contents

  * [Prequisities](#prequisities)
     * [Extracting Python 2.7 from MSI Windows Installer](#extracting-python-27-from-msi-windows-installer)
     * [Installing pip, setuptools, pywin32 (for SAP GUI Scripting)](#installing-pip-setuptools-pywin32-for-sap-gui-scripting)
     * [Installing PyRFC for Python 2.7 (for SAP RFC)](#installing-pyrfc-for-python-27-for-sap-rfc)
     * [Adding Python and _sapnwrfc.dll_ to PATH variable](#adding-python-and-sapnwrfcdll-to-path-variable)
     * [Related links](#related-links)
  * [SAP GUI Scripting](#sap-gui-scripting)
     * [Related links](#related-links-1)
  * [PyRFC](#parameters)
     * [Related links](#related-links-2)
  * [Similar topics](#similar-topics)
     * [Testing with Robot Framework and SapGui Library](#testing-with-robot-framework-and-sapgui-library)  
     * [Related links](#related-links-3)

## Prequisities

Because I don't have access to SAP NetWeaver RFC SDK, I am using a 32-bit _sapnwrfc.dll_ library provided with SAP GUI installation, which forces me to use a 32-bit version of PyRFC. PyRFC in compiled 32-bit version is only available for Python 2.7.

### Extracting Python 2.7 from MSI Windows Installer
I install this way because a normal installation requires admin privileges.
```dos
msiexec /a python-2.7.15.msi /qb TARGETDIR=C:\Python27
```

### Installing pip, setuptools, pywin32 (for SAP GUI Scripting)
```dos
C:\Python27\python.exe -m ensurepip --default-pip
C:\Python27\python.exe -m pip install --upgrade pip setuptools
C:\Python27\python.exe -m pip install --upgrade pywin32
```

### Installing PyRFC for Python 2.7 (for SAP RFC)
```dos
C:\Python27\Scripts\easy_install.exe https://github.com/SAP/PyRFC/raw/master/dist/pyrfc-1.9.3-py2.7-win32.egg
```

### Adding Python and _sapnwrfc.dll_ to PATH variable
1. Press `Windows logo key + R` and run `rundll32 sysdm.cpl,EditEnvironmentVariables`.  
1. Add `C:\Python27`, `C:\Python27\Scripts` and `C:\Windows\SysWOW64` to `PATH` variable.

### Related links
https://blogs.sap.com/2017/05/15/sap-gui-for-windows-7.50-offers-sapnwrfc-library-version-7.49-pl-0/

## SAP GUI Scripting

### Related links
https://blogs.sap.com/2017/09/19/how-to-use-sap-gui-scripting-inside-python-programming-language/  
https://blogs.sap.com/2014/11/20/scripting-tracker-development-tool-for-sap-gui-scripting/  
https://www.stschnell.de/  
https://sappython.blogspot.com/  
https://stackoverflow.com/questions/26764978/using-win32com-with-multithreading  
https://mail.python.org/pipermail/python-win32/2008-June/007788.html

## PyRFC

### Related links
http://wbarczynski.pl/calling-bapis-with-python-and-pyrfc/  
https://blogs.sap.com/2016/02/21/how-to-use-actual-sap-netweaver-rfc-library-with-python-call-abap-report/

## Similar topics

### Testing with Robot Framework and SapGui Library

### Related links
https://pypi.org/project/robotframework-sapguilibrary/  
https://frankvanderkuur.github.io/SapGuiLibrary.html
