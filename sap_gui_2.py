# -*- coding: utf-8 -*-
"""Simple example of SAP GUI Scripting via COM Objects and Interfaces
"""


import sys
import win32com.client

materials = ("<my_material_1>", "<my_material_2>", "<my_material_3>")


def main():

    try:

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not isinstance(SapGuiAuto, win32com.client.CDispatch):
            return

        application = SapGuiAuto.GetScriptingEngine
        if not isinstance(application, win32com.client.CDispatch):
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not isinstance(connection, win32com.client.CDispatch):
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not isinstance(session, win32com.client.CDispatch):
            connection = None
            application = None
            SapGuiAuto = None
            return

        for material in materials:
            # --- copied and changed VBS macro
            session.findById("wnd[0]/tbar[0]/okcd").text = "mm03"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02").select()
            document = session.findById(
                "wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMM:2004/subSUB5:SAPLMGD1:2004/txtMARA-ZEINR"
            ).text
            print(document)
            session.findById("wnd[0]/tbar[0]/btn[15]").press()

    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


if __name__ == "__main__":
    main()
