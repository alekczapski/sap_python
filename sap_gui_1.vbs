'Macro recorded with SAP "Script Recording and Playback..."

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 152,30,false
session.findById("wnd[0]/tbar[0]/okcd").text = "mm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "<my_material>"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMM:2004/subSUB5:SAPLMGD1:2004/txtMARA-ZEINR").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMM:2004/subSUB5:SAPLMGD1:2004/txtMARA-ZEINR").caretPosition = 12
session.findById("wnd[0]/tbar[0]/btn[15]").press
