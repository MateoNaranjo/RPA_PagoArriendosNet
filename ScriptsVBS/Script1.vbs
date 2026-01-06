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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtEL_LIFNR-LOW").text = "9000"
session.findById("wnd[0]/usr/ctxtLISTU").text = "alv"
session.findById("wnd[0]/usr/ctxtLISTU").setFocus
session.findById("wnd[0]/usr/ctxtLISTU").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[8]").press
