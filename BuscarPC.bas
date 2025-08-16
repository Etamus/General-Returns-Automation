Attribute VB_Name = "BuscarPC"
Sub busca_pc()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Cancelar Ordem").Select

nl = Application.WorksheetFunction.CountA(Range("D:D"))
cont = 1 + nl

Do While Cells(cont, 1) <> ""

If Not IsObject(App) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set App = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = App.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject App, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radP_PDF").Select
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = Cells(cont, 1).Value
session.findById("wnd[0]/usr/ctxtP_ARQ").Text = "c:\Temp\" & Cells(cont, 1).Value & ".pdf"
session.findById("wnd[0]/tbar[1]/btn[8]").press
On Error Resume Next
Cells(cont, 4).Value = session.findById("wnd[1]/usr/txtMESSTXT1").Text
teste = session.findById("wnd[0]").Text
If teste = "Relatório Pré-Cálculo Despesa Logística Reversa" Then
Cells(cont, 4).Value = "Documento de transporte não existe"
teste = ""
session.findById("wnd[0]/tbar[1]/btn[8]").press
GoTo pula
End If
On Error GoTo 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
pula:
cont = cont + 1
Loop

End Sub
