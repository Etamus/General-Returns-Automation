Attribute VB_Name = "AlterarNFD"
Sub Altera_NFD()

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

Windows("Planilha Reversa.xlsb").Activate
Sheets("Cancelar Ordem").Select
    
nl = Application.WorksheetFunction.CountA(Range("D:D"))
cont = nl + 1

Do While Cells(cont, 1) <> ""
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = Cells(cont, 1).Value
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
aviso = session.findById("wnd[1]/usr/txtMESSTXT1").Text

    If aviso = "Considerar documentos subsequentes" Then
    session.findById("wnd[1]").sendVKey 0
    End If

On Error GoTo 0

session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e" & Cells(cont, 2).Value
session.findById("wnd[0]").sendVKey 0

    If Cells(cont, 3).Value <> "" Then
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0004", "Column1"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0004", "Column1"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0004", "Column1"
    textoAnt = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Cells(cont, 3).Value & " - " & textoAnt
    session.findById("wnd[0]").sendVKey 0
    End If
    
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Cells(cont, 4).Value = "Alterado."
cont = cont + 1

Loop

End Sub


