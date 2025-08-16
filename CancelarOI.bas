Attribute VB_Name = "CancelarOI"
Sub Cancelar_OI()

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0

    Sheets("Cancelar Ordem").Select
    Final = Range("D100000").End(xlUp).Row + 1
        
    Range("A" & Final).Select
    

nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

ordem = ActiveCell.Offset(0, 0).Value
motivo = ActiveCell.Offset(0, 1).Value
texto = ActiveCell.Offset(0, 2).Value
vazio = ActiveCell.Offset(0, 0).Value
aviso = ""

If vazio = "" Then
GoTo fim
End If


session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ordem
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
aviso = session.findById("wnd[0]/sbar").Text
If aviso = "O documento SD " & ordem & " não existe no banco de dados ou não foi arquivado" Then
Windows("Criação.xlsb").Activate
    Sheets("Cancelar").Select
ActiveCell.Offset(0, 3).Value = "OI não existe"
GoTo nexiste
End If
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
On Error GoTo 0

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = "160"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").SetFocus

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").Select

volta555:

On Error Resume Next
barra = session.findById("wnd[0]/sbar").Text
On Error GoTo 0
teste = Right(barra, 7)
If Right(barra, 7) = "consumo" Then
session.findById("wnd[0]").sendVKey 0
barra = ""
GoTo volta555
End If


session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/usr/cmbRV45A-S_ABGRU").Key = "60"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[7]").press

On Error Resume Next
barra = session.findById("wnd[0]/sbar").Text
If barra = "Não foi efetuada qualquer modificação de dados" Then
On Error GoTo 0
GoTo texto
End If

On Error Resume Next
volta:
txtd = session.findById("wnd[2]/usr/txtMESSTXT1").Text
If txtd <> "" Then
session.findById("wnd[0]").sendVKey 0
txtd = ""
GoTo volta
On Error GoTo 0
End If

texto:

session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
nfxx = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text
If nfxx = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e1-1"
End If
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0005", "Column1"
textoAnt = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = texto & " - " & textoAnt
session.findById("wnd[0]/tbar[0]/btn[3]").press

On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Windows("Planilha Reversa.xlsb").Activate
    Sheets("Cancelar Ordem").Select

ActiveCell.Offset(0, 3).Value = "Cancelada."
nexiste:
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation
'
End Sub

Sub Zerar_OI()

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0

    Sheets("Cancelar Ordem").Select
    Final = Range("D100000").End(xlUp).Row + 1
        
    Range("A" & Final).Select
    
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

ordem = ActiveCell.Offset(0, 0).Value
motivo = ActiveCell.Offset(0, 1).Value
texto = ActiveCell.Offset(0, 2).Value
vazio = ActiveCell.Offset(0, 0).Value
aviso = ""
If vazio = "" Then
GoTo fim
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ordem
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
session.findById("wnd[1]").sendVKey 0
On Error GoTo 0
contar = 0
volta:
linha = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text
If linha <> "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = ""
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.Position = contar + 1
contar = contar + 1
GoTo volta
End If

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = "160"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").Select


session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/usr/cmbRV45A-S_ABGRU").Key = "60"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[7]").press

On Error Resume Next
barra = session.findById("wnd[0]/sbar").Text
If barra = "Não foi efetuada qualquer modificação de dados" Then
On Error GoTo 0
GoTo texto
End If
On Error Resume Next
volta1:
txtd = session.findById("wnd[2]/usr/txtMESSTXT1").Text
If txtd <> "" Then
session.findById("wnd[0]").sendVKey 0
txtd = ""
GoTo volta1
On Error GoTo 0
End If

texto:

session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
nfxx = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text
If nfxx = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = "e1-1"
End If
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0005", "Column1"
textoAnt = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = texto & " - " & textoAnt
session.findById("wnd[0]/tbar[0]/btn[3]").press
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Windows("Planilha Reversa.xlsb").Activate
    Sheets("Cancelar Ordem").Select

ActiveCell.Offset(0, 3).Value = "Ordem Zerada."
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation
'
End Sub

Sub Eliminar_OI()

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0


    Sheets("Cancelar Ordem").Select
    Final = Range("D100000").End(xlUp).Row + 1
        
    Range("A" & Final).Select
    

nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

OI = ActiveCell.Offset(0, 0).Value
vazio = ActiveCell.Offset(0, 0).Value

If vazio = "" Then
GoTo fim
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[0]/menu[10]").Select
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press



Windows("Planilha Reversa.xlsb").Activate
    Sheets("Cancelar Ordem").Select
    
ActiveCell.Offset(0, 3).Value = "Eliminada."
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation

End Sub

Sub Reativar_OI()

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0

    Sheets("Cancelar Ordem").Select
    Final = Range("D100000").End(xlUp).Row + 1
        
    Range("A" & Final).Select

nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

ordem = ActiveCell.Offset(0, 0).Value
motivo = ActiveCell.Offset(0, 1).Value
texto = ActiveCell.Offset(0, 2).Value
vazio = ActiveCell.Offset(0, 0).Value
aviso = ""

If vazio = "" Then
GoTo fim
End If


session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = ordem
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
aviso = session.findById("wnd[0]/sbar").Text
If aviso = "O documento SD " & ordem & " não existe no banco de dados ou não foi arquivado" Then
Windows("Criação.xlsb").Activate
    Sheets("Cancelar").Select
ActiveCell.Offset(0, 3).Value = "OI não existe"
GoTo nexiste
End If
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
On Error GoTo 0

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = motivo
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").Select

volta555:

On Error Resume Next
barra = session.findById("wnd[0]/sbar").Text
On Error GoTo 0
teste = Right(barra, 7)
If Right(barra, 7) = "consumo" Then
session.findById("wnd[0]").sendVKey 0
barra = ""
GoTo volta555
End If


session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/usr/cmbRV45A-S_ABGRU").Key = ""
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[7]").press

On Error Resume Next
barra = session.findById("wnd[0]/sbar").Text
If barra = "Não foi efetuada qualquer modificação de dados" Then
On Error GoTo 0
GoTo texto
End If

On Error Resume Next
volta:
txtd = session.findById("wnd[2]/usr/txtMESSTXT1").Text
If txtd <> "" Then
session.findById("wnd[0]").sendVKey 0
txtd = ""
GoTo volta
On Error GoTo 0
End If

texto:

session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0005", "Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0005", "Column1"
textoAnt = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = texto & " - " & textoAnt
session.findById("wnd[0]/tbar[0]/btn[3]").press

On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Windows("Planilha Reversa.xlsb").Activate
    Sheets("Cancelar Ordem").Select

ActiveCell.Offset(0, 3).Value = "Reativada."
nexiste:
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation
'
End Sub
