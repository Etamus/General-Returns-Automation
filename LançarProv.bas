Attribute VB_Name = "LançarProv"
Sub Lançar_prov()

    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Lançar Providência").Select
    
nl = Application.WorksheetFunction.CountA(Range("B:B"))
Range("B2:B" & nl).Copy

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_TRANSP-LOW").Text = "1"
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").Text = "010101"
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_TRANSP_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]").sendVKey 8
session.findById("wnd[0]/usr/ctxtS_CODOC-LOW").Text = Range("C4").Value
session.findById("wnd[0]").sendVKey 8

On Error Resume Next
erro01 = session.findById("wnd[1]/usr/txtMESSTXT1").Text
If erro01 = "Não há dados para essa seleção." Then
session.findById("wnd[1]/tbar[0]/btn[0]").press
GoTo pular
End If
On Error GoTo 0

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "HOROC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "VPAGDIF"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "DTPRCVLROC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "HOROC"
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "TXTOC2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "CODPROV"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "C"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 1
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").SetFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/tbar[0]/btn[0]").press

volta:

On Error Resume Next
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[9]").press
erro01 = session.findById("wnd[1]/usr/txtMESSTXT1").Text
If erro01 = "Selecionar um registro!" Then
session.findById("wnd[1]/tbar[0]/btn[0]").press
GoTo pular
End If
On Error GoTo 0

session.findById("wnd[1]/usr/ctxtW_SAIDA-CODPROV").SetFocus
session.findById("wnd[1]/usr/ctxtW_SAIDA-CODPROV").Text = Range("D4").Value
session.findById("wnd[1]/usr/cntlCC_PROVIDENCIA/shell").SetFocus
session.findById("wnd[1]/usr/cntlCC_PROVIDENCIA/shell").Text = Range("E4").Value
session.findById("wnd[1]/usr/btnSAVE").press
session.findById("wnd[2]/tbar[0]/btn[0]").press

GoTo volta


pular:


MsgBox "Lançamento Finalizado", vbInformation

End Sub
Sub Lançar_prov_Por_Oc()

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_TRANSP-LOW").Text = "1"
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").Text = "010101"
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_DTEXP-LOW").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell

    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Lançar Providência").Select

nl = Application.WorksheetFunction.CountA(Range("F:F")) + 1
cont = 1 + nl

Do While Cells(cont, 1) <> ""

session.findById("wnd[0]/usr/ctxtS_TRANSP-LOW").Text = Cells(cont, 2).Value
session.findById("wnd[0]/usr/ctxtS_CODOC-LOW").Text = Cells(cont, 3).Value
session.findById("wnd[0]/usr/ctxtS_CODOC-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_CODOC-LOW").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press

On Error Resume Next
erro01 = session.findById("wnd[1]/usr/txtMESSTXT1").Text
If erro01 = "Não há dados para essa seleção." Then
Cells(cont, 6).Value = "Não Lançado"
session.findById("wnd[1]/tbar[0]/btn[0]").press
erro01 = ""
GoTo pular
End If
On Error GoTo 0

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "HOROC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "VPAGDIF"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "DTPRCVLROC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "HOROC"
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "STATUS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "TXTOC2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "CODPROV"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "STATUS"
session.findById("wnd[0]/tbar[1]/btn[29]").press

If Cells(cont, 4).Value * 1 <> 22 Then
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "C"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 1
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").SetFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/tbar[0]/btn[0]").press
Else
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "C"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 1
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").SetFocus
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

On Error Resume Next
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[9]").press
erro01 = session.findById("wnd[1]/usr/txtMESSTXT1").Text
If erro01 = "Selecionar um registro!" Then
Cells(cont, 6).Value = "Não Lançado"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
erro01 = ""
GoTo pular
End If
On Error GoTo 0

session.findById("wnd[1]/usr/ctxtW_SAIDA-CODPROV").SetFocus
session.findById("wnd[1]/usr/ctxtW_SAIDA-CODPROV").Text = Cells(cont, 4).Value
session.findById("wnd[1]/usr/cntlCC_PROVIDENCIA/shell").SetFocus
session.findById("wnd[1]/usr/cntlCC_PROVIDENCIA/shell").Text = Cells(cont, 5).Value
session.findById("wnd[1]/usr/btnSAVE").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Cells(cont, 6).Value = "ok"
pular:
cont = cont + 1

Loop

MsgBox "Lançamento Finalizado", vbInformation

End Sub

