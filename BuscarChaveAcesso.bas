Attribute VB_Name = "BuscarChaveAcesso"
Sub Buscar_Chave_Acesso()
Attribute Buscar_Chave_Acesso.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
Range("A2").Select

QTYLINHAS = Range("A100000").End(xlUp).Row
ActiveSheet.Range("$A$1:$B$" & QTYLINHAS).RemoveDuplicates Columns:=1, Header:= _
        xlYes
QTYLINHAS = ""
    
linha = Range("A100000").End(xlUp).Row
SUBLINHA = Range("B100000").End(xlUp).Row
linha = SUBLINHA + 1
Range("A" & linha).Select

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVA03"
session.findById("wnd[0]").sendVKey 0

Dim OI, TIPO, CHAVE, PEDIDO As Variant
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1
For i = 0 To nl
OI = ActiveCell.Offset(0, 0).Value
If OI = "" Then
GoTo fim
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-AUART").SetFocus
TIPO = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-AUART").Text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK").SetFocus
PEDIDO = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

If TIPO = "REB" Or TIPO = "ZDRG" Then

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZV62"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_VBELN-LOW").Text = OI
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "010101"
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").caretPosition = 6
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "ACCKEY"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "DOCEST"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "ACCKEY"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]").sendVKey 4
CHAVEZV = session.findById("wnd[2]/usr/lbl[1,3]").Text
CHAVE = Right(CHAVEZV, 44)
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVA03"
session.findById("wnd[0]").sendVKey 0

If CHAVE = "" Then
Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave").Select
ActiveCell.Offset(0, 1).Value = "Buscar NF no Avalara"
ActiveCell.Offset(1, 0).Select
Else
GoTo seguir
End If
GoTo próximo
End If

If TIPO = "ZRRG" And PEDIDO = "ZLR6" Then
GoTo RECUSA
End If

If TIPO = "ZRRG" And PEDIDO = "ZLR8" Then
GoTo SIMULTÂNEA
End If

'COLETA NORMAL CHAVE DA NOTA DE ENTRADA
If PEDIDO = "ZLR1" Then
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").SetFocus
CHAVE = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
GoTo seguir
End If

'COLETA SIMULTÂNEA CHAVE DA NOTA DE ENTRADA
If PEDIDO = "ZLR2" Then
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").SetFocus
CHAVE = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
GoTo seguir
End If

'RECUSA CHAVE DA NOTA DE SAÍDA
If PEDIDO = "ZLR3" Then
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          3", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          3", "&Hierarchy"
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").SetFocus
CHAVE = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
GoTo seguir
End If

'RECUSA CHAVE DA NOTA DE SAÍDA
RECUSA:
If PEDIDO = "ZLR6" Then
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          3", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          3", "&Hierarchy"
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").SetFocus
CHAVE = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
GoTo seguir
End If

'COLETA SIMULTÂNEA CHAVE DA NOTA DE ENTRADA
SIMULTÂNEA:
If PEDIDO = "ZLR8" Then
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          2", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          2", "&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").SetFocus
CHAVE = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/txtJ_1B_NFE_SCREEN_FIELDS-ACCKEY").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
GoTo seguir
End If

seguir:
Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
ActiveCell.Offset(0, 1).Value = "'" & CHAVE
ActiveCell.Offset(1, 0).Select

próximo:
CHAVEZV = ""
CHAVE = ""
OI = ""
TIPO = ""
PEDIDO = ""
Next

fim:

End Sub
