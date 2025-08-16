Attribute VB_Name = "BuscarChaveAcessoMlog"
Sub Buscar_Chave_Acesso_Mlog()
Attribute Buscar_Chave_Acesso_Mlog.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
Range("A2").Select

QTYLINHAS = Range("A100000").End(xlUp).Row
ActiveSheet.Range("$A$1:$C$" & QTYLINHAS).RemoveDuplicates Columns:=1, Header:= _
        xlYes
QTYLINHAS = ""
    
linha = Range("A100000").End(xlUp).Row
SUBLINHA = Range("C100000").End(xlUp).Row
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

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
nl = Range("A10000").End(xlUp).Row
For i = 0 To nl
OI = ActiveCell.Offset(0, 0).Value
vazio = ActiveCell.Offset(0, 0).Value

If vazio = "" Then
GoTo fim
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/ctxtVBRK-VBELN").SetFocus
FATURAMENTO = session.findById("wnd[0]/usr/ctxtVBRK-VBELN").Text
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
ActiveCell.Offset(0, 2).Value = FATURAMENTO
ActiveCell.Offset(1, 0).Select

Next

PROXIMO = ActiveCell.Offset(0, 0).Value
If PROXIMO = "" Then
GoTo fim
End If

session.findById("wnd[0]").sendVKey 3

fim:

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
Range("A2").Select

linha = Range("A10000").End(xlUp).Row
LinhaSub = Range("B10000").End(xlUp).Row
linha = LinhaSub + 1
Range("A" & linha).Select

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
nl = Range("A10000").End(xlUp).Row

For i = 0 To nl
OI = ActiveCell.Offset(0, 0).Value
vazio = ActiveCell.Offset(0, 0).Value
doc_brasa = ActiveCell.Offset(0, 2).Value

If vazio = "" Then
GoTo CONCLUIDO
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZVAG13"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radR3").SetFocus
session.findById("wnd[0]/usr/radR3").Select
session.findById("wnd[0]/usr/ctxtP_FATUR").Text = doc_brasa
session.findById("wnd[0]/usr/ctxtP_FATUR").SetFocus
session.findById("wnd[0]/usr/ctxtP_FATUR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = "FATURA2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "ESTOQUE1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").doubleClickItem "          5", "&Hierarchy"
session.findById("wnd[0]/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/shellcont/shell").pressToolbarButton "&DETAIL"
session.findById("wnd[1]/tbar[0]/btn[71]").press
session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "em processamento"
session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").caretPosition = 16
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[2]/usr/txtGS_SEARCH-SEARCH_INFO").SetFocus
PROCESS = session.findById("wnd[2]/usr/txtGS_SEARCH-SEARCH_INFO").Text
session.findById("wnd[2]/tbar[0]/btn[12]").press
session.findById("wnd[0]/shellcont/shell").pressToolbarButton "&DETAIL"
session.findById("wnd[1]").Close

If PROCESS <> "A ocorrência será exibida : 1" Then
Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso e Mlog").Select
ActiveCell.Offset(0, 1).Value = "Não há faturamento para a ordem MLOG"
ActiveCell.Offset(1, 0).Select
PROCESS = ""
GoTo seguir
End If

session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          6", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          6", "&Hierarchy"
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
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

Windows("Planilha Reversa.xlsb").Activate
Sheets("Buscar Chave de Acesso & Mlog").Select
ActiveCell.Offset(0, 1).Value = "'" & CHAVE
ActiveCell.Offset(1, 0).Select
seguir:

Next

PROXIMO = ActiveCell.Offset(0, 0).Value
If PROXIMO = "" Then
GoTo CONCLUIDO
End If

CONCLUIDO:
MsgBox "Finalizado.", vbInformation

End Sub
