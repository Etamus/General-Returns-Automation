Attribute VB_Name = "LançarOcorrência"
Sub Lançar_Ocorrencia_NF()
Attribute Lançar_Ocorrencia_NF.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Lançar Ocorrência").Select
Final = Range("H10000").End(xlUp).Row + 1
Range("A" & Final).Select
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

NF = Format(ActiveCell.Offset(0, 1).Value, "000000000")
digito = ActiveCell.Offset(0, 2).Value
Deposito = ActiveCell.Offset(0, 3).Value
TR = ActiveCell.Offset(0, 4).Value
Oc = ActiveCell.Offset(0, 5).Value
texto = ActiveCell.Offset(0, 6).Value
vazio = ActiveCell.Offset(0, 1).Value

If vazio = "" Then
GoTo fim
End If

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr07"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtST_SELECAO-NFNUM").Text = NF
session.findById("wnd[0]/usr/txtST_SELECAO-SERIES").Text = digito
session.findById("wnd[0]/usr/txtST_SELECAO-VSTEL").Text = Deposito
session.findById("wnd[0]/usr/ctxtST_SELECAO-LIFNR").Text = TR
session.findById("wnd[0]/usr/txtST_SELECAO-CODOC").Text = Oc
igual = session.findById("wnd[0]/usr/ctxtVBAK-VDATU").Text
session.findById("wnd[0]/usr/ctxtVBAK-AUDAT").Text = igual
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER01/shell").Text = texto
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER01/shell").setSelectionIndexes 50, 50
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER04/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER04/shellcont/shell").setCurrentCell -1, "MESSAGE"
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER04/shellcont/shell").selectColumn "MESSAGE"
session.findById("wnd[0]/usr/cntlCUSTOM_CONTAINER04/shellcont/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[1]").sendVKey 4
info = session.findById("wnd[2]/usr/lbl[1,3]").Text
session.findById("wnd[2]/usr/lbl[1,3]").caretPosition = 5
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

Windows("Planilha Reversa").Activate
Sheets("Lançar Ocorrência").Select
ActiveCell.Offset(0, 7).Value = info
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation

End Sub
