Attribute VB_Name = "CriarRemessa"
Sub CRIAR_REMESSA()
Attribute CRIAR_REMESSA.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Alterar Remessa, OI ou TR").Select
    Range("A2").Select
    linha = Range("A100000").End(xlUp).Row
    LinhaSub = Range("B100000").End(xlUp).Row
    linha = LinhaSub + 1
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl01n"
session.findById("wnd[0]").sendVKey 0

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
Final = Range("A100000").End(xlUp).Row

Dim vazio, Deposito, ordem
nl = Application.WorksheetFunction.CountA(Range("A:A"))

For i = 0 To nl

vazio = ActiveCell.Offset(0, 0).Value
ordem = ActiveCell.Offset(0, 0).Value
Deposito = ActiveCell.Offset(0, 8).Value

If vazio = "" Then
GoTo fim
End If

Inicio:

session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").Text = Deposito
session.findById("wnd[0]/usr/ctxtLV50C-VBELN").Text = ordem
session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").SetFocus
session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
DOCNUM = session.findById("wnd[0]/sbar").Text

If DOCNUM = "Não se pode selecionar código de função" Then
GoTo Inicio
End If

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
DOCNUM_BASE = DOCNUM
DOCNUM_BASE = Mid(DOCNUM, 18, 10)
ActiveCell.Offset(0, 1).Value = DOCNUM_BASE
DOCNUM = ""
DOCNUM_BASE = ""
ActiveCell.Offset(1, 0).Select

Next

fim:

session.findById("wnd[0]").sendVKey 12
    
MsgBox "Remessa Criada."

End Sub
