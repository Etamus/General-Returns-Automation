Attribute VB_Name = "AlterarRFQ"
Sub Alterar_RCF_Coletiva()
Attribute Alterar_RCF_Coletiva.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Alterar RFQ").Select
    Range("A2").Select
    
    QTYLINHAS = Range("A10000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$C$10000").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
    QTYLINHAS = ""
    
    ActiveWorkbook.Worksheets("Alterar RFQ").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Alterar RFQ").Sort.SortFields.Add Key:= _
        Range("A2:A10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Alterar RFQ").Sort
        .SetRange Range("A1:C10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    linha = Range("A100000").End(xlUp).Row
    LinhaSub = Range("C100000").End(xlUp).Row
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0

    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Alterar RFQ").Select
    
    QTYLINHAS3 = Range("C10000").End(xlUp).Row
    QTYLINHAS2 = Range("C10000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS2 + 1
    Range("A" & QTYLINHAS3).Select

Dim OI
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

OI = ActiveCell.Offset(0, 0).Value
Cod = ActiveCell.Offset(0, 1).Value

If OI = "" Then
GoTo fim
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select

check = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text
If check = Cod Then
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
ActiveCell.Offset(0, 2).Value = "RFQ já Atualizada"
GoTo estaok
End If
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").Text = Cod

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar RFQ").Select
    
ActiveCell.Offset(0, 2).Value = "RFQ atualizada"

If ActiveCell.Offset(0, 1).Value = "01" Then
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem "          5", "&Hierarchy"
session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          5", "&Hierarchy"
session.findById("wnd[0]/tbar[1]/btn[8]").press
sobe:
session.findById("wnd[0]/mbar/menu[0]/menu[4]").Select

On Error Resume Next
teste = Left(session.findById("wnd[1]/usr/txtMESSTXT1").Text, 5)
If teste = "Ordem" Then
session.findById("wnd[0]").sendVKey 0
teste = ""
GoTo sobe
End If

On Error GoTo 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select

Dim CDC As Variant
linha = 0
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until CDC = ""
If CDC = "SP Transportador" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").Text = "5002359"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").caretPosition = 7
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
Exit Do
Else
linha = linha + 1
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
Loop
linha = ""
End If

estaok:
ActiveCell.Offset(1, 0).Select

Next

Range("A2").Select

fim:
MsgBox "Finalizado.", vbInformation

End Sub
