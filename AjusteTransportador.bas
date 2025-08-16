Attribute VB_Name = "AjusteTransportador"
Sub Macro_Ajuste_Transp_Zrec()
Attribute Macro_Ajuste_Transp_Zrec.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Windows("Ajuste Transportador ZREC").Activate
Sheets("ENTRADA").Select

Data_Inicial = Range("C4").Text
Data_Final = Range("D4").Text


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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzrec"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV").Select
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/chkP_290FIM").Selected = True
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/ctxtS_ERDAT-LOW").Text = Data_Inicial
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/ctxtS_ERDAT-HIGH").Text = Data_Final
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/ctxtS_ERDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/ctxtS_ERDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/btn%_S_VSTEL_%_APP_%-VALU_PUSH").press

session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1350"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1352"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1550"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/btnBTSELECIONAR").press

session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Base ajuste transportador Zrec.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 30

On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12
On Error GoTo 0

    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\Base ajuste transportador Zrec.XLS", _
        Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 4), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
        Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array( _
        28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), _
        Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 4), Array(40, 1)), _
        TrailingMinusNumbers:=True
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft

    Windows("Base ajuste transportador Zrec.XLS").Activate
    Sheets("Base ajuste transportador Zrec").Select
    Columns("W:W").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("Z:Z").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRigh
    
    Columns("Z:Z").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    
    Columns("W:W").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    ActiveWorkbook.Worksheets("Base ajuste transportador Zrec").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Base ajuste transportador Zrec").Sort.SortFields.Add Key:=Range("A2:A" & nl _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Base ajuste transportador Zrec").Sort
        .SetRange Range("A1:AM" & nl)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Windows("Base ajuste transportador Zrec.XLS").Activate
Sheets("Base ajuste transportador Zrec").Select
QTYLINHAS = Range("A10000").End(xlUp).Row
Range("A2:A" & QTYLINHAS).Select
Selection.Copy

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzv62"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "010101"
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 8

session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZVZREC.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6

On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12
On Error GoTo 0
    QTYLINHAS = ""

    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\ZVZREC.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 4), Array(7, 4), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 4), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 4), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1)), _
        TrailingMinusNumbers:=True
    
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp


    Windows("ZVZREC.XLS").Activate
    Sheets("ZVZREC").Select

    'ORDEM INVERSA CANCELADA
   
    cont = 2
    Do While Cells(cont, 2) <> ""
    If Cells(cont, 8).Value = "159" Or Cells(cont, 8).Value = "160" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
    'ORDEM INVERSA FATURADA
    
    cont = 2
    Do While Cells(cont, 2) <> ""
    If Cells(cont, 32).Value <> "" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
        
    Windows("Base ajuste transportador Zrec.XLS").Activate
    nl = Application.WorksheetFunction.CountA(Range("D:D"))
    Range("E2:E" & nl).Value = "=IFERROR(VLOOKUP(RC[-4],[ZVZREC.xls]ZVZREC!C2,1),""Excluir"")"
    Range("E2:E" & nl).Copy
    Range("E2:E" & nl).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Windows("Base ajuste transportador Zrec.XLS").Activate
    
    cont = 2
    Do While Cells(cont, 4) <> ""
    If Cells(cont, 5).Value = "Excluir" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
    Windows("Base ajuste transportador Zrec.XLS").Activate
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("A2:A" & nl).Select
    Selection.Copy

    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    Windows("Base ajuste transportador Zrec.XLS").Activate
    Sheets("Base ajuste transportador Zrec").Select
    Range("B2:B" & nl).Select
    Selection.Copy
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("B2").Select
    ActiveSheet.Paste

    Windows("Base ajuste transportador Zrec.XLS").Activate
    Sheets("Base ajuste transportador Zrec").Select
    Range("C2:C" & nl).Select
    Selection.Copy
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("C2").Select
    ActiveSheet.Paste
    
    Windows("Base ajuste transportador Zrec.XLS").Activate
    Sheets("Base ajuste transportador Zrec").Select
    Range("D2:D" & nl).Select
    Selection.Copy
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("D2").Select
    ActiveSheet.Paste

    
    Windows("Base ajuste transportador Zrec.XLS").Activate
    Sheets("Base ajuste transportador Zrec").Select
    Range("I2:I" & nl).Select
    Selection.Copy
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("E2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("F2:F" & nl).Value = "=IFERROR(VLOOKUP(RC[-5],[ZVZREC.xls]ZVZREC!C2:C22,21,0),""DESCONSIDERAR"")"
    Range("F2:F" & nl).Copy
    Range("F2:F" & nl).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("G2:G" & nl).Value = "=RC[-1]=RC[-3]"
    Range("G2:G" & nl).Copy
    Range("G2:G" & nl).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Rows("1:1").Select
   
    Sheets("DADOS_ZREC").Select

    cont = 2
    Do While Cells(cont, 1) <> ""
    If Cells(cont, 7).Value = True Or Cells(cont, 6).Value = "DESCONSIDERAR" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
    If Range("A2").Value = "" Then
    MsgBox "Sem dados para processamento"
    GoTo fim
    End If
    
    'verificar formulas
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("G2:G" & nl).Value = "=IFERROR(VLOOKUP(RC[-6],[ZVZREC.xls]ZVZREC!C2,1,0),""DESCONSIDERAR"")"
    Range("G2:G" & nl).Copy
    Range("G2:G" & nl).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    cont = 2
    Do While Cells(cont, 1) <> ""
    If Cells(cont, 5).Value = "DESCONSIDERAR" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
    
    
    If Range("A2").Value = "" Then
    MsgBox "Sem dados para processamento"
    GoTo fim
    End If
 
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("G2:G" & nl).Select
    Selection.Delete
    Range("A1").Select
fim:
    Call retirar_linha
    
    MsgBox "EXTRAÇÃO CONCLUÍDA"
    


    Application.DisplayAlerts = False
    Windows("Base ajuste transportador Zrec.XLS").Activate
    ActiveWindow.Close
    Windows("ZVZREC.XLS").Activate
    ActiveWindow.Close
    
End Sub

Sub Can_Zrec()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

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

    Sheets("Cancelar ZREC").Select
    Windows("Planilha Reversa.xlsb").Activate
    
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzrec"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV").Select
nl = Application.WorksheetFunction.CountA(Range("C:C"))
cont = 1 + nl
Do While Cells(cont, 1).Value <> ""

session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/txtS_NFENUM-LOW").Text = Cells(cont, 1).Value
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/txtS_NFENUM-LOW").SetFocus
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/subSUBSCREEN:SAPLZGPL204:9411/txtS_NFENUM-LOW").caretPosition = 9
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/btnBTSELECIONAR").press
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").currentCellColumn = "NFENUM"
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").clickCurrentCell
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/cntlCONTAINER_9150/shellcont/shell/shellcont[0]/shell").pressToolbarButton "ZPLPROCESSO"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = Cells(cont, 2).Value
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_9000/tabpTAB_9000_REV/ssubSUBSCREEN:SAPLZGPL204:9410/btnBTSELECIONAR").press
Cells(cont, 3).Value = "ok"
cont = cont + 1

Loop


End Sub

Sub retirar_linha()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Windows("Ajuste Transportador ZREC").Activate
Sheets("DADOS_ZREC").Select


    cont = 2
    Do While Cells(cont, 1) <> ""
    If Cells(cont, 4).Value = "5003255" Or Cells(cont, 4).Value = "5003254" Or Cells(cont, 4).Value = "5003482" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
End Sub

Sub TMA_3255()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Windows("Ajuste Transportador ZREC").Activate
Sheets("ENTRADA").Select

    Sheets("DADOS_ZREC").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzv62n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").Text = "TMA5003255"
session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "TMA.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6

On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12
On Error GoTo 0

    ChDir "C:\temp"
    Workbooks.OpenText Filename:="C:\temp\TMA.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 4), Array(7, 4), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 4), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 4), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1)), _
        TrailingMinusNumbers:=True
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    ActiveCell.Cells.Select
    ActiveWorkbook.Worksheets("TMA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TMA").Sort.SortFields.Add Key:=ActiveCell.Offset(0 _
        , 16).Range("A1:A1000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TMA").Sort
        .SetRange ActiveCell.Offset(-1, 0).Range("A1:AX1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    If Range("Q2").Value = "" Then
    GoTo fim
    End If
    
    nl = Application.WorksheetFunction.CountA(Range("Q:Q"))
    cont = 2
    Do While Cells(cont, 17) <> ""
    teste = Cells(cont, 20).Value
    If Cells(cont, 20).Value = "" Then
    Rows(cont & ":" & cont).Select
    Selection.Delete Shift:=xlUp
    Else
    cont = cont + 1
    End If
    Loop
    
    If Range("Q2").Value = "" Then
    GoTo fim
    End If
    
    Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Range("F1").Select
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Transportador"
    Range("D2").Select
    Columns("D:D").EntireColumn.AutoFit
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("D2:D" & nl).Value = "=IFERROR(IF(RC[2]=7520,VLOOKUP(RC[1],'[Ajuste Transportador ZREC.xlsm]CDC'!C7:C9,3,0),VLOOKUP(RC[1],'[Ajuste Transportador ZREC.xlsm]CDC'!C7:C8,2,0)),""Buscar"")"
    Range("D2:D" & nl).Copy
    Range("D2:D" & nl).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    cont = 2
    Do While Cells(cont, 2) <> ""
        If Cells(cont, 4).Value = "Buscar" Then
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva03"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = Cells(cont, 1).Value
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 9
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select
            recb = 1
volta15:
            ver1 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & recb & "]").Text
            If Left(ver1, 2) = "WE" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & recb & "]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            On Error Resume Next
            Cells(cont, 5).Value = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").Text
            session.findById("wnd[1]/tbar[0]/btn[12]").press
            On Error GoTo 0
            Cells(cont, 4).Value = "=IFERROR(IF(RC[2]=7520,VLOOKUP(RC[1],'[Ajuste Transportador ZREC.xlsm]CDC'!C7:C9,3,0),VLOOKUP(RC[1],'[Ajuste Transportador ZREC.xlsm]CDC'!C7:C8,2,0)),""Buscar"")"
            Cells(cont, 4).Copy
            Cells(cont, 4).PasteSpecial xlPasteValues
            GoTo pula30
            Else
            recb = recb + 1
            GoTo volta15
            End If
pula30:
        Else
        cont = cont + 1
        End If
    Loop
    
    Range("A2:D" & nl).Copy
    Windows("Ajuste Transportador ZREC.xlsm").Activate
    Sheets("DADOS_ZREC").Select
    Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Columns("A:J").Select
    nl = Application.WorksheetFunction.CountA(Range("A:A"))
    ActiveSheet.Range("$A$1:$J$" & nl).RemoveDuplicates Columns:=1, Header:=xlYes
    
fim:
    Windows("TMA.XLS").Activate
    ActiveWindow.Close
    
End Sub

