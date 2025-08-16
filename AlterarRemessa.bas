Attribute VB_Name = "AlterarRemessa"
Private Sub ALTERAR_REMESSA()

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
Range("B1").Select

Dim Remessa, Cod, Cod_CDC
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl

Remessa = ActiveCell.Offset(1, 0).Value
Cod = ActiveCell.Offset(1, 2).Value
Cod_CDC = ActiveCell.Offset(1, 3).Value

If Remessa = "" Then
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Remessa
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select

Dim CDC As Variant
If Cod_CDC <> "" Then
linha = 0
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until CDC = ""
If CDC = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod_CDC
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Exit Do
Else
linha = linha + 1
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
If CDC = "" And Cod_CDC <> "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "CD"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Cod_CDC
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
End If
Loop
linha = ""
GoTo ALTERAR_TRANP
End If

If Cod_CDC = "" Then
linha = 0
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until CDC = ""
If CDC = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Exit Do
Else
linha = linha + 1
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
Loop
linha = ""
End If

ALTERAR_TRANP:
Dim Transp As Variant
linha = 0
Transp = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until Transp = ""
If Transp = "SP Transportador" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
Exit Do
Else
linha = linha + 1
Transp = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
Loop
linha = ""

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
ActiveCell.Offset(1, 0).Select

Next

fim:

End Sub
Private Sub ALTERAR_OI()

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
Range("A1").Select

Dim OI, Cod
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

OI = ActiveCell.Offset(1, 0).Value
Cod = ActiveCell.Offset(1, 3).Value
Cod_CDC = ActiveCell.Offset(1, 4).Value

If OI = "" Then
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select

Dim CDC As Variant
If Cod_CDC <> "" Then
linha = 0
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until CDC = ""
If CDC = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod_CDC
Exit Do
Else
linha = linha + 1
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
If CDC = "" And Cod_CDC <> "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "CD"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Cod_CDC
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
End If
Loop
linha = ""
GoTo ALTERAR_TRANP
End If

If Cod_CDC = "" Then
linha = 0
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until CDC = ""
If CDC = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Exit Do
Else
linha = linha + 1
CDC = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
Loop
linha = ""
End If

ALTERAR_TRANP:
Dim Transp As Variant
linha = 0
Transp = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
Do Until Transp = ""
If Transp = "SP Transportador" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
Exit Do
Else
linha = linha + 1
Transp = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
End If
Loop
linha = ""

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
    
ActiveCell.Offset(1, 0).Select

Next

fim:

End Sub
Private Sub ALTERAR_TR()

Dim Custo As String

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
Range("C1").Select

Dim TR
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl
TR = ActiveCell.Offset(1, 0).Value
Cod = ActiveCell.Offset(1, 1).Value

If TR = "" Then
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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/btnSCD_DISPLAY_1").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STLBG").press
dt = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DAREG").Text
If dt <> "" Then
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STREG").press
End If
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1021/ctxtVTTK-TDLNR").Text = Cod
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").SetFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").Key = "PB"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[0,0]").Selected = True
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STREG").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").Text = ""
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").SetFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/btnPUSHB_PICK").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STLBG").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
dt = ""

'CUSTO

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvi01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").SetFocus
Custo = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").caretPosition = 21
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
ActiveCell.Offset(1, 3).Value = Custo

'ZSTR01 E ZSTR64

Windows("Criação Transporte.xlsm").Activate
Sheets("Alterar Remessa, OI ou TR").Select

retorno:
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtP_TKNUM").caretPosition = 9
session.findById("wnd[0]").sendVKey 8
On Error Resume Next
logzstr = session.findById("wnd[0]/usr/lbl[0,0]").Text
If logzstr = "Carga de Documentos de Transporte e Notas Fiscais do Transporte" Then
logzstr = ""
GoTo retorno
End If
On Error GoTo 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 8

session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr64"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtP_TKNUM").caretPosition = 9
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
On Error Resume Next
x = session.findById("wnd[1]/usr/txtMESSTXT2").Text
If x = "existe" Then
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
GoTo seguir
End If
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[1]").sendVKey 0
seguir:
session.findById("wnd[0]").sendVKey 3
On Error GoTo 0

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
ActiveCell.Offset(1, 4).Value = "ZSTR OK"

'ZSTR44 NOTFIS

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = "0"
session.findById("wnd[0]/usr/radP_OPT2").SetFocus
session.findById("wnd[0]/usr/radP_OPT2").Select

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select

Dim Notfis, vazio, msg1, msg2
session.findById("wnd[0]/usr/ctxtP_TKNUM").SetFocus
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtP_TKNUM").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
msg2 = session.findById("wnd[0]/sbar").Text
If msg2 = "A Declaração de Devolução foi enviada para o e-mail do Transportador" Then
GoTo notfi
End If
msg1 = session.findById("wnd[1]/usr/txtMESSTXT1").Text
If msg1 = "Não foi encontrado e-mail para envio da Declaração" Then
session.findById("wnd[1]/tbar[0]/btn[0]").press
GoTo notfi
End If

notfi:
Notfis = session.findById("wnd[0]/usr/txtP_NOTFIS").Text

ActiveCell.Offset(1, 5).Value = Notfis

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alterar Remessa, OI ou TR").Select
ActiveCell.Offset(1, 0).Select
While ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(1, 0).Value
If ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(1, 0).Value Then
ActiveCell.Offset(1, 0).Select
End If
Wend

Next

fim:

End Sub
Sub ALTERAR_COD_OI_REMESSA_TR()
Attribute ALTERAR_COD_OI_REMESSA_TR.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
Inicio:
 TIPO = InputBox("Inserir informação que deseja alterar o código transportador (OI), (Remessa) ou (TR)")
If TIPO = Empty Then
    MsgBox ("Inserir informação que deseja alterar (OI), (Remessa) ou (TR)"), vbCritical
    GoTo Inicio
End If

If TIPO = "Remessa" Then
Call ALTERAR_REMESSA
End If

If TIPO = "OI" Then
Call ALTERAR_OI
End If

If TIPO = "TR" Then
Call ALTERAR_TR
End If

MsgBox "Finalizado.", vbInformation

End Sub
