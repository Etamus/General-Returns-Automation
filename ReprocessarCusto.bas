Attribute VB_Name = "ReprocessarCusto"
Sub Excluir_Custo()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("B2").Select

Dim TR
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl
TR = ActiveCell.Offset(0, 0).Value

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
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID").Select
'placa = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").Text
'If placa = "" Then
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").SetFocus
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").Text = "REVERSE"
'placa = ""
'End If
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC").Select
processo = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/cmbVTTK-FBGST").Text
    
    If processo = "A não processado                                                                                                                                                                                                                                              " Then
    GoTo nprocesso
    End If

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/btnSCD_DISPLAY_1").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

On Error Resume Next
cte = session.findById("wnd[1]").Text
If cte = "Cancelamento" Then
ActiveCell.Offset(0, 6).Value = "CTe associado"
GoTo pula
End If
On Error GoTo 0
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
retor:
session.findById("wnd[0]").sendVKey 0
nprocesso:

On Error Resume Next
If Left(session.findById("wnd[0]/sbar").Text, 12) = "O transporte" Then
session.findById("wnd[0]").sendVKey 0
GoTo retor
End If
On Error GoTo 0

'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID").Select
'placa = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").Text
'If placa = "" Then
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").SetFocus
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_ID/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1022/txtVTTK-EXTI2").Text = "REVERSE"
'placa = ""
'session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR").Select
'End If

If session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DALBG").Text <> "" Then
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STLBG").press
End If

If session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DAREG").Text <> "" Then
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STREG").press
End If

If session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DTDIS").Text <> "" Then
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STDIS").press
End If

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").Text = ""

'fnovo

session.findById("wnd[0]/tbar[0]/btn[11]").press

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
pula:

'verificar necessidade
repete:
ActiveCell.Offset(1, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
    GoTo fim
    Else
        If ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(-1, 0).Value Then
        
        GoTo repete
        End If
    End If
'f'verificar necessidade
segue:
cte = ""
dt = ""
Next

fim:

    Sheets("Alteração Geral").Select

End Sub
Sub Alterar_Remessa_OI()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("C1").Select

Dim Remessa, Cod
nl = Application.WorksheetFunction.CountA(Range("C:C")) - 1

For i = 0 To nl

Remessa = ActiveCell.Offset(1, 0).Value
erro = ActiveCell.Offset(1, 5).Value
Cod = ActiveCell.Offset(1, 1).Value

If Remessa = "" Then
GoTo fim
End If

If erro = "CTe associado" Then
GoTo pula
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

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select
expe = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV50A:2110/ctxtLIKP-VSTEL").Text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
CD = ""

contador3 = 2

Do While Sheets("CDC").Cells(contador3, 1) <> ""

            
            If Sheets("CDC").Cells(contador3, 1).Value = Cod Then
            
                If expe = "7420" Or expe = "7520" Then
                CD = ""
                GoTo CDC
                End If
                If expe = "1400" Or expe = "1441" Or expe = "1443" Or expe = "1444" Then
                CD = Sheets("CDC").Cells(contador3, 4).Value
                GoTo CDC
                Else
                CD = Sheets("CDC").Cells(contador3, 3).Value
                GoTo CDC
                End If
              Else
              contador3 = contador3 + 1
            End If
            
Loop
CD = ""
CDC:
'elimina C2
Dim linha
linha = 0
confere:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
GoTo C2eliminado
End If
If ver = "C2 Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 0
GoTo C2eliminado
Else
linha = linha + 1
GoTo confere
End If
C2eliminado:

If CD = "" Then
GoTo SemCD
End If

'incluir CD
linha = 0
confere2:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "CD"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = CD
session.findById("wnd[0]").sendVKey 0
GoTo CDincluido
End If
If ver = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = CD
session.findById("wnd[0]").sendVKey 0
GoTo CDeliminado
Else
linha = linha + 1
GoTo confere2
End If

SemCD:

'elimina CD
linha = 0
confere1:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
GoTo CDeliminado
End If
If ver = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 0
GoTo CDeliminado
Else
linha = linha + 1
GoTo confere1
End If
CDeliminado:
CDincluido:

'verifica transportador
linha = 0
confere3:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "SP"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Cod
session.findById("wnd[0]").sendVKey 0
GoTo STincluido
End If
If ver = "SP Transportador" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV50A:2114/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod
session.findById("wnd[0]").sendVKey 0
GoTo STincluido
Else
linha = linha + 1
GoTo confere3
End If
STincluido:

session.findById("wnd[0]/tbar[0]/btn[11]").press

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
    
pula:
ActiveCell.Offset(1, 0).Select

Next

fim:

Call Ajustar_OI

End Sub
Sub Ajustar_OI()


Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("A1").Select

Dim OI, Cod
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

OI = ActiveCell.Offset(1, 0).Value
Cod = ActiveCell.Offset(1, 3).Value
erro = ActiveCell.Offset(1, 7).Value

If OI = "" Then
GoTo fim
End If
If erro = "CTe associado" Then
GoTo pula
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
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error GoTo 0
expe = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VSTEL[9,0]").Text
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select


CD = ""

contador3 = 2

Do While Sheets("CDC").Cells(contador3, 1) <> ""

            If Sheets("CDC").Cells(contador3, 1).Value = Cod Then
                If expe = "7420" Or expe = "7520" Then
                CD = ""
                GoTo CDC
                End If
                If expe = "1400" Or expe = "1441" Or expe = "1443" Or expe = "1444" Then
                CD = Sheets("CDC").Cells(contador3, 4).Value
                GoTo CDC
                Else
                CD = Sheets("CDC").Cells(contador3, 3).Value
                GoTo CDC
                End If
              Else
              contador3 = contador3 + 1
            End If
            Loop
CD = ""
CDC:
'elimina C2
Dim linha
linha = 0
confere:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
GoTo C2eliminado
End If
If ver = "C2 Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 0
GoTo C2eliminado
Else
linha = linha + 1
GoTo confere
End If
C2eliminado:

If CD = "" Then
GoTo SemCD
End If

'incluir CD
linha = 0
confere2:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "CD"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = CD
session.findById("wnd[0]").sendVKey 0
GoTo CDincluido
End If
If ver = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = CD
session.findById("wnd[0]").sendVKey 0
GoTo CDeliminado
Else
linha = linha + 1
GoTo confere2
End If

SemCD:

'elimina CD
linha = 0
confere1:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
GoTo CDeliminado
End If
If ver = "CD Cross Docking" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").SetFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnDELETE").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 0
GoTo CDeliminado
Else
linha = linha + 1
GoTo confere1
End If
CDeliminado:
CDincluido:

'verifica transportador
linha = 0
confere3:
ver = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & linha & "]").Text
If ver = "" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").Key = "SP"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").Text = Cod
session.findById("wnd[0]").sendVKey 0
GoTo STincluido
End If
If ver = "SP Transportador" Then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & linha & "]").Text = Cod
session.findById("wnd[0]").sendVKey 0
GoTo STincluido
Else
linha = linha + 1
GoTo confere3
End If
STincluido:

session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
gravar = session.findById("wnd[1]/usr/txtSPOP-TEXTLINE3").Text
If gravar <> "" Then
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
gravar = ""
End If
On Error GoTo 0

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
pula:
ActiveCell.Offset(1, 0).Select

Next

fim:

Call Organizar_TR

End Sub
Sub Organizar_TR()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("B1").Select

Dim Transporte
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl

Transporte = ActiveCell.Offset(1, 0).Value
erro = ActiveCell.Offset(1, 6).Value
codTR = ActiveCell.Offset(1, 2).Value

If Transporte = "" Then
GoTo fim
End If

If erro = "CTe associado" Then
GoTo pula
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
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = Transporte
session.findById("wnd[0]").sendVKey 0


DATA_CUSTO = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DTDIS").Text
If DATA_CUSTO <> "" Then
'MsgBox " Verificar se o " & Transporte & " tem CTE associado."
ActiveCell.Offset(1, 6).Value = "Erro ao excluir/Gerar custo"
DATA_CUSTO = ""
GoTo seguir
End If

If codTR = "" Then
GoTo semTR
End If

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1021/ctxtVTTK-TDLNR").Text = codTR
session.findById("wnd[0]").sendVKey 0
semTR:
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").SetFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").Key = "PB"
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE").Select

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STDIS").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[0,0]").Selected = True
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").Text = ""
seguir:

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").SetFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STLBG").press
session.findById("wnd[0]/tbar[0]/btn[11]").press

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
pula:
ActiveCell.Offset(1, 0).Select
While ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(1, 0).Value
If ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(1, 0).Value Then
ActiveCell.Offset(1, 0).Select
End If
Wend

Next

fim:

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

End Sub
Sub Custo()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("B2").Select

Range("B2").Select

Dim TR, Custo
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl

TR = ActiveCell.Offset(0, 0).Value
erro = ActiveCell.Offset(0, 6).Value

If TR = "" Then
GoTo fim
End If

If erro <> "" Then
GoTo pula
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

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

ActiveCell.Offset(0, 3).Value = Custo * 1

'ZSTR01 E ZSTR64

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

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

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
ActiveCell.Offset(0, 4).Value = "ZSTR OK"

'ZSTR44 NOTFIS

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = "0"
session.findById("wnd[0]/usr/radP_OPT2").SetFocus
session.findById("wnd[0]/usr/radP_OPT2").Select

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

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

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

ActiveCell.Offset(0, 5).Value = Notfis
pula:
ActiveCell.Offset(1, 0).Select

Next

fim:

'MsgBox "PROCESSO ENCERRADO", vbInformation
Sheets("Alteração Geral").Select


End Sub
Sub Full()

Call Excluir_Custo
Call Alterar_Remessa_OI
Call Custo

End Sub
Sub Reprocessar_custo()

Application.ScreenUpdating = False
Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
Range("B2").Select

Dim TR
nl = Application.WorksheetFunction.CountA(Range("B:B")) - 1

For i = 0 To nl
TR = ActiveCell.Offset(0, 0).Value

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt03n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC").Select
processo = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/cmbVTTK-FBGST").Text
    
        'If processo = "A não processado                                                                                                                                                                                                                                              " Then
        'GoTo nprocesso
        'End If

session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/btnSCD_DISPLAY_1").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-POSTX[2,0]").caretPosition = 18
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[5]").press

ActiveCell.Offset(0, 6).Value = session.findById("wnd[0]/usr/tabsTABSTRIP_ITEM/tabpPPRI/ssubSCD_ITEM:SAPMV54A:0041/txtVFKP-NETWR").Text
session.findById("wnd[0]/tbar[0]/btn[11]").press

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

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select
ActiveCell.Offset(0, 4).Value = "ZSTR OK"

'ZSTR44 NOTFIS

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = "0"
session.findById("wnd[0]/usr/radP_OPT2").SetFocus
session.findById("wnd[0]/usr/radP_OPT2").Select

Windows("Planilha Reversa").Activate
Sheets("Ajuste Transportador").Select

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

Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

ActiveCell.Offset(0, 5).Value = Notfis
pula:
ActiveCell.Offset(1, 0).Select


Windows("Planilha Reversa").Activate
Sheets("Alteração Geral").Select

Next

fim:

    Sheets("Alteração Geral").Select


End Sub
