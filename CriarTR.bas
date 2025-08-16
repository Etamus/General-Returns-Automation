Attribute VB_Name = "CriarTR"
Sub Criar_TR_Peças()
Attribute Criar_TR_Peças.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Custo As String

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar Transporte").Select
Range("A2").Select
    
nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

Dim TR, dt, codTR, condexp

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar Transporte").Select

Remessa = ActiveCell.Offset(0, 0).Value
Deposito = ActiveCell.Offset(0, 1).Value
codTR = ActiveCell.Offset(0, 2).Value
xp = ActiveCell.Offset(0, 3).Value
dtinicial = ActiveCell.Offset(0, 4).Value
dtfinal = ActiveCell.Offset(0, 5).Value
condexp = ActiveCell.Offset(0, 6).Value

OI = ActiveCell.Offset(0, 0).Value

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr06"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radP_OUTR").Select
session.findById("wnd[0]/usr/radP_SMLOG").Select
session.findById("wnd[0]/usr/radP_REVER").Select

session.findById("wnd[0]/usr/ctxtS_VSTEL-LOW").Text = Deposito
session.findById("wnd[0]/usr/ctxtS_REMES-LOW").Text = Remessa
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = "010101"
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").caretPosition = 6
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5, "TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TVFTZT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VSTEL"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ROUTE"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "COCKPIT_NUM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ENVIADO_CPL"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ERDAT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LFDAT_DES"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WADAT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LFDAT_ENT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KUNNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "NAME1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ORT01"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "REGIO"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TKNUM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SA_DESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DDTEXT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BEZEI_ESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BEZEI_EQU"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "QTDIAS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CD_PARC_ATUA"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SORTL_CDC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AUTLF_AUX"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KUKLA"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VTEXT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DDTEXT_STS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DDTEXT_RES"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "INCO1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SORTL_OPL"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MATNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "QTDE"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BRGEW"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "GEWEI"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLUM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLEH"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EMBARQUE"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VGBEL"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BSARK"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LIFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VSTEL_ORI"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "OBS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "OBS1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "OBS2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "OBS3"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "OBS4"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DIV_REM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ZMOCUP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ZTERM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "REGBR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CIT_LEX"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "COD_FAM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DES_FAM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "COD_MOD"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DES_MOD"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "COD_MAR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DES_MAR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MVGR1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BEZEI"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ERNAM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ERDAT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ERZET"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TP_PARC_AT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CD_PARC_PROP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PARC_PROP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TP_PARC_PROP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DATA_FRQ"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "CD_PARC_COCKPIT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PARC_COCKPIT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TP_PARC_COCKPIT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "1_PARC_PROP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "1_DATA_FRQ"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VLR_ITM_REM"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VSBED"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/chkST_CARGA-CK_ORGANIZAR").Selected = True
session.findById("wnd[1]/usr/ctxtST_CARGA-TDLNR").Text = codTR
session.findById("wnd[1]/usr/ctxtST_CARGA-VSART").Text = xp
session.findById("wnd[1]/usr/ctxtST_CARGA-SDABW").Text = "23"
session.findById("wnd[1]/usr/ctxtVTTK-VSBED").Text = condexp
session.findById("wnd[1]/usr/chkST_CARGA-CK_ORGANIZAR").SetFocus
session.findById("wnd[1]/usr/btnSALVAR").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/btn*RV56A-ICON_STLBG").press
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1025/ctxtVTTK-DPLBG").SetFocus
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").SetFocus
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_TE/ssubG_HEADER_SUBSCREEN2:SAPMZV56A:1035/cmbVTTK-TNDRST").Key = "PB"
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[2]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

' Criar_Custo Macro

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvi01"
session.findById("wnd[0]").sendVKey 0
TR = session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text

session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").SetFocus
Custo = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").caretPosition = 21
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar Transporte").Select

ActiveCell.Offset(0, 7).Value = TR
ActiveCell.Offset(0, 8).Value = Custo

' ZSTR01_64 Macro

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar Transporte").Select
    
TR = ActiveCell.Offset(0, 7).Value
    
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
Sheets("Criar Transporte").Select
    
ActiveCell.Offset(0, 9).Value = "ZSTR OK"

' Notfis Macro

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = "0"
session.findById("wnd[0]/usr/radP_OPT2").SetFocus
session.findById("wnd[0]/usr/radP_OPT2").Select

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar Transporte").Select
    
    Dim Notfis, msg1, msg2

TR = ActiveCell.Offset(0, 7).Value

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

ActiveCell.Offset(0, 10).Value = Notfis
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation



End Sub
