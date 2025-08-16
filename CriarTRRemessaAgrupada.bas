Attribute VB_Name = "CriarTRRemessaAgrupada"
Sub Criar_TR_Remessa_Agrupada()
Attribute Criar_TR_Remessa_Agrupada.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select
Range("A2").Select

Dim TR, dt, codTR, condexp, OI
Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select

Deposito = ActiveCell.Offset(0, 1).Value
codTR = ActiveCell.Offset(0, 2).Value
xp = ActiveCell.Offset(0, 3).Value
DtRemessa = ActiveCell.Offset(0, 4).Value
condexp = ActiveCell.Offset(0, 5).Value

OI01 = Range("A2").Value
OI02 = Range("A3").Value
OI03 = Range("A4").Value
OI04 = Range("A5").Value
OI05 = Range("A6").Value
OI06 = Range("A7").Value
OI07 = Range("A8").Value
OI08 = Range("A9").Value
OI09 = Range("A10").Value
OI10 = Range("A11").Value
OI11 = Range("A12").Value
OI12 = Range("A13").Value
OI13 = Range("A14").Value
OI14 = Range("A15").Value
OI15 = Range("A16").Value
OI16 = Range("A17").Value
OI17 = Range("A18").Value
OI18 = Range("A19").Value
OI19 = Range("A20").Value
OI20 = Range("A21").Value
OI21 = Range("A22").Value
OI22 = Range("A23").Value
OI23 = Range("A24").Value
OI24 = Range("A25").Value
OI25 = Range("A26").Value
OI26 = Range("A27").Value
OI27 = Range("A28").Value
OI28 = Range("A29").Value
OI29 = Range("A30").Value
OI30 = Range("A31").Value
OI31 = Range("A32").Value
OI32 = Range("A33").Value
OI33 = Range("A34").Value
OI34 = Range("A35").Value
OI35 = Range("A36").Value
OI36 = Range("A37").Value
OI37 = Range("A38").Value
OI38 = Range("A39").Value
OI39 = Range("A40").Value
OI40 = Range("A41").Value
OI41 = Range("A42").Value
OI42 = Range("A43").Value
OI43 = Range("A44").Value
OI44 = Range("A45").Value
OI45 = Range("A46").Value
OI46 = Range("A47").Value
OI47 = Range("A48").Value
OI48 = Range("A49").Value
OI49 = Range("A50").Value
OI50 = Range("A51").Value

If OI01 = "" Then
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
session.findById("wnd[0]/usr/radP_TODOS").Select
session.findById("wnd[0]/usr/radP_TTTRN").Select
session.findById("wnd[0]/usr/radP_REVER").Select
session.findById("wnd[0]/usr/radP_REVER").SetFocus
session.findById("wnd[0]/usr/ctxtS_VSTEL-LOW").Text = Deposito
session.findById("wnd[0]/usr/ctxtS_ORDEM-LOW").caretPosition = 8
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").Text = DtRemessa
session.findById("wnd[0]/usr/btn%_S_ORDEM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = OI01
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI02
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI03
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI04
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI05
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI06
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI07
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI08
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 8
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI09
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI10
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI11
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI12
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI13
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI14
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI15
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 14
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI16
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI17
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI18
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI19
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI20
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI21
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI22
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 21
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI23
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI24
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI25
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI26
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI27
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI28
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI29
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 28
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI30
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI31
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI32
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI33
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI34
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI35
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI36
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 35
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI37
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI38
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI39
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI40
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI41
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI42
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI43
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 42
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = OI44
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = OI45
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = OI46
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = OI47
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = OI48
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = OI49
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = OI50
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_ORDEM-LOW").SetFocus
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
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
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

'Criar_Custo

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvi01"
session.findById("wnd[0]").sendVKey 0
TR = session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").SetFocus
custo01 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,0]").Text
custo02 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,1]").Text
custo03 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,2]").Text
custo04 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,3]").Text
custo05 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,4]").Text
custo06 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,5]").Text
custo07 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,6]").Text
custo08 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,7]").Text
custo09 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,8]").Text
custo10 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,9]").Text
custo11 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,10]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP").verticalScrollbar.Position = 10
custo12 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,1]").Text
custo13 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,2]").Text
custo14 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,3]").Text
custo15 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,4]").Text
custo16 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,5]").Text
custo17 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,6]").Text
custo18 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,7]").Text
custo19 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,8]").Text
custo20 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,9]").Text
custo21 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,10]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP").verticalScrollbar.Position = 19
custo22 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,1]").Text
custo23 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,2]").Text
custo24 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,3]").Text
custo25 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,4]").Text
custo26 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,5]").Text
custo27 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,6]").Text
custo28 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,7]").Text
custo29 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,8]").Text
custo30 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,9]").Text
custo31 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,10]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP").verticalScrollbar.Position = 28
custo32 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,1]").Text
custo33 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,2]").Text
custo34 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,3]").Text
custo35 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,4]").Text
custo36 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,5]").Text
custo37 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,6]").Text
custo38 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,7]").Text
custo39 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,8]").Text
custo40 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,9]").Text
custo41 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,10]").Text
session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP").verticalScrollbar.Position = 37
custo42 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,1]").Text
custo43 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,2]").Text
custo44 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,3]").Text
custo45 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,4]").Text
custo46 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,5]").Text
custo47 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,6]").Text
custo48 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,7]").Text
custo49 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,8]").Text
custo50 = session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-NETWR[4,9]").Text
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select
Range("A2").Select
linha2 = 2
COLUNA2 = 7
linha3 = 1
COLUNA3 = 1
While ActiveCell.Offset(0, 0).Value <> ""
If Cells(linha2, COLUNA2).Value = "" Then
Cells(linha2, COLUNA2).Value = TR
End If
linha2 = linha2 + 1
linha3 = linha3 + 1
Cells(linha2, COLUNA3).Select
Wend
linha2 = ""
COLUNA2 = ""
linha3 = ""
COLUNA3 = ""
Range("A2").Select

If OI01 <> "" Then
Range("H2").Value = custo01
End If
If OI02 <> "" Then
Range("H3").Value = custo02
End If
If OI03 <> "" Then
Range("H4").Value = custo03
End If
If OI04 <> "" Then
Range("H5").Value = custo04
End If
If OI05 <> "" Then
Range("H6").Value = custo05
End If
If OI06 <> "" Then
Range("H7").Value = custo06
End If
If OI07 <> "" Then
Range("H8").Value = custo07
End If
If OI08 <> "" Then
Range("H9").Value = custo08
End If
If OI09 <> "" Then
Range("H10").Value = custo09
End If
If OI10 <> "" Then
Range("H11").Value = custo10
End If
If OI11 <> "" Then
Range("H12").Value = custo11
End If
If OI12 <> "" Then
Range("H13").Value = custo12
End If
If OI13 <> "" Then
Range("H14").Value = custo13
End If
If OI14 <> "" Then
Range("H15").Value = custo14
End If
If OI15 <> "" Then
Range("H16").Value = custo15
End If
If OI16 <> "" Then
Range("H17").Value = custo16
End If
If OI17 <> "" Then
Range("H18").Value = custo17
End If
If OI18 <> "" Then
Range("H19").Value = custo18
End If
If OI19 <> "" Then
Range("H20").Value = custo19
End If
If OI20 <> "" Then
Range("H21").Value = custo20
End If
If OI21 <> "" Then
Range("H22").Value = custo21
End If
If OI22 <> "" Then
Range("H23").Value = custo22
End If
If OI23 <> "" Then
Range("H24").Value = custo23
End If
If OI24 <> "" Then
Range("H25").Value = custo24
End If
If OI25 <> "" Then
Range("H26").Value = custo25
End If
If OI26 <> "" Then
Range("H27").Value = custo26
End If
If OI27 <> "" Then
Range("H28").Value = custo27
End If
If OI28 <> "" Then
Range("H29").Value = custo28
End If
If OI29 <> "" Then
Range("H30").Value = custo29
End If
If OI30 <> "" Then
Range("H31").Value = custo30
End If
If OI31 <> "" Then
Range("H32").Value = custo31
End If
If OI32 <> "" Then
Range("H33").Value = custo32
End If
If OI33 <> "" Then
Range("H34").Value = custo33
End If
If OI34 <> "" Then
Range("H35").Value = custo34
End If
If OI35 <> "" Then
Range("H36").Value = custo35
End If
If OI36 <> "" Then
Range("H37").Value = custo36
End If
If OI37 <> "" Then
Range("H38").Value = custo37
End If
If OI38 <> "" Then
Range("H39").Value = custo38
End If
If OI39 <> "" Then
Range("H40").Value = custo39
End If
If OI40 <> "" Then
Range("H41").Value = custo40
End If
If OI41 <> "" Then
Range("H42").Value = custo41
End If
If OI42 <> "" Then
Range("H43").Value = custo42
End If
If OI43 <> "" Then
Range("H44").Value = custo43
End If
If OI44 <> "" Then
Range("H45").Value = custo44
End If
If OI45 <> "" Then
Range("H46").Value = custo45
End If
If OI46 <> "" Then
Range("H47").Value = custo46
End If
If OI47 <> "" Then
Range("H48").Value = custo47
End If
If OI48 <> "" Then
Range("H49").Value = custo48
End If
If OI49 <> "" Then
Range("H50").Value = custo49
End If
If OI50 <> "" Then
Range("H51").Value = custo50
End If

'ZSTR01_64

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select

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
Sheets("Criar TR Remessa Agrupada").Select
Range("A2").Select
linha2 = 2
COLUNA2 = 9
linha3 = 1
COLUNA3 = 1
While ActiveCell.Offset(0, 0).Value <> ""
If Cells(linha2, COLUNA2).Value = "" Then
Cells(linha2, COLUNA2).Value = "TABELA ZSTR OK"
End If
linha2 = linha2 + 1
linha3 = linha3 + 1
Cells(linha2, COLUNA3).Select
Wend
linha2 = ""
COLUNA2 = ""
linha3 = ""
COLUNA3 = ""
Range("A2").Select

'Notfis

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzstr44"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_TKNUM").Text = "0"
session.findById("wnd[0]/usr/radP_OPT2").SetFocus
session.findById("wnd[0]/usr/radP_OPT2").Select

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select
    
Dim Notfis, msg1, msg2

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

Windows("Planilha Reversa.xlsb").Activate
Sheets("Criar TR Remessa Agrupada").Select
Range("A2").Select
linha2 = 2
COLUNA2 = 10
linha3 = 1
COLUNA3 = 1
While ActiveCell.Offset(0, 0).Value <> ""
If Cells(linha2, COLUNA2).Value = "" Then
Cells(linha2, COLUNA2).Value = Notfis
End If
linha2 = linha2 + 1
linha3 = linha3 + 1
Cells(linha2, COLUNA3).Select
Wend
linha2 = ""
COLUNA2 = ""
linha3 = ""
COLUNA3 = ""
Range("A2").Select

TR = ""

fim:
MsgBox "Finalizado.", vbInformation

End Sub
