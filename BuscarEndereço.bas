Attribute VB_Name = "BuscarEndereço"
Sub Buscar_N_endereço()
Attribute Buscar_N_endereço.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Buscar_N Macro
'


Windows("Planilha Reversa.xlsb").Activate
    Sheets("Endereço").Select
    
    Final = Range("C10000").End(xlUp).Row + 1
    
    Range("A" & Final).Select
 

nl = Application.WorksheetFunction.CountA(Range("A:A")) - 1

For i = 0 To nl

OI = ActiveCell.Offset(0, 0).Value
vazio = ActiveCell.Offset(0, 0).Value

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
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OI
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").Select

linha0 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,0]").Text
linha1 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,1]").Text
linha2 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,2]").Text
linha3 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,3]").Text
linha4 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").Text
linha5 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").Text
linha6 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,6]").Text

If linha0 = "WE Recebedor da mercadoria" Then
GoTo segue0
End If

If linha1 = "WE Recebedor mercadoria" Then
GoTo segue1
End If

If linha2 = "WE Recebedor mercadoria" Then
GoTo segue2
End If

If linha3 = "WE Recebedor mercadoria" Then
GoTo segue3
End If

If linha4 = "WE Recebedor mercadoria" Then
GoTo segue4
End If

If linha5 = "WE Recebedor mercadoria" Then
GoTo segue5
End If

If linha6 = "WE Recebedor mercadoria" Then
GoTo segue6
End If

segue0:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,0]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue1:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue2:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,2]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue3:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,3]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue4:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue5:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,5]").SetFocus
session.findById("wnd[0]").sendVKey 2
GoTo segue7

segue6:

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").SetFocus
session.findById("wnd[0]").sendVKey 2

segue7:

rua = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-STREET").Text
numero = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").Text
'CPF = session.findById("wnd[1]/usr/subGCS_ATTRIBUTES:SAPLV09C:5001/txtGDF_STCD2").Text

Windows("Planilha Reversa.xlsb").Activate
    Sheets("Endereço").Select
    
ActiveCell.Offset(0, 1).Value = rua
ActiveCell.Offset(0, 2).Value = numero
'ActiveCell.Offset(0, 3).Value = CPF
ActiveCell.Offset(1, 0).Select

Next

fim:
MsgBox "Finalizado.", vbInformation

'
End Sub

Sub Buscar_Peso()
'
' Buscar peso bruto e liquido
'
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

    Sheets("Buscar Peso").Select
    Final = Range("F100000").End(xlUp).Row + 1
        
    Range("B" & Final).Select
    

Do Until Range("B" & Final).Select = ""
SKU = ActiveCell.Offset(0, 0).Value
qtd = ActiveCell.Offset(0, 1).Value
vazio = ActiveCell.Offset(0, 0).Value

If vazio = "" Then
GoTo fim
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmm03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = SKU
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[18]").press
material = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1005/txtMAKT-MAKTX").Text
bruto = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2703/txtMARA-BRGEW").Text
liquido = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2703/txtMARA-NTGEW").Text
volume = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2703/txtMARA-VOLUM").Text
unidade_item = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2703/ctxtMARA-GEWEI").Text
unidade_volume = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2703/ctxtMARA-VOLEH").Text

total1 = bruto * qtd
If total1 = "0" Then
total1 = bruto * 1
End If

total2 = liquido * qtd
If total2 = "0" Then
total2 = liquido * 1
End If

total3 = volume * qtd
If total3 = "0" Then
total3 = volume * 1
End If

Windows("Planilha Reversa.xlsb").Activate
    Sheets("Buscar Peso").Select
    
ActiveCell.Offset(0, -1).Value = material
ActiveCell.Offset(0, 2).Value = total1
ActiveCell.Offset(0, 3).Value = total2
ActiveCell.Offset(0, 4).Value = unidade_item
ActiveCell.Offset(0, 5).Value = total3
ActiveCell.Offset(0, 6).Value = unidade_volume
    
ActiveCell.Offset(0, 7).Value = "OK"
'ActiveCell.Offset(1, 0).Select

Final = Final + 1

total1 = 0
total2 = 0
total3 = 0

Loop

fim:
MsgBox "Finalizado.", vbInformation
'
End Sub

