Attribute VB_Name = "ExcluirTR"
Sub EXCLUIR_TR()

Application.ScreenUpdating = False
Windows("Planilha Reversa.xlsb").Activate
Sheets("Alteração Geral").Select
Range("A1").Select

    QTYLINHAS = Range("B10000").End(xlUp).Row
    ActiveSheet.Range("$A$1:$C$10000").RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    QTYLINHAS = ""

    QTYLINHAS3 = Range("B10000").End(xlUp).Row
    QTYLINHAS2 = Range("H10000").End(xlUp).Row
    QTYLINHAS3 = QTYLINHAS2 + 1
    Range("B" & QTYLINHAS3).Select

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
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC").Select
session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_FC/ssubG_HEADER_SUBSCREEN1:SAPMZV56A:1028/btnSCD_DISPLAY_1").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nyt02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = TR
session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press

Windows("Planilha Reversa.xlsb").Activate
Sheets("Alteração Geral").Select
ActiveCell.Offset(0, 6).Value = "Transporte Excluído"
ActiveCell.Offset(1, 0).Select

Next

fim:
session.findById("wnd[0]").sendVKey 12

End Sub
Sub EXCLUIR_REMESSA()

    ' Desativa a atualização da tela para acelerar a macro
    Application.ScreenUpdating = False
    
    ' Ativa a planilha correta
    Windows("Planilha Reversa.xlsb").Activate
    Sheets("Alteração Geral").Select
    
    ' Remove linhas duplicadas com base nas colunas A, B e C
    ActiveSheet.Range("$A$1:$C$10000").RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    
    ' Seleciona a primeira célula de dados para iniciar o loop (C2)
    Range("C2").Select

    ' Inicia o loop que continuará enquanto houver valores na coluna C
    Do While ActiveCell.Value <> ""
    
        ' Verifica se a coluna I (offset 6 da coluna C) NÃO contém "Remessa Excluída"
        If ActiveCell.Offset(0, 6).Value <> "Remessa Excluída" Then
        
            ' Pega o valor da remessa da célula ativa na coluna C
            Remessa = ActiveCell.Value
            
            ' Bloco de código para conexão e automação do SAP
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
            
            ' Bloco de código para executar a transação no SAP
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Remessa
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[1]/btn[14]").press
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

            ' Ativa o Excel novamente
            Windows("Planilha Reversa.xlsb").Activate
            Sheets("Alteração Geral").Select
            
            ' Escreve "Remessa Excluída" na coluna I (offset 6) para marcar que foi processada
            ActiveCell.Offset(0, 6).Value = "Remessa Excluída"
            
        End If ' Fim da condição de verificação

        ' Move para a próxima linha na coluna C para continuar o loop
        ActiveCell.Offset(1, 0).Select
        
    Loop ' Repete o processo para a próxima linha

    ' Fecha a janela da transação no SAP ao final do processo
    session.findById("wnd[0]").sendVKey 12
    
    ' Reativa a atualização da tela
    Application.ScreenUpdating = True

End Sub


Sub EXCLUIR_TRANSPORTE_REMESSA()

Application.ScreenUpdating = False
Inicio:
TIPO = InputBox("Qual informação deseja excluir, (Remessa) ou (TR)?")
If TIPO = Empty Then
    MsgBox ("Escolha uma das opções."), vbCritical
    GoTo Inicio
End If

TIPO = UCase(TIPO)

If TIPO = "REMESSA" Then
    Call EXCLUIR_REMESSA
End If

If TIPO = "TR" Then
    Call EXCLUIR_TR
End If

MsgBox "Finalizado.", vbInformation

End Sub


