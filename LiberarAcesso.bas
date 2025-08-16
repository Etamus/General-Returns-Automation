Attribute VB_Name = "LiberarAcesso"
Public tempoLimite As Date
Public acessoAtivo As Boolean
Public usuarioLogado As String

Sub LiberarAcesso()
    Dim ws As Worksheet
    Dim abasLiberadas As Variant
    Dim login As String
    Dim senha As String

    frmLogin.Show

    If frmLogin.txtLogin = "" Or frmLogin.txtSenha = "" Then
        MsgBox "Login ou senha não preenchidos.", vbExclamation
        Exit Sub
    End If

    login = LCase(Trim(frmLogin.txtLogin))
    senha = frmLogin.txtSenha
    Unload frmLogin
    usuarioLogado = login

    Select Case True
        Case login = "nascia35" And senha = "120995"
            tempoLimite = Now + TimeSerial(1, 0, 0)
            MsgBox "Acesso liberado por 1 hora.", vbInformation
            abasLiberadas = Array("Liberar Acesso", "Alteração Geral", "Alterar Remessa, OI ou TR", "Cancelar Ordem", _
                                  "Lançar Providência", "Buscar Peso", "Buscar Chave de Acesso e Mlog", "Endereço")
        
        Case login = "santot20" And senha = "130817"
            tempoLimite = Now + TimeSerial(1, 0, 0)
            MsgBox "Acesso liberado por 1 hora.", vbInformation
            abasLiberadas = Array("Liberar Acesso", "Alteração Geral", "Alterar Remessa, OI ou TR", "Cancelar Ordem", _
                                  "Lançar Providência", "Buscar Peso", "Buscar Chave de Acesso e Mlog", "Endereço")
                                  
        Case login = "anjosl6" And senha = "qb7p7Z001UQTwL"
            tempoLimite = Now + TimeSerial(1, 0, 0)
            MsgBox "Acesso liberado por 1 hora.", vbInformation
            abasLiberadas = Array("Liberar Acesso", "Alteração Geral", "Alterar Remessa, OI ou TR", "Cancelar Ordem", _
                                  "Lançar Providência", "Buscar Peso", "Buscar Chave de Acesso e Mlog", "Endereço")
                                  
        Case login = "lopesm21" And senha = "Whirlpoolcsi@2025"
            tempoLimite = Now + TimeSerial(1, 0, 0)
            MsgBox "Acesso liberado por 1 hora.", vbInformation
            abasLiberadas = Array("Liberar Acesso", "Alteração Geral", "Criar Transporte", "Alterar Remessa, OI ou TR", _
                                  "Cancelar Ordem", "Alterar RFQ")

        Case login = "santof78" And senha = "maurilo26"
            tempoLimite = Now + TimeSerial(1, 0, 0)
            MsgBox "Acesso liberado por 1 hora.", vbInformation
            abasLiberadas = Array("Liberar Acesso", "Alteração Geral", "Criar Transporte", "Alterar Remessa, OI ou TR", _
                                  "Cancelar Ordem", "Alterar RFQ")

        Case login = "admin" And senha = "admin"
            tempoLimite = Now + TimeSerial(5, 0, 0)
            MsgBox "Acesso liberado por 5 horas.", vbInformation
            abasLiberadas = Array() ' todas, exceto CDC

        Case Else
            MsgBox "Login ou senha incorretos.", vbCritical
            Exit Sub
    End Select

    acessoAtivo = True
    Application.OnTime Now + TimeSerial(0, 0, 5), "MonitorarAcesso"

    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "CDC" Then
            ws.Visible = xlSheetVeryHidden
        ElseIf login = "admin" Then
            ws.Visible = xlSheetVisible
        Else
            If Not IsError(Application.Match(ws.Name, abasLiberadas, 0)) Then
                ws.Visible = xlSheetVisible
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next ws

    Sheets("Alteração Geral").Select
    Application.ScreenUpdating = True
End Sub

Sub MonitorarAcesso()
    Dim ws As Worksheet

    If acessoAtivo Then
        If Now >= tempoLimite Then
            Application.ScreenUpdating = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> "Liberar Acesso" Then
                    ws.Visible = xlSheetVeryHidden
                End If
            Next ws
            Application.ScreenUpdating = True

            MsgBox "Tempo de acesso expirado.", vbExclamation
            acessoAtivo = False
        Else
            Application.OnTime Now + TimeSerial(0, 0, 5), "MonitorarAcesso"
        End If
    End If
End Sub


