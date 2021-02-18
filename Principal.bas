Attribute VB_Name = "Principal"
'Public Cn As New ADODB.Connection
Public Command                                              As String
Public Caminho                                              As String
Public FrmTelaLogin                                         As Boolean
Public CaminhoBanco                                         As String
Public Aminacao                                             As String
Public LocalImagem                                          As String
'Public Usuario As Usuario_01
Public Comando                                              As String
Public Cores                                                As TipoCores
Public Html                                                 As String
Public Conexao                                              As String
Public Versao                                               As String
Public Sql                                                  As String
Public SenhaSistema                                         As String
Public BancoDeDados                                         As DAO.Database
Public LocalBanco                                           As String
Public Usuario                                              As Usuario_Class
Public ShowTipoCliente                                      As Long
Public ShowFornecedor                                       As Long


Type TipoCores
    Azul                                                    As String
    Preto                                                   As String
    Branco                                                  As String
    Vermelho                                                As String
    Verde                                                   As String
    Amarelho                                                As String
    AmrelhoClaro                                            As String
    DeskTop                                                 As String
    Padrão                                                  As String

End Type

Type Usuario_01
    Nome                                                    As String
    Login                                                   As String
    Senha                                                   As String
    Acesso                                                  As String
    Data                                                    As Date
    Hora                                                    As String
    Id                                                      As String
End Type
Type Usuario_Class
    CadastrodeCliente                                       As Boolean
    CadastrodeProduto                                       As Boolean
    CadastrodeVendedor                                      As Boolean
    CadastrodeTipodeCliente                                 As Boolean
    CadastrodeCentrodeCusto                                 As Boolean
    CadastrodeUsuario                                       As Boolean
    Cheque                                                  As Boolean
    Unidade                                                 As Boolean
    Entradanoestoque                                        As Boolean
    CondicaodePagamento                                     As Boolean
    CadastroFornecedor                                      As Boolean
    GrupodeProdutos                                         As Boolean
    BalcaodeVendas                                          As Boolean
    ListagemdeProdutos                                      As Boolean
    RelatoriodeVendas                                       As Boolean
    ListagemdeVendasporVendedor                             As Boolean
    Nome                                                    As String
    Id                                                      As Long
    Login                                                   As String
    Senha                                                   As String
    DataHoraLogin                                           As String

End Type

Dim strNovoSQl                                              As String
Dim strCampo()                                              As String
Dim strValores()                                            As Variant
Public Function NovoEdit(strParamentro As String)
    strNovoSQl = strParamentro
    ReDim strCampo(0) As String
    ReDim strValores(0) As Variant
End Function

Public Function Edit(pCampo As String, pValor As Variant)
    Dim lngUltimo As Long
    
    lngUltimo = UBound(strCampo) + 1
    ReDim Preserve strCampo(lngUltimo) As String
    ReDim Preserve strValores(lngUltimo) As Variant

    strCampo(lngUltimo) = pCampo
    strValores(lngUltimo) = pValor
End Function
Public Function BuscaClienteRel(lngId As Long) As String
    On Error Resume Next
    BuscaClienteRel = BancoDeDados.OpenRecordset("select Razao from cliente where id =" & lngId).Fields("Razao").Value
End Function

Public Function buscaVendedor(lngId As Long) As String
    On Error Resume Next
    buscaVendedor = BancoDeDados.OpenRecordset("select Razao from vendedor where id =" & lngId).Fields("Razao").Value
End Function


Public Function MontaSql(pstrTabela As String)
    Dim i                                                   As Long
    Dim strSql                                              As String
    Dim strSqlParte2                                        As String
    If Trim(strNovoSQl) = "" Then
        strSql = "insert into " & pstrTabela & "("
        strSqlParte2 = ""
        For i = 0 To UBound(strCampo)
            strSql = strSql & strCampo(i) & ", "
            strSqlParte2 = strSqlParte2 & "'" & strValores(i) & "', "
        Next
        strSqlParte2 = Left(strSqlParte2, Len(strSqlParte2) - 2) & ") "
        strSql = Left(strSql, Len(strSql) - 2) & ") values (" & strSqlParte2
    Else
        strSql = "update " & pstrTabela & " Set "
        For i = 0 To UBound(strCampo)
            strSql = strSql & strCampo(i) & " = '" & strValores(i) & ", "
        Next
        strSql = Left(strSql, Len(strSql) - 2) & "  where " & strNovoSQl
    End If
    BancoDeDados.Execute strSql

End Function

Public Sub Main()
    Dim Provider                                            As String

    Versao = App.Major & "." & App.Minor & "." & App.Revision

    If Right(App.Path, 1) = "\" Then
        CaminhoBanco = Ler("Banco", "Arquivo", "", App.Path & "cadcon.ini")
        Aminacao = IIf(Ler("Tela", "Aminação", "1", App.Path & "cadcon.ini") = "1", True, False)
        LocalImagem = Ler("Tela", "Log", "", App.Path & "cadcon.ini")
    Else
        CaminhoBanco = Ler("Banco", "Arquivo", "", App.Path & "\cadcon.ini")
        Aminacao = IIf(Ler("Tela", "Aminação", "1", App.Path & "\cadcon.ini") = "1", True, False)
        LocalImagem = Ler("Tela", "Log", "", App.Path & "\cadcon.ini")
    End If

    If Right(CaminhoBanco, 1) <> "\" Then CaminhoBanco = CaminhoBanco & "\"


    If CaminhoBanco = "" Or Dir(CaminhoBanco) = "" Then
        MsgBox "Caro Usuario, Impossivel localizar o banco de dados", vbCritical, App.Title
        End
    End If

    Provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CaminhoBanco & "cadcon2000.mdb"

    Cn.Open Provider

    frmSplash.Show

End Sub

Public Function E()
    Dim msg                                                 As String

    If Err.Number > 0 Then

        msg = "Caro Usuario, Ocorreu um erro" & Chr(13)
        msg = msg & "Numero do Erro :" & Err.Number & Chr(13)
        msg = msg & "Descrição : " & Err.Description & Chr(13) & Chr(13)
        msg = msg & "Favor entrar em contato com o suporte tecnico." & Chr(13)
        msg = msg & "31 8806-5049 ou suporte@neobh.com.br"

        MsgBox msg, vbCritical, "Atenção"

    End If

End Function

Public Function Erro()
    Dim msg                                                 As String

    If Err.Number > 0 Then

        msg = "Caro Usuario, Ocorreu um erro" & Chr(13)
        msg = msg & "Numero do Erro :" & Err.Number & Chr(13)
        msg = msg & "Descrição : " & Err.Description & Chr(13) & Chr(13)
        msg = msg & "Favor entrar em contato com o suporte tecnico." & Chr(13)
        msg = msg & "31 8806-5049 ou suporte@neobh.com.br"

        MsgBox msg, vbCritical, "Atenção"

    End If

End Function

Function Centra(Frm As Form)
    Frm.Top = (FrmPrincipal.ScaleHeight / 2) - (Frm.Height / 2)
    Frm.Left = (FrmPrincipal.ScaleWidth / 2) - (Frm.Width / 2)
End Function

Public Function Num(KeyAscii As Integer)
    On Error GoTo Trata_Erro
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then

    ElseIf KeyAscii = vbKeyBack Then

    ElseIf KeyAscii = 46 Then
        KeyAscii = 44
    Else
        KeyAscii = 0
    End If
    Num = KeyAscii
Trata_Erro:
    E
End Function

Public Function Valida(NomeData As MaskEdBox, Optional msg As Boolean)
    On Error GoTo Trata_Erro
    Dim Data                                                As String

    Data = NomeData.Text
    If Right(Data, 4) = "____" Then
        Data = Left(Data, Len(Data) - 4) & IIf(Len("" & Year(Date)) = 2, "20" & Year(Date), Year(Date))
    End If
    If Mid(Data, 4, 2) = "__" Then
        Data = Left(Data, 3) & Month(Date) & Right(Data, 5)
    End If
    If Left(Data, 2) = "__" Then
        Data = Day(Date) & Right(Data, Len(Data) - 2)
    End If
    Data = Replace(Data, "_", "")

    Data = Format(Data, "DD/MM/YYYY")

    If msg = False Then
        If Not IsDate(Data) Then
            MsgBox "Data em formato invalido, favor Digitar novamente!", vbInformation, "Atenção"
            Valida = "__/__/____"
            NomeData.SetFocus
        Else
            Valida = Data
        End If
    Else
        Valida = Data
    End If
Trata_Erro:
    E
End Function

Public Function Calc_CGC(Valor As String) As Boolean
    Dim Mult1                                               As String
    Dim Mult2                                               As String
    Dim dig1                                                As Integer
    Dim dig2                                                As Integer
    Dim X                                                   As Integer
    Mult1 = "543298765432"
    Mult2 = "6543298765432"
    For X = 1 To 12
        dig1 = dig1 + (Val(Mid$(Valor, X, 1)) * Val(Mid$(Mult1, X, 1)))
    Next
    For X = 1 To 13
        dig2 = dig2 + (Val(Mid$(Valor, X, 1)) * Val(Mid$(Mult2, X, 1)))
    Next
    dig1 = (dig1 * 10) Mod 11
    dig2 = (dig2 * 10) Mod 11
    If dig1 = 10 Then
        dig1 = 0
    End If
    If dig2 = 10 Then
        dig2 = 0
    End If
    Calc_CGC = True
    If dig1 <> Val(Mid$(Valor, 13, 1)) Then
        Calc_CGC = False
    End If
    If dig2 <> Val(Mid$(Valor, 14, 1)) Then
        Calc_CGC = False
    End If
End Function
Public Function Calc_CPF(Valor As String) As Boolean
'Inicializa variaveis
    Dim dig1                                                As Integer
    Dim dig2                                                As Integer
    Dim Mult1                                               As Integer
    Dim Mult2                                               As Integer
    Dim X                                                   As Integer
    Mult1 = 10
    Mult2 = 11
    For X = 1 To 9
        dig1 = dig1 + (Val(Mid$(Valor, X, 1)) * Mult1)
        Mult1 = Mult1 - 1
    Next
    For X = 1 To 10
        dig2 = dig2 + (Val(Mid$(Valor, X, 1)) * Mult2)
        Mult2 = Mult2 - 1
    Next
    dig1 = (dig1 * 10) Mod 11
    dig2 = (dig2 * 10) Mod 11
    If dig1 = 10 Then
        dig1 = 0
    End If
    If dig2 = 10 Then
        dig2 = 0
    End If
    Calc_CPF = True
    If Val(Mid$(Valor, 10, 1)) <> dig1 Then Calc_CPF = False
    If Val(Mid$(Valor, 11, 1)) <> dig2 Then Calc_CPF = False
End Function

Public Sub Pause(HowManySecs)
'   pause for HowManySecs seconds
    Dim EndWait
    EndWait = DateAdd("s", HowManySecs, Now)
    While Now < EndWait
        '        this is dummy text...nothing is actually done during the wait
    Wend

End Sub

Public Function Cel(Txt As TextBox)
    On Error GoTo Trata_Erro
    Txt.SelStart = 0
    Txt.SelLength = Len(Txt.Text)
Trata_Erro:
    E
End Function

Public Function MemoRead(Arquivo As String)
    On Error GoTo Trata_Erro
    Dim Texto                                               As String

    If Dir(Arquivo) = "" Then
        MsgBox "Impossivel Localizar o Arquivo de Configuração do Sistema", vbExclamation, App.Title
        Exit Function
    End If

    Texto = String(FileLen(Arquivo), " ")

    Open Arquivo For Binary As #1
    Get #1, , Texto
    Close #1
    MemoRead = Texto

Trata_Erro:
    E
End Function

Public Function EstadoCobo(Box As ComboBox)
    On Error GoTo Trata_Erro
    Box.Clear
    Box.AddItem "Ac"
    Box.AddItem "AL"
    Box.AddItem "AM"
    Box.AddItem "AP"
    Box.AddItem "BA"
    Box.AddItem "CE"
    Box.AddItem "DF"
    Box.AddItem "ES"
    Box.AddItem "GO"
    Box.AddItem "MA"
    Box.AddItem "MG"
    Box.AddItem "MS"
    Box.AddItem "MT"
    Box.AddItem "Pa"
    Box.AddItem "PB"
    Box.AddItem "PE"
    Box.AddItem "PI"
    Box.AddItem "PR"
    Box.AddItem "RJ"
    Box.AddItem "RN"
    Box.AddItem "RO"
    Box.AddItem "RR"
    Box.AddItem "RS"
    Box.AddItem "SC"
    Box.AddItem "SE"
    Box.AddItem "SP"
    Box.AddItem "TO"
Trata_Erro:
    E
End Function

'Public Function ValidaCep(CCep As TextBox, Estado)
'On Error GoTo Trata_Erro
'Dim Cep1() As String
'Dim Cep2() As String
'ReDim Cep1(27) As String
'ReDim Cep2(27) As String
'Dim Tmp As Neo
'
'Cep1(1) = "RO"
'Cep1(2) = "AC"
'Cep1(3) = "AM"
'Cep1(4) = "RR"
'Cep1(5) = "PA"
'Cep1(6) = "AP"
'Cep1(7) = "TO"
'Cep1(8) = "MA"
'Cep1(9) = "PI"
'Cep1(10) = "CE"
'Cep1(11) = "RN"
'Cep1(12) = "PB"
'Cep1(13) = "PE"
'Cep1(14) = "AL"
'Cep1(15) = "SE"
'Cep1(16) = "BA"
'Cep1(17) = "MG"
'Cep1(18) = "ES"
'Cep1(19) = "RJ"
'Cep1(20) = "SP"
'Cep1(21) = "PR"
'Cep1(22) = "SC"
'Cep1(23) = "RS"
'Cep1(24) = "MS"
'Cep1(25) = "MT"
'Cep1(26) = "GO"
'Cep1(27) = "DF"
'
'Cep2(1) = "7890078999"
'Cep2(2) = "6990069999"
'Cep2(3) = "6900069899"
'Cep2(4) = "6930069399"
'Cep2(5) = "6600068899"
'Cep2(6) = "6890068999"
'Cep2(7) = "7700077999"
'Cep2(8) = "6500065999"
'Cep2(9) = "6400064999"
'Cep2(10) = "6000063999"
'Cep2(11) = "5900059999"
'Cep2(12) = "5800058999"
'Cep2(13) = "5000056999"
'Cep2(14) = "5700057999"
'Cep2(15) = "4900049999"
'Cep2(16) = "4000048999"
'Cep2(17) = "3000039999"
'Cep2(18) = "2900029999"
'Cep2(19) = "2000028999"
'Cep2(20) = "0100019999"
'Cep2(21) = "8000087999"
'Cep2(22) = "8800089999"
'Cep2(23) = "9000099999"
'Cep2(24) = "7900079999"
'Cep2(25) = "7800078999"
'Cep2(26) = "7280076799"
'Cep2(27) = "7000073399"
'
'Tmp.Xy = 1
'Tmp.Texto = "Cep Invalido, Este Cep Não Pertece ao Estado Informado ! ! !"
'Do While True
'    If UCase(Estado) = UCase(Cep1(Tmp.Xy)) Then
'        Tmp.Valor1 = Left(Cep2(Tmp.Xy), 5)
'        Tmp.Valor2 = Right(Cep2(Tmp.Xy), 5)
'        If Val(Left(CCep.Text, 5)) < Tmp.Valor1 Then
'            MsgBox Tmp.Texto, vbInformation, App.Title
'            Exit Function
'        End If
'        If Val(Left(CCep.Text, 5)) > Tmp.Valor2 Then
'            MsgBox Tmp.Texto, vbInformation, App.Title
'            Exit Function
'        End If
'    End If
'    Tmp.Xy = Tmp.Xy + 1
'    If Tmp.Xy = 28 Then Exit Do
'Loop
'Trata_Erro:
'E
'End Function

Public Sub CarCores()
    On Error GoTo Trata_Erro

    Cores.Amarelho = &HFFFF&
    Cores.AmrelhoClaro = &H80000018
    Cores.Azul = &H800000
    Cores.Branco = &HFFFFFF
    Cores.DeskTop = &H80000001
    Cores.Padrão = &HC0C0C0
    Cores.Preto = &H0&
    Cores.Verde = &H8000&
    Cores.Vermelho = &HC0&

Trata_Erro:
    E
End Sub

Public Function LimparCampos(Frm As Form)
    Dim X                                                   As Long

    For X = 0 To Frm.Count


    Next X

End Function


Public Function Resp(VResp, VLong)
    On Error GoTo Trata_Erro
    Dim r                                                   As String
    Select Case VResp
    Case 0
        MsgBox "Matricula inicial inexistente", vbCritical, App.Title
    Case 1
        MsgBox "Matricula final inexistente", vbCritical, App.Title
    Case 2
        MsgBox "Parametro nao Cadastrado.", vbCritical, App.Title
    Case 3
        MsgBox "Movimento incoerente.", vbCritical, App.Title
    Case 4
        MsgBox "Valores incoerentes.", vbCritical, App.Title
    Case 5
        MsgBox "Funcionario ja demitido. ", vbCritical, App.Title
    Case 6
        MsgBox "Funcionario nao demitido. ", vbCritical, App.Title
    Case 7
        MsgBox "Digito verificador nao confere. ", vbCritical, App.Title
    Case 8
        MsgBox "Titulo ja foi baixado. ", vbCritical, App.Title
    Case 9
        MsgBox "Titulo ja foi cancelado. ", vbCritical, App.Title
    Case 10
        MsgBox "Titulo ja foi lancado. ", vbCritical, App.Title
    Case 11
        MsgBox "Cheque foi emitido. Fazer cancelamento. ", vbCritical, App.Title
    Case 12
        MsgBox "Cheque ja foi cancelado. ", vbCritical, App.Title
    Case 13
        MsgBox "Cheque nao foi emitido fazer exclusao. ", vbCritical, App.Title
    Case 14
        MsgBox "Acesso nao autorizado. ", vbCritical, App.Title
    Case 15
        MsgBox "Senha nao confere. ", vbCritical, App.Title
    Case 16
        MsgBox "Registro inexistente. ", vbCritical, App.Title
    Case 17
        MsgBox "Registro existente. ", vbCritical, App.Title
    Case 18
        MsgBox "Registro gravado com codigo => " + Str(VLong), vbInformation, App.Title
    Case 19
        MsgBox "Final de Arquivo. ", vbCritical, App.Title
    Case 20
        Resp = MsgBox("Confirma Deleção?", vbCritical + vbYesNo + vbDefaultButton2, App.Title)
    Case 21
        Resp = MsgBox("Confirma Dados ?", vbInformation + vbYesNo, App.Title)
    Case 22
        Resp = MsgBox("O Banco de Dados da Empresa será Excluido!" + Chr(13) + "Deseja Continuar? ", vbCritical + vbYesNo + vbDefaultButton2)
    Case 23
        Resp = MsgBox("Não foi Possivel organizar as Tabelas !", vbCritical, App.Title)
    Case 24
        Resp = MsgBox("Não foi possivel Encontrar o Relatorio" + Chr(13) + Chr(13) + VLong, vbCritical, App.Title)
    Case 25
        Resp = MsgBox("Não foi possivel Encontrar o Back-Up", vbCritical, App.Title)
    Case 26
        Resp = MsgBox("Data em Formanto invalido !!!", vbCritical, App.Title)
    Case 27
        Resp = MsgBox("Acesso Negado ! ! !", vbCritical + vbApplicationModal, App.Title + " ! ! !")
    Case 28
        Resp = MsgBox("Periodo Contabil Invalido ! ! !", vbCritical, App.Title)
    Case 29
        Resp = MsgBox("CNPJ Invalido!", vbCritical, App.Title)
    Case 30
        Resp = MsgBox("Cpf Invalido!", vbCritical, App.Title)
    Case 31
        Resp = MsgBox("Centro de Custro Invalido!!!", vbCritical, App.Title)
    Case 32
        Resp = MsgBox("Classificação incoerente falta conta titulo", vbCritical, App.Title)
    Case 33
        Resp = MsgBox("Não a Registro a Ser Listado!!!", vbCritical, App.Title)
    Case 34
        Resp = MsgBox("Filial Invalida!", vbCritical, App.Title)
    Case 35
        Resp = MsgBox("Conta Invalida ! ! !", vbCritical, App.Title)
    Case 36
        MsgBox "Relatorio Invalido ! ! !", vbCritical, App.Title
    Case 37
        MsgBox "Plano de Conta ja Cadastrada com Esta Classificação ! ! !", vbCritical, App.Title
    Case 38
        MsgBox "Não a Plano de Conta no Balancente", vbSystemModal + vbCritical, App.Title
    End Select
Trata_Erro:
    E
End Function
Public Function Alert(Mensagem As String, Optional Tipo As VbMsgBoxStyle) As VbMsgBoxResult

    If Tipo = 0 Then Tipo = vbCritical

    Alert = MsgBox(Mensagem, Tipo)

End Function
Public Function MontaEmpresa()

End Function
