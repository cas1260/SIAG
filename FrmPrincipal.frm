VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SIAG - Sistema de Informações e Administração Gerencial"
   ClientHeight    =   5910
   ClientLeft      =   1815
   ClientTop       =   2010
   ClientWidth     =   8055
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "29/10/2010"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "11:25"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            MinWidth        =   2646
            Text            =   "Usuario :"
            TextSave        =   "Usuario :"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6562
            MinWidth        =   5292
            Text            =   "Empresa :"
            TextSave        =   "Empresa :"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Left            =   1320
      Top             =   1320
   End
   Begin MSComctlLib.ImageList MList 
      Left            =   2250
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":191E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":21FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":264E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":2AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":337E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":405A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":4936
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnuCadClinete 
         Caption         =   "Cliente"
      End
      Begin VB.Menu mnuCadProduo 
         Caption         =   "Produto"
      End
      Begin VB.Menu mnuCentroCursto 
         Caption         =   "Centro de Custo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCondPag 
         Caption         =   "Condições de Pagamento"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFornecedor 
         Caption         =   "Fornecedor"
      End
      Begin VB.Menu mnuGrupoProd 
         Caption         =   "Grupo de Produto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTipoCliente 
         Caption         =   "Tipo de Cliente"
      End
      Begin VB.Menu mnuTipoDesp 
         Caption         =   "Tipo de Despesas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCadVendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnuUsuaruio 
         Caption         =   "Usuario"
      End
      Begin VB.Menu MenuUnidade 
         Caption         =   "Unidade"
      End
   End
   Begin VB.Menu mnuMov 
      Caption         =   "Movimentação"
      Begin VB.Menu mnuBalcaoVendas 
         Caption         =   "Balção de de Vendas"
      End
      Begin VB.Menu mnuCheques 
         Caption         =   "Cheques"
      End
      Begin VB.Menu mnuEntradasProduto 
         Caption         =   "Entradas de Produto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDespensas 
         Caption         =   "Despensas "
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatorio"
      Begin VB.Menu mnuListaProd 
         Caption         =   "Listagem de Produto"
      End
      Begin VB.Menu mnRelaVendas 
         Caption         =   "Relatorio de Vendas / Caixa"
      End
      Begin VB.Menu mnuListaGemVendas 
         Caption         =   "Listagem de Vendas por Vendedor"
      End
      Begin VB.Menu CmdHisto 
         Caption         =   "Historio de Cliente"
      End
   End
   Begin VB.Menu mnuDiversos 
      Caption         =   "Diversos"
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre"
      End
      Begin VB.Menu mnubackUp 
         Caption         =   "Back-Up"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Back - Up"
      End
      Begin VB.Menu mnuDadosDaEmpresa 
         Caption         =   "Dados da Empresa"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MudarDesenho                                            As Boolean
Dim XX As Long, YY As Long, PassoX As Boolean, PassoY       As Boolean
Attribute YY.VB_VarUserMemId = 1073938433
Attribute PassoX.VB_VarUserMemId = 1073938433
Attribute PassoY.VB_VarUserMemId = 1073938433

Private Sub CmdHisto_Click()
1   FrmhistoricoCliente.Show
End Sub

Private Sub MDIForm_Activate()
1   If FrmTelaLogin = False Then
2       FrmSenha.Show 1
3       FrmTelaLogin = True
4       Sbar.Panels(3).Text = Usuario.Nome
5       MontaEmpresa
6   End If
7   mnuCadClinete.Enabled = Usuario.CadastrodeCliente
8   mnuCadProduo.Enabled = Usuario.CadastrodeProduto
9   mnuCadVendedores.Enabled = Usuario.CadastrodeVendedor
10  mnuGrupoProd.Enabled = Usuario.GrupodeProdutos
11  mnuFornecedor.Enabled = Usuario.CadastroFornecedor
12  mnuTipoCliente.Enabled = Usuario.CadastrodeTipodeCliente
13  mnuCentroCursto.Enabled = Usuario.CadastrodeCentrodeCusto
14  mnuCondPag.Enabled = Usuario.CadastrodeCentrodeCusto
15  mnuBalcaoVendas.Enabled = Usuario.BalcaodeVendas
16  mnuUsuaruio.Enabled = Usuario.CadastrodeUsuario
17  mnuCheques.Enabled = Usuario.Cheque
18  MenuUnidade.Enabled = Usuario.Unidade
19  mnuListaProd.Enabled = Usuario.ListagemdeProdutos
20  mnRelaVendas.Enabled = Usuario.RelatoriodeVendas
21  mnuListaGemVendas.Enabled = Usuario.ListagemdeVendasporVendedor


End Sub

Private Sub MDIForm_Load()
1   On Error Resume Next


    'Menu.SubClassMenu Me
    'Menu.ImageList = MList
    'Menu.ItemIcon("mnuCadClinete") = 0
    'Menu.ItemIcon("mnuCadProduo") = 1
    'Menu.ItemIcon("mnuCadVendedores") = 2
    'Menu.ItemIcon("mnuGrupoProd") = 3
    'Menu.ItemIcon("mnuFornecedor") = 4
    'Menu.ItemIcon("mnuTipoCliente") = 5
    'Menu.ItemIcon("mnuCentroCursto") = 6
    'Menu.ItemIcon("mnuCondPag") = 7
    'Menu.ItemIcon("mnuBalcaoVendas") = 8
    'Menu.ItemIcon("mnuUsuaruio") = 9
2   If Aminacao = True Then
3       Randomize
4       FrmTela.Top = Rnd * 9000
5       Randomize
6       FrmTela.Left = Rnd * 9000
7       FrmTela.Show
8       XX = 0
9       YY = 0
10      MudarDesenho = False
11      Timer.Interval = 1
12      Timer.Enabled = True
13  Else

14      Centra FrmTela
15      Timer.Interval = 0
16      Timer.Enabled = False
17  End If
    'SQL = "Alter table Vendas add data Datetime" ' datetime default date()"
    'BancoDeDados.Execute SQL


    ''SQL = "Alter table ItenVenda add data Datetime" ' datetime default date()"
    'BancoDeDados.Execute SQL

    'SQL = "Create Table Empresa (Id counter,Nome Text(50) null ,"
    'SQL = SQL & "Doc Text(20)  null,Endereco Text(100) null ,"
    'SQL = SQL & "Bairro Text(50) null,Cidade Text(50) null,"
    'SQL = SQL & "Estado Text(50) null, Cep Text(50) null,"
    'SQL = SQL & "Telefone Text(20) null)"
    'BancoDeDados.Execute SQL

18  Sql = "Create Table Cheques (Id Counter, Data Datetime null,Cliente Text(100),"
19  Sql = Sql & " Numero long null, Banco Text(50) null, Agencia Long null, Conta Long null,"
20  Sql = Sql & " Vencimento datetime null,Valor double null, Obs memo)"

    '21    BancoDeDados.Execute SQL

21  MontaEmpresa
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
1   If MsgBox("Deseja sair do Sistema?", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal, App.Title) = vbYes Then
2       BancoDeDados.Close
3       Cancel = 0
4       End
5   Else
6       Cancel = 1
7   End If

End Sub

Private Sub MenuUnidade_Click()
1   FrmCadUnidade.Show 1
End Sub

Private Sub mnRelaVendas_Click()
1   FrmRelCaixa.Show 1
End Sub

Private Sub mnuBalcaoVendas_Click()
'Me.WindowState = 2
'mnuCadastro.Visible = False
'mnuMov.Visible = False
'mnuRelatorios.Visible = False
'mnuDiversos.Visible = False
'Timer.Enabled = False
1   FrmBalcao.Show 1

End Sub

Private Sub mnuCadClinete_Click()
1   FrmCadCLiente.Show
End Sub

Private Sub mnuCadProduo_Click()
1   FrmCadProd.Show
End Sub

Private Sub mnuCadVendedores_Click()
1   FrmVendedor.Show
End Sub

Private Sub mnuCentroCursto_Click()
1   FrmCentroCusto.Show
End Sub

Private Sub mnuCheques_Click()
1   FrmCheques.Show 1
End Sub

Private Sub mnuCondPag_Click()
1   FrmCondPag.Show
End Sub

Private Sub mnuDadosDaEmpresa_Click()
1   FrmCadEmpresa.Show 1
End Sub

Private Sub mnuFornecedor_Click()
1   FrmFornecedor.Show
End Sub

Private Sub mnuGrupoProd_Click()
1   FrmGrupoProd.Show
End Sub

Private Sub mnuListaGemVendas_Click()
1   FrmRelVendedor.Show
End Sub

Private Sub mnuListaProd_Click()
1   FrmListagemProd.Show 1
End Sub

Private Sub mnuSobre_Click()
1   MsgBox "SIAG " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & "Neo SoftWare - Telefone:31 88065049" & Chr(13) & "suporte: suporte@neobh.com.br" & Chr(13) & "Reclamações ou sugestões : sac@neobh.com.br" & Chr(13) & "Vendas : vendas@neobh.com.br" & Chr(13) & "Site : www.neobh.com.br", vbInformation, "Sobre"
End Sub

Private Sub mnuTipoCliente_Click()
1   FrmTipoClie.Show
End Sub

Private Sub mnuUsuaruio_Click()
1   FrmCadUsuario.Show 1
End Sub

Private Sub Timer_Timer()
1   If MudarDesenho = False Then
2       If PassoX = True Then
3           FrmTela.Left = FrmTela.Left - 50
4           If FrmTela.Left <= 0 Then
5               PassoX = False
6           End If
7       Else
8           FrmTela.Left = FrmTela.Left + 50
9           If FrmTela.Left >= FrmPrincipal.ScaleWidth - FrmTela.Width Then
10              PassoX = True
11          End If
12      End If

13      If PassoY = True Then
14          FrmTela.Top = FrmTela.Top - 50
15          If FrmTela.Top <= 0 Then
16              PassoY = False
17          End If
18      Else
19          FrmTela.Top = FrmTela.Top + 50
20          If FrmTela.Top >= FrmPrincipal.ScaleHeight - FrmTela.Height Then
21              PassoY = True
22          End If
23      End If
24  Else
25      If PassoX = True Then
26          FrmTela.Left = FrmTela.Left + 50
27          If FrmTela.Left <= 0 Then
28              PassoX = False
29          End If
30      Else
31          FrmTela.Left = FrmTela.Left - 50
32          If FrmTela.Left >= FrmPrincipal.ScaleWidth - FrmTela.Width Then
33              PassoX = True
34          End If
35      End If

36      If PassoY = True Then
37          FrmTela.Top = FrmTela.Top + 50
38          If FrmTela.Top <= 0 Then
39              PassoY = False
40          End If
41      Else
42          FrmTela.Top = FrmTela.Top - 50
43          If FrmTela.Top >= FrmPrincipal.ScaleHeight - FrmTela.Height Then
44              PassoY = True
45          End If
46      End If
47  End If

End Sub

Private Sub Timer2_Timer()
1   MudarDesenho = Not MudarDesenho
End Sub
