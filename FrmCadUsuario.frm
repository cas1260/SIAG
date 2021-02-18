VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCadUsuario 
   Caption         =   "Cadastro de Usuario"
   ClientHeight    =   6195
   ClientLeft      =   1635
   ClientTop       =   1860
   ClientWidth     =   8415
   Icon            =   "FrmCadUsuario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4725
      Left            =   3990
      TabIndex        =   11
      Top             =   1350
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   8334
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      FormatString    =   "Login       | Nome do Usuario                                   "
   End
   Begin VB.Frame fd 
      Caption         =   "Permissões"
      Height          =   4785
      Left            =   60
      TabIndex        =   9
      Top             =   1290
      Width           =   3855
      Begin MSComctlLib.ImageList ImageList222 
         Left            =   1710
         Top             =   3270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCadUsuario.frx":0442
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView La 
         Height          =   4515
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   7964
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ImageList222"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox Txtnome 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Top             =   930
      Width           =   4215
   End
   Begin VB.TextBox TxtConfirmacao 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2730
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   930
      Width           =   1305
   End
   Begin VB.TextBox TxtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   930
      Width           =   1305
   End
   Begin VB.TextBox TxtLogin 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   930
      Width           =   1305
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":09AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadUsuario.frx":0F0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   953
      ButtonWidth     =   1005
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "C"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "D"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "G"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Label lblId 
      Height          =   195
      Left            =   2640
      TabIndex        =   12
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nome do usuario"
      Height          =   255
      Left            =   4110
      TabIndex        =   8
      Top             =   690
      Width           =   1845
   End
   Begin VB.Label Label2 
      Caption         =   "Confirmação"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   690
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Senha"
      Height          =   255
      Left            =   1410
      TabIndex        =   6
      Top             =   690
      Width           =   885
   End
   Begin VB.Label lblLogin 
      Caption         =   "Login"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   885
   End
End
Attribute VB_Name = "FrmCadUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   Select Case Button.index
    Case 1
2       Novo
3   Case 2
4       Salvar
5   Case 3
6       Excluir
7   Case 5
8       Unload Me
9   End Select
End Sub

Private Sub Form_Activate()
1   Novo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   MontaGrid
End Sub

Private Sub Salvar()
1   On Error GoTo Trata_Erro
    Dim Nivel As String, X                                  As Long

2   If UCase(TxtSenha.Text) <> UCase(TxtConfirmacao.Text) Then
3       MsgBox "A Senha não conferir com sua confirmação!", vbCritical, App.Title
4       TxtSenha.SetFocus
5       Exit Sub
6   End If

7   If Trim(TxtLogin.Text) = "" Then
8       MsgBox "Login invalido!", vbCritical, App.Title
9       TxtLogin.SetFocus
10      Exit Sub
11  End If

12  If Trim(TxtNome.Text) = "" Then
13      MsgBox "Nome invalido!"
14      TxtNome.SetFocus
15      Exit Sub
16  End If

17  If Trim(TxtSenha.Text) = "" Then
18      MsgBox "Senha invalida!", vbCritical, App.Title
19      TxtSenha.SetFocus
20      Exit Sub
21  End If
    Dim Rs                                                  As Recordset

22  Sql = "Select * From Usuario Where Login ='" & TxtLogin.Text & "'"
23  Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)

24  If Rs.EOF = False Then
25      Rs.Edit
26  Else
27      Rs.AddNew
28  End If

29  Rs!Login = TxtLogin.Text
30  Rs!Senha = TxtSenha.Text
31  Rs!Nome = TxtNome.Text
32  Nivel = ""
33  For X = 1 To La.Nodes.Count
34      Nivel = Nivel & IIf(La.Nodes(X).Checked = True, "1", "0")
35  Next X
36  Rs!Nivel = Nivel
37  Rs.Update
38  MontaGrid
39  MsgBox "Usuario Cadastrado com sucesso!", vbInformation, App.Title
40  Novo


Trata_Erro:
41  E
End Sub
Private Sub Novo()
1   TxtConfirmacao.Text = ""
2   TxtSenha.Text = ""
3   TxtLogin.Text = ""
4   TxtNome.Text = ""
5   La.Nodes.Clear

6   With La.Nodes
7       .Add , , "cadusu", "Cadastro de Usuario", 1
8       .Add , , "cadcliente", "Cadastro de Cliente", 1
9       .Add , , "cadproduto", "Cadastro de Produto", 1
10      .Add , , "entrar", "Entrada no estoque", 1
11      .Add , , "cadVend", "Cadastro de Vendedor", 1
12      .Add , , "For", "Cadastro Fornecedor", 1
13      .Add , , "grupo", "Grupo de Produtos", 1
14      .Add , , "balcao", "Balção de Vendas", 1
15      .Add , , "Cheque", "Cheques", 1
16      .Add , , "uni", "Cadastro de Unidade", 1
17      .Add , , "Tipo", "Cadastro de Tipo de Cliente", 1
18      .Add , , "ListProd", "Listagem de Produtos", 1
19      .Add , , "ListRelVed", "Relatorio de Vendas", 1
20      .Add , , "listvend", "Listagem de Vendas por Cliente", 1
21  End With

    'La.Nodes(5).Visible = False
    'La.Nodes("CodPg").Visible = False

22  La.Refresh
23  TxtLogin.SetFocus
End Sub

'Private Sub Grid_Click()
'MostraReg
'End Sub

Private Sub grid_DblClick()
1   MostraReg
End Sub

Private Sub La_NodeClick(ByVal Node As MSComctlLib.Node)
1   Node.Checked = Not Node.Checked
End Sub

Private Sub Excluir()
1   On Error Resume Next
2   If Grid.TextMatrix(Grid.Row, 0) <> "" Then
3       If MsgBox("Confirma exclusão deste usuario?" & Chr(13) & Grid.TextMatrix(Grid.Row, 2), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) = vbYes Then
4           Sql = "Delete from Usuario Where Id =" & Grid.TextMatrix(Grid.Row, 0)
5           BancoDeDados.Execute Sql
6           MsgBox "Usuario excluido com sucesso!", vbInformation, "Ok"
7           MontaGrid
8           Novo
9       End If
10  End If
Trata_Erro:
11  E
End Sub
Private Sub MostraReg()

1   If Grid.TextMatrix(1, 0) = "" Then Exit Sub

    Dim Nivel                                               As String
2   X = Grid.Row
3   lblId.Caption = Grid.TextMatrix(X, 0)
4   Nivel = Grid.TextMatrix(X, 1)
5   TxtLogin.Text = Grid.TextMatrix(X, 2)
6   TxtNome.Text = Grid.TextMatrix(X, 3)
7   TxtSenha.Text = Grid.TextMatrix(X, 4)

8   For X = 1 To La.Nodes.Count
9       La.Nodes(X).Checked = IIf(Mid(Nivel, X, 1) = "1", True, False)
10  Next X

End Sub
Private Sub MontaGrid()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   Grid.Rows = 2
3   Grid.Clear
4   Grid.FormatString = "Id|Nivel|Login       | Nome do Usuario                                   |Senha"
5   Grid.ColWidth(0) = 0
6   Grid.ColWidth(1) = 0
7   Grid.ColWidth(4) = 0

8   Sql = "Select * From Usuario"
9   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
10  If Rs.EOF = True Then Exit Sub
11  X = 1
12  Rs.MoveLast
13  Rs.MoveFirst
14  Grid.Rows = Rs.RecordCount + 1
15  Do While Not Rs.EOF
16      Grid.TextMatrix(X, 0) = Rs!Id
17      Grid.TextMatrix(X, 1) = Rs!Nivel
18      Grid.TextMatrix(X, 2) = Rs!Login
19      Grid.TextMatrix(X, 3) = Rs!Nome
20      Grid.TextMatrix(X, 4) = Rs!Senha
21      X = X + 1
22      Rs.MoveNext
23  Loop

24  Rs.Close
Trata_Erro:
25  E
End Sub

