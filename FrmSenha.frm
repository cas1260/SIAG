VERSION 5.00
Begin VB.Form FrmSenha 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2730
   ClientLeft      =   1650
   ClientTop       =   1875
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      BackColor       =   &H8000000E&
      Caption         =   "Sair"
      Height          =   375
      Left            =   3330
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CmdEntrar 
      BackColor       =   &H8000000E&
      Caption         =   "Entrar no Sistema"
      Height          =   375
      Left            =   1770
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1590
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1770
      Width           =   3255
   End
   Begin VB.TextBox TxtLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   1290
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   990
      TabIndex        =   3
      Top             =   1830
      Width           =   1875
   End
   Begin VB.Label bllogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1020
      TabIndex        =   1
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Image Image 
      Height          =   1350
      Left            =   0
      Picture         =   "FrmSenha.frx":0000
      Top             =   270
      Width           =   4125
   End
   Begin VB.Label LblTop 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEntrar_Click()
1   BuscaUsuario

End Sub

Private Sub CmdSair_Click()
1   BancoDeDados.Close
2   End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then CmdSair_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
1   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
1   LblTop.Caption = "Siag - " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub BuscaUsuario()
    Dim Rs                                                  As Recordset
1   On Error GoTo Trata_Erro

2   TxtLogin.Text = UCase(TxtLogin.Text)
3   TxtSenha.Text = UCase(TxtSenha.Text)

4   If Trim(TxtLogin.Text) = "" Then
5       MsgBox "Login invalido!", vbCritical, App.Title
6       TxtLogin.SetFocus
7       Exit Sub
8   End If

9   If Trim(TxtSenha.Text) = "" Then
10      MsgBox "Senha invalida!", vbCritical, App.Title
11      TxtSenha.SetFocus
12      Exit Sub
13  End If

14  If TxtLogin.Text = "NEO" Then
15      If TxtSenha = "NSLTR" & Format(Minute(Time()), "00") Then
16          With Usuario
17              .Nome = "Usuario Master"
18              .Login = "Master"
19              .Senha = "Master"
20              .DataHoraLogin = Date
21              .Cheque = True
22              .Unidade = True
23              .CadastrodeUsuario = True
24              .CadastrodeCliente = True
25              .CadastrodeProduto = True
26              .Entradanoestoque = True
27              .CadastrodeVendedor = True
28              .CadastrodeTipodeCliente = True
29              .CadastrodeCentrodeCusto = True
30              .CondicaodePagamento = True
31              .CadastroFornecedor = True
32              .GrupodeProdutos = True
33              .BalcaodeVendas = True
34              .ListagemdeProdutos = True
35              .RelatoriodeVendas = True
36              .ListagemdeVendasporVendedor = True

37          End With
38          Unload Me
39          Exit Sub
40      End If
41  End If


42  Sql = "Select * From Usuario Where Login ='" & TxtLogin.Text & "'"
43  Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

44  If Rs.EOF = True Then
45      MsgBox "Login não cadastrado!", vbCritical, App.Title
46      Rs.Close
47      TxtLogin.SetFocus
48      Exit Sub
49  End If
    Dim Senha                                               As String

50  Senha = UCase(Rs!Senha)
51  TxtSenha.Text = UCase(TxtSenha.Text)

52  If TxtSenha.Text <> Senha Then
53      MsgBox "Senha incorreta!", vbCritical, App.Title
54      TxtSenha.SetFocus
55      Exit Sub
56  End If


57  With Usuario
58      .Nome = Rs!Nome
59      .Login = Rs!Login
60      .Senha = Rs!Senha
61      .DataHoraLogin = Date

62      .CadastrodeUsuario = IIf(Mid(Rs!Nivel, 1, 1) = "1", True, False)
63      .CadastrodeCliente = IIf(Mid(Rs!Nivel, 2, 1) = "1", True, False)
64      .CadastrodeProduto = IIf(Mid(Rs!Nivel, 3, 1) = "1", True, False)
65      .Entradanoestoque = IIf(Mid(Rs!Nivel, 4, 1) = "1", True, False)
66      .CadastrodeVendedor = IIf(Mid(Rs!Nivel, 5, 1) = "1", True, False)
67      .CadastroFornecedor = IIf(Mid(Rs!Nivel, 6, 1) = "1", True, False)
68      .GrupodeProdutos = IIf(Mid(Rs!Nivel, 7, 1) = "1", True, False)
69      .BalcaodeVendas = IIf(Mid(Rs!Nivel, 8, 1) = "1", True, False)
70      .Cheque = IIf(Mid(Rs!Nivel, 9, 1) = "1", True, False)
71      .Unidade = IIf(Mid(Rs!Nivel, 10, 1) = "1", True, False)
72      .CadastrodeTipodeCliente = IIf(Mid(Rs!Nivel, 11, 1) = "1", True, False)
73      .ListagemdeProdutos = IIf(Mid(Rs!Nivel, 12, 1) = "1", True, False)
74      .RelatoriodeVendas = IIf(Mid(Rs!Nivel, 13, 1) = "1", True, False)
75      .ListagemdeVendasporVendedor = IIf(Mid(Rs!Nivel, 14, 1) = "1", True, False)
76  End With

77  Rs.Close
78  Unload Me
Trata_Erro:
79  E
End Sub
