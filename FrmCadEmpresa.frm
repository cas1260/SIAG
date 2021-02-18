VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCadEmpresa 
   Caption         =   "Cadastro da Empresa"
   ClientHeight    =   2445
   ClientLeft      =   1620
   ClientTop       =   1905
   ClientWidth     =   6480
   Icon            =   "FrmCadEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   9
      Top             =   1890
      Width           =   1155
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Salvar"
      Height          =   345
      Left            =   4140
      TabIndex        =   8
      Top             =   1890
      Width           =   1155
   End
   Begin VB.TextBox TxtCep 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   30
      TabIndex        =   6
      Top             =   1950
      Width           =   2505
   End
   Begin VB.TextBox TxtEndereco 
      Height          =   285
      Left            =   2100
      TabIndex        =   2
      Top             =   930
      Width           =   4305
   End
   Begin VB.TextBox TxtBairro 
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   1410
      Width           =   4605
   End
   Begin VB.TextBox TxtCidade 
      Height          =   285
      Left            =   4650
      TabIndex        =   4
      Top             =   1410
      Width           =   1155
   End
   Begin VB.TextBox TxtEstado 
      Height          =   285
      Left            =   5820
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1410
      Width           =   585
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   6375
   End
   Begin MSMask.MaskEdBox MskDoc 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   930
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTelefone 
      Height          =   285
      Left            =   2580
      TabIndex        =   7
      Top             =   1950
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(##) ####-####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label11 
      Caption         =   "Endereço"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2130
      TabIndex        =   17
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Bairro"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   30
      TabIndex        =   16
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Cidade"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4650
      TabIndex        =   15
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Estado"
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   5820
      TabIndex        =   14
      Top             =   1230
      Width           =   585
   End
   Begin VB.Label Label15 
      Caption         =   "Cep"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   30
      TabIndex        =   13
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Telefone"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2580
      TabIndex        =   12
      Top             =   1740
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Cpf/Cnpj:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   30
      TabIndex        =   11
      Top             =   690
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Nome da Empresa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCadEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
1   Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Rs                                                  As Recordset

1   Sql = "Select * From Empresa"
2   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
3   If Rs.EOF = False Then
4       Rs.Edit
5   Else
6       Rs.AddNew
7   End If
8   Rs!Nome = TxtNome.Text
9   Rs!Doc = MskDoc.Text
10  Rs!Endereco = TxtEndereco.Text
11  Rs!Bairro = TxtBairro.Text
12  Rs!Cidade = TxtCidade.Text
13  Rs!Estado = TxtEstado.Text
14  Rs!Cep = TxtCep.Text
15  Rs!Telefone = MskTelefone.Text
16  Rs.Update
17  MsgBox "Dados da Empresa salvo com sucesso!", vbInformation, App.Title
18  MontaEmpresa
19  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   TxtNome.Text = Empresa.Nome
2   MskDoc.Text = Empresa.CNPJ
3   TxtEndereco.Text = Empresa.Endereco
4   TxtBairro.Text = Empresa.Bairro
5   TxtCidade.Text = Empresa.Cidade
6   TxtEstado.Text = Empresa.Estado
7   TxtCep.Text = Empresa.Cep
8   MskTelefone.Text = Empresa.Telefone
End Sub
