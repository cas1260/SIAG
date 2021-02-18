VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesqCliente 
   Caption         =   "Pesquisa de Cliente"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "FrmPesqCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   30
      ScaleHeight     =   3315
      ScaleWidth      =   6225
      TabIndex        =   6
      Top             =   750
      Width           =   6285
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmPesqCliente.frx":0442
         Height          =   3315
         Left            =   0
         OleObjectBlob   =   "FrmPesqCliente.frx":0456
         TabIndex        =   3
         Top             =   0
         Width           =   6225
      End
   End
   Begin VB.Data Banco 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VbRc\0002\Rc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cliente"
      Top             =   1260
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Pesquisa"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   6315
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2700
         TabIndex        =   7
         Top             =   30
         Width           =   30
      End
      Begin VB.TextBox TxtPesq 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   270
         Width           =   2625
      End
      Begin VB.OptionButton O3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Cpf/Cnpj"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1680
         TabIndex        =   1
         Top             =   345
         Width           =   1005
      End
      Begin VB.OptionButton O1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Razao/Nome"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   270
         TabIndex        =   0
         Top             =   345
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.Label L 
         Caption         =   "Razao"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   2790
         TabIndex        =   5
         Top             =   330
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmPesqCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_DblClick()
1   BuscaAbrir
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       BuscaAbrir
3   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       SendKeys "{TAB}"
3   ElseIf KeyCode = 27 Then
4       Unload Me
5   End If
End Sub

Private Sub Form_Load()
1   Banco.Connect = ";pwd=" & SenhaSistema
2   Banco.DatabaseName = LocalBanco
3   O1_Click
End Sub

Private Sub O1_Click()
1   L.Caption = "Razao"
2   Banco.RecordSource = "Select * from Cliente Order By Razao"
3   Banco.Refresh
End Sub

Private Sub O3_Click()
1   L.Caption = "Cnpj"
2   Banco.RecordSource = "Select * from Cliente Order By cnpj"
3   Banco.Refresh
End Sub

Private Sub txtPesq_Change()
1   If O1.Value = True Then
2       Banco.Recordset.FindFirst "Razao like '" & txtPesq.Text & "*'"
3   Else
4       Banco.Recordset.FindFirst "Cnpj like '" & txtPesq.Text & "*'"
5   End If
End Sub

Private Sub BuscaAbrir()
1   On Error GoTo Trata_Erro
2   If Banco.Recordset.EOF = False Then
3       ShowCliente = Banco.Recordset!Codigo
4       Unload Me
5   End If
Trata_Erro:
6   E
End Sub
