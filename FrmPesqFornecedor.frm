VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesqFornecedor 
   Caption         =   "Pesquisa de Fornecedores"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   Icon            =   "FrmPesqFornecedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPesq 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   7935
   End
   Begin VB.Data Banco 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Emmanuel\Fontes\Controle de Estoque Simples\Controle.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Fornecedor"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmPesqFornecedor.frx":058A
      Height          =   3825
      Left            =   0
      OleObjectBlob   =   "FrmPesqFornecedor.frx":059E
      TabIndex        =   1
      Top             =   630
      Width           =   7965
   End
   Begin VB.Label Label1 
      Caption         =   "Noome / Razao do Fornecedor"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   90
      Width           =   2325
   End
End
Attribute VB_Name = "FrmPesqFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_DblClick()
1   Abrir
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Abrir
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   Banco.DatabaseName = LocalBanco
2   Banco.Connect = ";pwd=" & SenhaSistema
3   ShowPesqFornecedor = 0
End Sub

Private Sub txtPesq_Change()
1   Banco.RecordSource = "Select * From Fornecedor Where Razao Like '" & txtPesq.Text & "*' Order By Razao"
2   Banco.Refresh
End Sub

Private Sub Abrir()
1   On Error Resume Next
2   If Banco.Recordset.EOF = False Then
3       ShowFornecedor = Banco.Recordset!Codigo
4       Unload Me
5   End If
End Sub

