VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmLocalTipoCliente 
   Caption         =   "Localizar Tipo de Cliente"
   ClientHeight    =   3720
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   6375
   ClipControls    =   0   'False
   Icon            =   "frmLocalTipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   270
      Width           =   6315
   End
   Begin VB.Data Banco 
      Connect         =   ";pwd=1906bili"
      DatabaseName    =   "D:\Cleber\Fontes\Siag\Siag97.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   750
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vendedor"
      Top             =   3750
      Width           =   5205
   End
   Begin MSDBGrid.DBGrid grid 
      Bindings        =   "frmLocalTipo.frx":014A
      Height          =   3075
      Left            =   30
      OleObjectBlob   =   "frmLocalTipo.frx":015E
      TabIndex        =   2
      Top             =   570
      Width           =   6255
   End
   Begin VB.Label Label 
      Caption         =   "Nome do Vendedor"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   1605
   End
End
Attribute VB_Name = "FrmLocalTipoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
1   Banco.DatabaseName = LocalBanco
2   Banco.Connect = ";pwd=" & SenhaSistema
3   Banco.RecordSource = "Select * From TipoCli Order By Descricao"
4   Banco.Refresh
End Sub

Private Sub grid_DblClick()
1   If Banco.Recordset.EOF = False Then
2       ShowTipoCliente = Banco.Recordset!Codigo
3       Unload Me
4   End If
End Sub

Private Sub TxtNome_Change()
1   Banco.Recordset.FindFirst "Descricao like '" & TxtNome.Text & "*'"
End Sub
