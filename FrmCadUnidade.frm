VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCadUnidade 
   Caption         =   "Cadastro de Unidades"
   ClientHeight    =   6060
   ClientLeft      =   1665
   ClientTop       =   1590
   ClientWidth     =   5520
   Icon            =   "FrmCadUnidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Banco 
      Connect         =   ";pwd=1906bili"
      DatabaseName    =   "D:\Cleber\Fontes\Siag\Siag97.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblUnidade"
      Top             =   6060
      Visible         =   0   'False
      Width           =   2265
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FrmCadUnidade.frx":0442
      Height          =   5925
      Left            =   60
      OleObjectBlob   =   "FrmCadUnidade.frx":0456
      TabIndex        =   0
      Top             =   60
      Width           =   5415
   End
End
Attribute VB_Name = "FrmCadUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
1   Banco.DatabaseName = CaminhoBanco
End Sub
