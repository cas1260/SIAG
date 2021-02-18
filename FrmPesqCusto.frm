VERSION 5.00
Begin VB.Form FrmPesqCusto 
   Caption         =   "Pesquisa de Centro de Custo"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPesqCusto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPesq 
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   90
      Width           =   4905
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   435
   End
End
Attribute VB_Name = "FrmPesqCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_DblClick()
1   Busca
End Sub
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Busca
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
2   Banco.DatabaseName = CaminhoBanco
End Sub

Private Sub txtPesq_Change()
1   On Error GoTo Trata_Erro
2   Banco.Recordset.FindFirst "Descricao like '" & TxtPesq.Text & "*'"
Trata_Erro:
3   E
End Sub

Private Sub Busca()
1   On Error GoTo Trata_Erro
2   If Banco.Recordset.EOF = False Then
3       ShowCentro = Banco.Recordset!Codigo
4       Unload Me
5   End If
Trata_Erro:
6   E
End Sub
