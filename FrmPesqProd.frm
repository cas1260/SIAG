VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPesqProd 
   Caption         =   "Pesquisa de Produtos"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   Icon            =   "FrmPesqProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3825
      Left            =   30
      TabIndex        =   2
      Top             =   690
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   6747
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   $"FrmPesqProd.frx":0442
   End
   Begin VB.Frame F 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Nome do Produto"
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   8355
      Begin VB.TextBox TxtPesq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   8235
      End
   End
End
Attribute VB_Name = "FrmPesqProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Order                                                   As String
Dim IndexGeral                                              As String

Private Sub DBGrid1_DblClick()
1   Abrir 13
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
1   Abrir KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 27 Then
2       ShowProduto = ""
3       Unload Me
4   End If
End Sub

Private Sub Form_Load()
1   IndexGeral = 0
2   txtPesq_Change
End Sub

Private Sub grid_DblClick()
1   ShowProduto = ""
2   If Grid.TextMatrix(1, 0) <> "" Then
3       ShowProduto = Grid.TextMatrix(Grid.Row, 0)
4   End If
5   Unload Me
End Sub

Private Sub txtPesq_Change()
1   Sql = "Descricao Like '" & TxtPesq.Text & "*'"
2   Sql = "Select * From Produtos Where " & Sql & " Order By Descricao"
    Dim Rs                                                  As Recordset
3   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
4   Grid.Rows = 2
5   Grid.Clear
6   Grid.FormatString = "Codigo        |Descrição                                                                                                           |Qtd no Estoque"
7   If Rs.EOF = False Then
8       Rs.MoveLast
9       Rs.MoveFirst
10      Grid.Rows = Rs.RecordCount + 1
11      X = 1
12      Do While Not Rs.EOF
13          Grid.TextMatrix(X, 0) = Rs!Codigo
14          Grid.TextMatrix(X, 1) = Rs!Descricao
15          Grid.TextMatrix(X, 2) = Rs!Atual
16          X = X + 1
17          Rs.MoveNext
18      Loop
19  End If
20  Rs.Close
End Sub

Private Sub Abrir(Key As Integer)
1   On Error Resume Next
2   If Key = 13 Then
3       If Grid.TextMatrix(Grid.Row, 0) <> "" Then
4           ShowProduto = Grid.TextMatrix(Grid.Row, 0)
5           Unload Me
6       End If
7   End If
End Sub

Private Sub TxtPesq_KeyDown(KeyCode As Integer, Shift As Integer)
1   Abrir KeyCode
2   If KeyCode = vbKeyDown Then
3       If Banco.Recordset.EOF = False Then
4           DBGrid1.SetFocus
5           Banco.Recordset.MoveNext
6       End If
7   End If
End Sub
