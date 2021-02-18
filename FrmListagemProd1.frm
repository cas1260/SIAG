VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmListagemProd 
   Caption         =   "Listagem de Produto"
   ClientHeight    =   5730
   ClientLeft      =   1620
   ClientTop       =   1830
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8370
   Begin Project1.Imp rel 
      Left            =   900
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   873
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5085
      Left            =   150
      TabIndex        =   6
      Top             =   540
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Codigo          |Descrição                                                |Qtd.              |Valor Compra      |Valor Venda.     "
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   6690
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   5130
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtFinal 
      Height          =   285
      Left            =   3540
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtIni 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Final"
      Height          =   225
      Left            =   2550
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Labeldd 
      Caption         =   "Codigo Inicial"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "FrmListagemProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim RsProd                                              As Recordset

1   TxtIni.Text = IIf(Trim(TxtIni.Text) = "", "0", TxtIni.Text)
2   TxtFinal.Text = IIf(Trim(TxtFinal.Text) = "", "999999", TxtFinal.Text)

3   Comando = "Select * From Produtos Where val(Codigo ) >=" & TxtIni.Text & " And Val(Codigo) <=" & TxtFinal.Text & " Order By Descricao"

4   Grid.Clear
5   Grid.FormatString = "Codigo          |Descrição                                                |Qtd.              |Valor Compra      |Valor Venda.     "
6   Grid.Rows = 2
7   Set RsProd = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)
8   If RsProd.EOF = False Then
9       X = 1
10      RsProd.MoveLast
11      RsProd.MoveFirst
12      Grid.Rows = RsProd.RecordCount + 1
13      Do While Not RsProd.EOF
14          Grid.TextMatrix(X, 0) = RsProd!Codigo
15          Grid.TextMatrix(X, 1) = RsProd!Descricao
16          Grid.TextMatrix(X, 2) = Format(RsProd!Atual, "###,##0.000")
17          Grid.TextMatrix(X, 3) = Format(RsProd!Compra, "###,###,##0.00")
18          Grid.TextMatrix(X, 4) = Format(RsProd!Venda, "###,###,##0.00")
19          RsProd.MoveNext
20          X = X + 1
21      Loop
22  End If
23  RsProd.Close

End Sub

Private Sub Command2_Click()
1   Set Rel.GridImpir = Grid
2   Rel.Titulo = "Listagem de Produtos"
3   Rel.SubTitulo = ""
4   Rel.Rodape = "Totais de Produto :" & Grid.Rows - 1
5   Rel.NovoRelatorio
6   Rel.DefineCampos "Codigo", 0, 20, True, True
7   Rel.DefineCampos "Descrição", 1, 70, True, False
8   Rel.DefineCampos "Qtd", 2, 15, True, True
9   Rel.DefineCampos "Valor Compras", 3, 25, True, True
10  Rel.DefineCampos "Valor Venda", 4, 25, True, True
11  Rel.ImprimirRelatorios


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

