VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRelCaixa 
   Caption         =   "Relatorio de Caixa / Fecha Caixa"
   ClientHeight    =   5355
   ClientLeft      =   1635
   ClientTop       =   1470
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Imp Rel 
      Left            =   1590
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   873
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   5100
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   6300
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   $"FrmRelCaixa.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   4740
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MskInicial 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskFinal 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "Desconto :"
      Height          =   195
      Left            =   2580
      TabIndex        =   11
      Top             =   5130
      Width           =   855
   End
   Begin VB.Label LblDesconto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3450
      TabIndex        =   10
      Top             =   4920
      Width           =   1965
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   6030
      TabIndex        =   8
      Top             =   4920
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Total :"
      Height          =   195
      Left            =   5490
      TabIndex        =   6
      Top             =   5130
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   2400
      TabIndex        =   3
      Top             =   330
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRelCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'MsgBox "Erro number 404 not found is obj.", vbCritical, "Err"
    Dim RsProd As Recordset, RsCompra                       As Recordset
    Dim Total As Double, TotalProd As Double, TotalQtd      As Double
    Dim Desconto                                            As Double

1   Sql = "Select * From Produtos Order By Descricao , Codigo"
2   Set RsProd = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
3   MskFinal.Text = Valida(MskFinal)
4   MskInicial.Text = Valida(MskInicial)
5   Barra.Min = 0
6   Barra.Value = 0
7   Grid.Clear
8   Grid.Rows = 2
9   Grid.FormatString = "Codigo    |Produto                                                                           |Qtd                  |Total                       "
10  LblDesconto.Caption = Format(Desconto, "###,###,##0.00")
11  lblTotal.Caption = Format(Total, "###,###,##0.00")

12  If RsProd.EOF = False Then
13      RsProd.MoveLast
14      RsProd.MoveFirst
15      Barra.Max = RsProd.RecordCount
16      Grid.Rows = RsProd.RecordCount + 1
17      X = 1
18      Total = 0
19      TotalProd = 0
20      TotalQtd = 0
21      Do While Not RsProd.EOF
22          Comando = "Select * From ItenVenda Where Codigo =" & RsProd!Codigo & " And "
23          Comando = Comando & "Data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# And Data <=#" & Format(MskFinal.Text, "mm/DD/yyyy") & "# Order By Data Desc"
24          Set RsCompra = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)
25          Grid.TextMatrix(X, 0) = RsProd!Codigo
26          Grid.TextMatrix(X, 1) = RsProd!Descricao
27          TotalProd = 0
28          TotalQtd = 0

29          If RsCompra.EOF = False Then
30              TotalProd = 0
31              TotalQtd = 0
32              Do While Not RsCompra.EOF
33                  TotalProd = TotalProd + RsCompra!Total
34                  TotalQtd = TotalQtd + RsCompra!Qtd
35                  RsCompra.MoveNext
36              Loop
37          End If
38          Grid.TextMatrix(X, 2) = Format(TotalQtd, "###,###,##0.000")
39          Grid.TextMatrix(X, 3) = Format(TotalProd, "###,###,##0.00")
40          Total = Total + TotalProd
41          RsProd.MoveNext
42          Barra.Value = Barra.Value + 1
43          X = X + 1
44          DoEvents
45          DoEvents
46      Loop
47      RsProd.Close
48      Barra.Value = 0
49      Barra.Min = 0
50      Barra.Max = Grid.Rows
51      For X = 1 To Grid.Rows
52          If X = Grid.Rows Then Exit For
53          DoEvents
54          If Val(Grid.TextMatrix(X, 2)) = 0 And Val(Grid.TextMatrix(X, 3)) = 0 Then
55              If X = 1 And Grid.Rows = 2 Then
56                  Grid.Clear
57                  Grid.Rows = 2
58                  Grid.FormatString = "Codigo    |Produto                                                                           |Qtd                  |Total                       "
59                  Exit For
60              Else
61                  Grid.RemoveItem X
62              End If
63              X = X - 1
64          End If
65          Barra.Value = Barra.Value + 1
66          DoEvents
67      Next X



68      Comando = "Select * From Vendas Where Data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# And Data <=#" & Format(MskFinal.Text, "mm/DD/yyyy") & "# Order By Data Desc"
69      Set RsCompra = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
70      Desconto = 0
71      If RsCompra.EOF = False Then
72          Do While Not RsCompra.EOF
73              Desconto = Desconto + RsCompra!Desconto
74              RsCompra.MoveNext
75          Loop
76      End If
77      LblDesconto.Caption = Format(Desconto, "###,###,##0.00")
78      lblTotal.Caption = Format(Total, "###,###,##0.00")
79  End If

80  Barra.Value = 0

End Sub

Private Sub Command2_Click()
1   Set Rel.GridImpir = Grid
2   Rel.Titulo = "Relatorio de Caixa"
3   Rel.SubTitulo = "Periodo :" & MskInicial.Text & "  a  " & MskFinal.Text
4   Rel.Rodape = "Descontos R$ " & LblDesconto.Caption & " Valor Brunto R$ " & LblDesconto.Caption & " Valor Liquido " & Format(CCur(lblTotal.Caption) - CCur(LblDesconto.Caption), "###.###.##0,00") & " "
5   Rel.NovoRelatorio
6   Rel.DefineCampos "Codigo", 0, 20, True, True
7   Rel.DefineCampos "Produto", 1, 50, True, True
8   Rel.DefineCampos "Qtd", 2, 20, True, True
9   Rel.DefineCampos "Total", 3, 30, True, True
10  Rel.ImprimirRelatorios
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub MskFinal_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskFinal.Text = Valida(MskFinal)
3   End If
End Sub

Private Sub MskInicial_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskInicial.Text = Valida(MskInicial)
3   End If
End Sub
