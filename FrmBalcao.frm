VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalcao 
   BackColor       =   &H00808080&
   ClientHeight    =   8775
   ClientLeft      =   1635
   ClientTop       =   1950
   ClientWidth     =   11370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Timer Ativar 
      Left            =   3810
      Top             =   5040
   End
   Begin VB.Frame FPreco 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1050
      TabIndex        =   26
      Top             =   3810
      Visible         =   0   'False
      Width           =   8715
      Begin VB.Label lblPreco 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Width           =   8595
      End
   End
   Begin VB.Frame FCTop 
      Caption         =   "Dados Principais"
      Height          =   1185
      Left            =   -30
      TabIndex        =   20
      Top             =   1110
      Width           =   12015
      Begin VB.TextBox TxtObs 
         Height          =   495
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   600
         Width           =   3225
      End
      Begin VB.TextBox TxtMotivo 
         Height          =   495
         Left            =   8940
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   600
         Width           =   2985
      End
      Begin VB.TextBox TxtDesconto 
         Height          =   285
         Left            =   10260
         TabIndex        =   3
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox TxtCodVend 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox TxtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1110
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
      Begin VB.ComboBox CboCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2325
      End
      Begin VB.ComboBox CboVendedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   630
         Width           =   2325
      End
      Begin VB.ComboBox CboCond 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmBalcao.frx":0000
         Left            =   5670
         List            =   "FrmBalcao.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs.:"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   4260
         TabIndex        =   28
         Top             =   690
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo do Desconto :"
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   8010
         TabIndex        =   25
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto :"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   9420
         TabIndex        =   24
         Top             =   360
         Width           =   795
      End
      Begin VB.Label LblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   510
         TabIndex        =   23
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedores :"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Condição de Pgt. :"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   4320
         TabIndex        =   21
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.ComboBox CboProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2745
      Left            =   4950
      Style           =   1  'Simple Combo
      TabIndex        =   19
      Top             =   810
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   1710
      TabIndex        =   17
      Top             =   3300
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Ftotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Total da Vendas"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      TabIndex        =   15
      Top             =   7890
      Width           =   3465
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   60
         TabIndex        =   16
         Top             =   270
         Width           =   3255
      End
   End
   Begin VB.Frame Tbar 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   960
      TabIndex        =   14
      Top             =   3870
      Width           =   7845
      Begin VB.CommandButton B 
         BackColor       =   &H00808080&
         Caption         =   "Nova Venda (F2)"
         Height          =   975
         Index           =   0
         Left            =   0
         Picture         =   "FrmBalcao.frx":0095
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton B 
         BackColor       =   &H00808080&
         Caption         =   "Incluir"
         Height          =   975
         Index           =   1
         Left            =   1560
         Picture         =   "FrmBalcao.frx":0D5F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton B 
         BackColor       =   &H00808080&
         Caption         =   "Excluir"
         Height          =   975
         Index           =   2
         Left            =   3120
         Picture         =   "FrmBalcao.frx":11EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton B 
         BackColor       =   &H00808080&
         Caption         =   "Finalizar (F3)"
         Height          =   975
         Index           =   3
         Left            =   4680
         Picture         =   "FrmBalcao.frx":1EB4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton B 
         BackColor       =   &H00808080&
         Caption         =   "Sair (F1)"
         Height          =   975
         Index           =   4
         Left            =   6240
         Picture         =   "FrmBalcao.frx":277E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList Ilist 
      Left            =   5760
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBalcao.frx":2BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBalcao.frx":349C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBalcao.frx":4178
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBalcao.frx":4A54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6735
      Left            =   90
      TabIndex        =   18
      Top             =   450
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      TextStyleFixed  =   1
      GridLinesFixed  =   1
      MergeCells      =   4
      PictureType     =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"FrmBalcao.frx":4EA8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label l 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "B a l c ã o   d e   V e n d a s"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   810
      TabIndex        =   13
      Top             =   0
      Width           =   4425
   End
End
Attribute VB_Name = "FrmBalcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CboPassa                                                As Boolean
Public QtdAtual                                             As Long

Private Sub Command_Click()
1   Form_Resize
End Sub

Private Sub Ativar_Timer()
1   FPreco.Visible = False
2   Ativar.Enabled = False
End Sub

Private Sub B_Click(index As Integer)
1   If index = 0 Then
2       Novo
3   ElseIf index = 1 Then
4       NovoIten
5   ElseIf index = 2 Then
6       Excluir
7   ElseIf index = 3 Then
8       FinalizarVenda
9   ElseIf index = 4 Then
10      If MsgBox("Deseja Sair da Tela de Vendas?", vbYesNo + vbDefaultButton2 + vbQuestion, App.Title) = vbYes Then
            '        FrmPrincipal.Timer.Enabled = True
            '        FrmPrincipal.mnuCadastro.Visible = True
            '        FrmPrincipal.mnuMov.Visible = True
            '        FrmPrincipal.mnuRelatorios.Visible = True
            '        FrmPrincipal.mnuDiversos.Visible = True
11          Unload Me
12      End If
13  End If
End Sub

Private Sub CboCliente_Click()
1   If CboCliente.ListIndex >= 0 Then
2       TxtCodCliente.Text = CboCliente.ItemData(CboCliente.ListIndex)
3   End If
End Sub

Private Sub CboCond_GotFocus()
1   SendKeys "{F4}"
End Sub

Private Sub CboCond_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then TxtDesconto.SetFocus
End Sub

Private Sub CboProd_Change()
    Dim X                                                   As Long
    Dim St                                                  As Long
    Dim Texto                                               As String

1   If CboPassa = True Then Exit Sub
2   St = CboProd.SelStart
3   If St = 0 Then Exit Sub


4   Texto = UCase(Left(CboProd.Text, St))
5   For X = 0 To CboProd.ListCount - 1
6       If Texto = UCase(Left(CboProd.List(X), Len(Texto))) Then
7           Texto = CboProd.List(X)
8           Exit For
9       End If
10  Next X
11  CboPassa = True
12  CboProd.Text = Texto
13  CboProd.SelStart = St
14  CboProd.SelLength = Len(CboProd.Text) - St

15  CboPassa = False

End Sub

Private Sub CboProd_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       BuscaProdNome
3   ElseIf KeyCode = vbKeyBack Then
4       CboPassa = True
5       ListaChage = False
6       X = CboProd.SelStart
7       If X = 0 Then
8           CboPassa = True
9           Exit Sub
10      End If
        'CboProd.SelText = ""
11      If Len(CboProd.Text) > 0 Then
12          CboProd.Text = Left(CboProd.Text, X)    'Len(CboProd.Text) - 1)
13          CboProd.SelStart = Len(CboProd.Text)
14      End If
15      CboPassa = False
16      ListaChage = True
17      KeyCode = 0
18  ElseIf KeyCode = 27 Then
19      CboPassa = False
20      CboProd.Visible = False
21      Grid.SetFocus
22  End If

End Sub

Private Sub CboProd_LostFocus()
1   CboProd.Visible = False
End Sub

Private Sub CboVendedor_Click()
1   If CboVendedor.ListIndex >= 0 Then
2       TxtCodVend.Text = CboVendedor.ItemData(CboVendedor.ListIndex)
3   End If
End Sub

Private Sub Form_Activate()
1   Form_Resize
2   FPreco.Left = (Me.ScaleWidth / 2) - (FPreco.Width / 2)
3   FPreco.Top = (Me.ScaleHeight / 2) - (FPreco.Height / 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = vbKeyF1 Then B_Click 4
2   If KeyCode = vbKeyF3 Then FinalizarVenda
3   If KeyCode = vbKeyF2 Then Novo
End Sub

Private Sub Form_Load()
1   QtdAtual = 0
2   Form_Resize
3   AddProd
4   AddCboCliente
5   AddCboVendedor
End Sub

Private Sub Form_Resize()
1   L.Top = 0
2   L.Left = 0
    'l.Height = Me.ScaleHeight
3   L.Width = Me.ScaleWidth
4   FCTop.Top = L.Height
5   FCTop.Left = 0
6   FCTop.Width = Me.ScaleWidth

7   Grid.Left = 0
8   Grid.Top = L.Height + FCTop.Height + 100
9   Grid.Width = Me.ScaleWidth
10  Grid.Height = Me.ScaleHeight - B(0).Height - L.Height - 100 - FCTop.Height
11  Ftotal.Left = 0
12  Ftotal.Top = Me.ScaleHeight - Ftotal.Height
13  Tbar.Left = Me.ScaleWidth - Tbar.Width
14  Tbar.Top = Me.ScaleHeight - Tbar.Height
15  Tbar.Refresh
End Sub


Private Static Sub Novo()
1   On Error GoTo Trata_Erro

2   lblTotal.Caption = "0,00"
3   B(1).Enabled = True
4   B(2).Enabled = True
5   B(3).Enabled = True
6   Grid.Clear
7   Grid.Rows = 2
8   Grid.FormatString = "Codigo|Descrição                                                                          |Qtd              |Valor                     |Total                    "
9   Grid.Enabled = True
10  Grid.SetFocus
11  B(0).SetFocus
12  CboCliente.ListIndex = -1
13  CboCond.ListIndex = -1
14  CboVendedor.ListIndex = -1
15  TxtCodCliente.Text = ""
16  TxtCodVend.Text = ""
17  TxtDesconto.Text = ""
18  TxtMotivo.Text = ""
19  TxtCodCliente.SetFocus
Trata_Erro:
20  E
End Sub


Private Sub NovoIten()
1   On Error GoTo Trata_Erro
2   Grid.Refresh
3   X = Grid.Rows - 1
4   Grid.Col = 1
5   If Grid.TextMatrix(X, 0) <> "" Then
6       Grid.AddItem ""
7   End If
8   Grid.Row = Grid.Rows - 1
9   grid_DblClick

Trata_Erro:
10  E
End Sub

Private Sub grid_DblClick()
1   On Error GoTo Trata_Erro
2   If Grid.Col <> 2 Then
3       Grid.Col = 0
4   End If

5   Txt.Top = Grid.Top + Grid.CellTop
6   Txt.Left = Grid.Left + Grid.CellLeft
7   Txt.Height = Grid.CellHeight
8   Txt.Width = Grid.CellWidth
9   Txt.Text = Grid.Text
10  Txt.Visible = True
11  Txt.SelStart = 0
12  Txt.SelLength = Len(Txt.Text)
13  Txt.SetFocus
Trata_Erro:
14  E
End Sub


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyDelete Then
2       grid_DblClick
3   ElseIf KeyCode = vbKeyDelete Then
4       B_Click 2
5   End If
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       If Grid.Col = 0 Then
3           If BuscaProd() = True Then
4               Grid.Col = 2
5               grid_DblClick
6           Else
7               Txt.SelStart = 0
8               Txt.SelLength = Len(Txt.Text)
9               Txt.SetFocus
10          End If
11      ElseIf Grid.Col = 2 Then
12          If Val(Txt.Text) > QtdAtual Then
13              MsgBox "Caro usuario, não há produtos suficiente no estoque!" & Chr(13) & "Atualmente há somente " & QtdAtual, vbCritical, "Atenção"
14              Txt.SetFocus
15              Exit Sub
16          End If
17          If Trim(Txt.Text) = "" Then
18              Grid.TextMatrix(Grid.Row, 2) = "1"
19          Else
                '            Grid.TextMatrix(Grid.Row, 2) = Format(Txt.Text, "####0.000")
20              Grid.TextMatrix(Grid.Row, 2) = Format(Txt.Text, "####0")
21          End If
22          SomaTotal
23          lblPreco.Caption = Grid.TextMatrix(Grid.Row, 4)
24          Ativar.Enabled = True
25          Ativar.Interval = 2000
26          FPreco.Visible = True
27          Grid.Row = Grid.Rows - 1
28          Grid.Col = 0
29          B_Click 1
30      End If

31  ElseIf KeyCode = 27 Then
32      Txt.Visible = False
33      Grid.SetFocus
34  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
35      Txt.Visible = False
36      Grid.SetFocus
37  End If

End Sub

Private Sub Txt_KeyPress(KeyAscii As Integer)
1   If KeyAscii = Asc("+") Or KeyAscii = Asc("-") Or KeyAscii = Asc("*") Or KeyAscii = Asc("/") Then
2       If Grid.Col = 0 Then
3           Grid.Col = 1
4           CboProd.Top = Grid.Top + Grid.CellTop
5           CboProd.Left = Grid.Left + Grid.CellLeft
            'Txt.Height = Grid.CellHeight
6           CboProd.Width = Grid.CellWidth
7           CboProd.Text = Grid.Text
8           CboProd.Visible = True
9           CboProd.SelStart = 0
10          CboProd.SelLength = Len(Txt.Text)
11          CboProd.SetFocus
12      Else
13          KeyAscii = 0
14      End If
15  End If
16  KeyAscii = Num(KeyAscii)
End Sub

Private Sub Txt_LostFocus()
1   Txt.Visible = False
End Sub
Private Function BuscaProd() As Boolean
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As DAO.Recordset

2   If Trim(Txt.Text) = "" Then Exit Function

3   Sql = "Select * From Produtos Where Codigo ='" & Txt.Text & "'"
4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
5   If Rs.EOF = False Then
6       If Rs!Atual <= 0 Then
7           MsgBox "Caro usuario, no momento não há produtos no estoque!", vbCritical, "Atenção"
8           BuscaProd = False
9           Exit Function
10      End If
11      QtdAtual = Rs!Atual
12      With Grid
13          .TextMatrix(.Row, 0) = Rs!Codigo
14          .TextMatrix(.Row, 1) = Rs!Descricao
15          .TextMatrix(.Row, 2) = "1"
16          .TextMatrix(.Row, 3) = Format(Rs!Venda, "###,###,##0.00")
17          .TextMatrix(.Row, 4) = Rs!Venda
18      End With
19      BuscaProd = True
20  Else
21      MsgBox "Produto não Encontrado!", vbCritical, App.Title
22      BuscaProd = False
23  End If
24  Rs.Close
25  SomaTotal
Trata_Erro:
26  E
End Function

Private Sub SomaTotal()
1   On Error GoTo Trata_Erro
    Dim X As Long, Total                                    As Double

2   Total = 0
3   For X = 1 To Grid.Rows - 1
4       If Grid.TextMatrix(X, 2) <> "" Then
5           Grid.TextMatrix(X, 4) = Format((CCur(Grid.TextMatrix(X, 2)) * CCur(Grid.TextMatrix(X, 3))), "###,###,##0.00")
6       End If
7   Next X



8   For X = 1 To Grid.Rows - 1
9       If Grid.TextMatrix(X, 4) <> "" Then
10          Total = Total + CCur(Grid.TextMatrix(X, 4))
11      End If
12  Next X
13  lblTotal.Caption = Format(Total, "###,###,##0.00")

Trata_Erro:
14  E
End Sub

Private Sub Excluir()
1   On Error GoTo Trata_Erro

2   If Trim(Grid.TextMatrix(Grid.Row, 0)) <> "" Then
3       If MsgBox("Confirma exclusão deste iten?" & Chr(13) & Grid.TextMatrix(Grid.Row, 1), vbYesNo + vbDefaultButton2 + vbQuestion + vbSystemModal, App.Title) = vbYes Then
4           If Grid.Row = 1 And Grid.Rows = 2 Then
5               Grid.Clear
6               Grid.Rows = 2
7               Grid.FormatString = "Codigo|Descrição                                                                          |Qtd              |Valor                     |Total                    "
8           Else
9               Grid.RemoveItem Grid.Row
10          End If
11      End If
12  Else
13      If Grid.Row = 1 And Grid.Rows = 2 Then
14          Grid.Clear
15          Grid.Rows = 2
16          Grid.FormatString = "Codigo|Descrição                                                                          |Qtd              |Valor                     |Total                    "
17      Else
18          Grid.RemoveItem Grid.Row
19      End If
20  End If
21  SomaTotal
Trata_Erro:
22  E
End Sub


Private Sub FinalizarVenda()
1   On Error GoTo Trata_Erro

    Dim RsCompra                                            As Recordset
    Dim IdCompra                                            As String

2   If Grid.TextMatrix(1, 4) = "" Then Exit Sub

3   If CboVendedor.ListIndex = -1 Then
4       MsgBox "Caro usuario, Vendedor invalido!", vbCritical, "Atenção"
5       CboVendedor.SetFocus
6       Exit Sub
7   End If


8   If CboCond.ListIndex = -1 Then
9       MsgBox "Caro Usuario, Codição de pagamento invalida!", vbCritical, "Atenção"
10      CboCond.SetFocus
11      Exit Sub
12  End If

13  If Trim(TxtDesconto.Text) <> "" Then
14      If Trim(TxtMotivo.Text) = "" Then
15          MsgBox "Voce Tem que esplicar o motivo do Desconto!", vbCritical, "Atenção"
16          TxtMotivo.SetFocus
17          Exit Sub
18      End If
19  End If


20  Vendas.Cancelar = False

21  FrmFinalVenda.Show 1

22  If Vendas.Cancelar = True Then Exit Sub

23  Sql = "Select * From Vendas"
24  Set RsCompra = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)

25  If Grid.TextMatrix(1, 0) = "" Then
26      MsgBox "Impossivel finalizar esta compra, pois não há itens nela!", vbCritical, "Atenção"
27      Grid.SetFocus
28      Exit Sub
29  End If

30  Txt.Visible = False
31  RsCompra.AddNew
32  If CboCliente.ListIndex <> -1 Then
33      Vendas.IdCliente = CboCliente.ItemData(CboCliente.ListIndex)
34  Else
35      Vendas.IdCliente = -1
36  End If
37  RsCompra!IdCliente = Vendas.IdCliente
38  RsCompra!IdVendedor = CboVendedor.ItemData(CboVendedor.ListIndex)
39  RsCompra!IdCondPag = CboCond.ListIndex
40  RsCompra!Motivo = TxtMotivo.Text
41  RsCompra!Desconto = IIf(Trim(TxtDesconto.Text) = "", "0", TxtDesconto.Text)
42  RsCompra!TotalNota = lblTotal.Caption
43  RsCompra!Obs = TxtObs.Text
44  RsCompra!Data = Date
45  RsCompra.Update
46  RsCompra.MoveLast
47  IdCompra = RsCompra!Id
48  IdCompraGeral = IdCompra
49  RsCompra.Close


50  For X = 1 To Grid.Rows - 1
51      If Grid.TextMatrix(X, 0) <> "" Then
52          NovoEdit " id =-1"
53          Edit "IdVendas", IdCompra
54          Edit "Codigo", Grid.TextMatrix(X, 0)
55          Edit "Descricao", Grid.TextMatrix(X, 1)
56          Edit "Qtd", Grid.TextMatrix(X, 2)
57          Edit "Valor", Grid.TextMatrix(X, 3)
58          Edit "Total", Grid.TextMatrix(X, 4)
59          Edit "Data", Date
60          MontaSql "ItenVenda"
61          Sql = "Update Produtos Set Atual = atual - " & Val(Grid.TextMatrix(X, 2)) & " where Codigo ='" & Grid.TextMatrix(X, 0) & "'"
62          BancoDeDados.Execute Sql
63      End If
64  Next X
    'If Vendas.IdCliente = -1 Then
    '    FrmRelVenda1.LblCond.Caption = BuscaCond()
    '    FrmRelVenda1.LblVendedor.Caption = buscaVendedor(Vendas.IdVendedor)
    '    FrmRelVenda1.Config
    'Else
    '    FrmRelVenda1.LblCond.Caption = BuscaCond()
    '    FrmRelVenda.LblVendedor.Caption = buscaVendedor(Vendas.IdVendedor)
    '    FrmRelVenda.Config
    'End If

65  For XCont = 0 To 1
66      Printer.FontSize = 20
67      Printer.Print Tab(2); "Orçamento"
68      Printer.Print
69      Printer.FontSize = 10
70      Printer.Print Tab(3); Empresa.Nome; Tab(60); "CNPJ:" & Empresa.CNPJ
71      Printer.Print Tab(3); "Endereço:" & Empresa.Endereco; Tab(60); "Bairro:" & Empresa.Bairro
72      Printer.Print Tab(3); "Cidade:" & Empresa.Cidade; Tab(60); "Uf " & Empresa.Estado
73      Printer.Print Tab(3); "Cep:" & Empresa.Cep; Tab(60); "Telefone:" & Empresa.Telefone
74      If Vendas.IdCliente <> -1 Then
75          Printer.Print Tab(3); "Cliente :" & BuscaClienteRel(Vendas.IdCliente)
76      End If
77      Printer.Print Tab(3); "Vendedores :" & buscaVendedor(Vendas.IdVendedor); Tab(60); "Condição de Pagamento : " & Vendas.DescricaoCond
78      Printer.Print

79      Printer.Print Tab(3); "Codigo"; Tab(12); "Descrição"; Tab(55); "Qtd."; Tab(75); "Valor"; Tab(90); "Total"

80      For X = 1 To Grid.Rows - 1
81          XX = Printer.CurrentY
82          Printer.Print Tab(3); Grid.TextMatrix(X, 0); Tab(12); Grid.TextMatrix(X, 1)

83          YY = 6300 - Printer.TextWidth(Grid.TextMatrix(X, 2))
84          Printer.CurrentY = XX
85          Printer.CurrentX = YY
86          Printer.Print Grid.TextMatrix(X, 2)


            '    Printer.CurrentY = XX
            '    Printer.Print Tab(60); Grid.TextMatrix(X, 2)

87          YY = 8500 - Printer.TextWidth(Grid.TextMatrix(X, 3))
88          Printer.CurrentY = XX
89          Printer.CurrentX = YY

90          Printer.Print Grid.TextMatrix(X, 3)

91          YY = 10100 - Printer.TextWidth(Grid.TextMatrix(X, 4))
92          Printer.CurrentY = XX
93          Printer.CurrentX = YY
94          Printer.Print Grid.TextMatrix(X, 4)


95      Next X

96      XX = Printer.CurrentY
97      Printer.Print Tab(70); "Total "
98      YY = 11000 - Printer.TextWidth(lblTotal.Caption)

99      Printer.CurrentY = XX
100     Printer.CurrentX = YY
101     Printer.Print lblTotal.Caption


102     XX = Printer.CurrentY
103     Printer.Print Tab(70); "Desconto "
104     YY = 11000 - Printer.TextWidth(Format(Vendas.Desconto, "###,###,##0.00"))

105     Printer.CurrentY = XX
106     Printer.CurrentX = YY
107     Printer.Print Format(Vendas.Desconto, "###,###,##0.00")


108     XX = Printer.CurrentY
109     Printer.Print Tab(70); "Total a Pagar :"
110     YY = 11000 - Printer.TextWidth(Format(CCur(lblTotal.Caption) - Vendas.Desconto, "###,###,##0.00"))

111     Printer.CurrentY = XX
112     Printer.CurrentX = YY
113     Printer.Print Format(CCur(lblTotal.Caption) - Vendas.Desconto, "###,###,##0.00")
114     Printer.Print ""
115     Printer.Print "Obs.: " & TxtObs.Text
116     Printer.Print ""
117     Printer.Print ""
118     Printer.Print ""
119     If XCont = 0 Then
120         Printer.Print String(200, "-")
121         Printer.Print ""
122     End If
123 Next XCont
124 Printer.EndDoc

125 Novo

Trata_Erro:
126 E

End Sub

Private Sub AddProd()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   Sql = "Select * From Produtos Order By Descricao"

3   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
4   CboProd.Clear
5   If Rs.EOF = False Then
6       Do While Not Rs.EOF
7           CboProd.AddItem Rs!Descricao
8           CboProd.ItemData(CboProd.ListCount - 1) = Rs!Id
9           Rs.MoveNext
10      Loop
11  End If

12  Rs.Close
Trata_Erro:
13  E
End Sub
Private Sub BuscaProdNome()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   If Trim(CboProd.Text) = "" Then Exit Sub

3   Sql = "Select * From Produtos Where Descricao ='" & CboProd.Text & "'"
4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
5   If Rs.EOF = False Then
6       Grid.TextMatrix(Grid.Row, 0) = Rs!Codigo
7       Grid.TextMatrix(Grid.Row, 1) = Rs!Descricao
8       Grid.TextMatrix(Grid.Row, 2) = "1"
9       Grid.TextMatrix(Grid.Row, 3) = Rs!Venda
10      Grid.TextMatrix(Grid.Row, 4) = Rs!Venda
11      Grid.Col = 2
12      grid_DblClick
13  Else
14      MsgBox "Não foi possivel localizar o produto com esta descrição!", vbCritical, App.Title
15      CboProd.Text = ""
16      CboProd.SetFocus
17  End If

18  Rs.Close
Trata_Erro:
19  E
End Sub
Private Sub AddCboCliente()
1   On Error GoTo Trata_Erro

    Dim Rs                                                  As DAO.Recordset

2   CboCliente.Clear
3   Sql = "Select * From Cliente Order by Razao"
4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
5   If Rs.EOF = False Then
6       Do While Not Rs.EOF
7           CboCliente.AddItem Rs!Razao
8           CboCliente.ItemData(CboCliente.ListCount - 1) = Rs!Id
9           Rs.MoveNext
10      Loop
11  End If
12  Rs.Close

Trata_Erro:
13  E
End Sub
Private Sub AddCboVendedor()
1   On Error GoTo Trata_Erro

    Dim Rs                                                  As DAO.Recordset

2   CboVendedor.Clear
3   Sql = "Select * From Vendedor Order by Razao"
4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
5   If Rs.EOF = False Then
6       Do While Not Rs.EOF
7           CboVendedor.AddItem Rs!Razao
8           CboVendedor.ItemData(CboVendedor.ListCount - 1) = Rs!Id
9           Rs.MoveNext
10      Loop
11  End If
12  Rs.Close

Trata_Erro:
13  E
End Sub

Private Sub TxtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Passa                                               As Boolean
1   If KeyCode = 13 Then
2       Passa = False
3       For X = 0 To CboCliente.ListCount - 1
4           If CboCliente.ItemData(X) = TxtCodCliente.Text Then
5               Passa = True
6               CboCliente.ListIndex = X
7               CboCond.SetFocus
8               Exit Sub
9           End If
10      Next X
11      If Passa = False Then
12          MsgBox "Cliente não  encontrado !", vbCritical, "Atenção"
13          CboCliente.ListIndex = -1
14          TxtCodCliente.Text = ""
15          TxtCodCliente.SetFocus
16      End If
17  End If
End Sub
Private Sub TxtCodVend_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Passa                                               As Boolean

1   If KeyCode = 13 Then
2       Passa = False
3       For X = 0 To CboVendedor.ListCount - 1
4           If CboVendedor.ItemData(X) = TxtCodVend.Text Then
5               CboVendedor.ListIndex = X
6               Grid.SetFocus
7               Passa = True
8               Exit Sub
9           End If
10      Next X
11      If Passa = False Then
12          MsgBox "Vendedor não Encontrado!", vbCritical, "Atenção"
13          TxtCodVend.Text = ""
14          TxtCodVend.SetFocus
15          Exit Sub
16      End If
17      Grid.SetFocus
18  End If
End Sub

Private Sub TxtDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then TxtCodVend.SetFocus
End Sub

Private Sub TxtDesconto_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub
