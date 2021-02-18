VERSION 5.00
Begin VB.Form FrmFinalVenda 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6090
   ClientLeft      =   1530
   ClientTop       =   2655
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtValorRecebido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   90
      TabIndex        =   0
      Top             =   2970
      Width           =   7185
   End
   Begin VB.CommandButton B 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Voltar a Tela"
      Height          =   975
      Index           =   2
      Left            =   5790
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5070
      Width           =   1575
   End
   Begin VB.CommandButton B 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Finalizar"
      Height          =   975
      Index           =   1
      Left            =   4230
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5070
      Width           =   1575
   End
   Begin VB.Label Troco 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Troco :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   7
      Top             =   3960
      Width           =   7185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Recebido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   6
      Top             =   2100
      Width           =   7185
   End
   Begin VB.Label LblValorPagar 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor A Pagar "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   7365
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Finalizar Venda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "FrmFinalVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub B_Click(index As Integer)
1   If index = 1 Then
        '    If Trim(TxtDesconto.Text) = "" Then TxtDesconto.Text = "0"
        '    Vendas.MotivoDesconto = TxtMotivo.Text
2       Vendas.Cancelar = False
        '    Vendas.Desconto = TxtDesconto.Text

3       Unload Me
4   ElseIf index = 2 Then
5       Vendas.Cancelar = True
6       Unload Me
7   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{tab}"
2   If KeyCode = 27 Then
3       Vendas.Cancelar = True
4       Unload Me
5   End If
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


Private Sub BuscaCliente()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As DAO.Recordset

2   If CboCliente.ListIndex = -1 Then Exit Sub


3   Sql = "Select * From Cliente Where Id =" & CboCliente.ItemData(CboCliente.ListIndex)
4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

5   If Rs.EOF = False Then
6       LblDoc.Caption = Rs!CNPJ
7       LblTelefone.Caption = Rs!TelefoneE
8       LblTipoCliente.Caption = Rs!TipoCliente
9   Else
10      LblDoc.Caption = "???.???.???-??"
11      LblTelefone.Caption = "?? ???? - ????"
12      LblTipoCliente.Caption = "????????????????????????????????"
13  End If
14  Rs.Close

Trata_Erro:
15  E
End Sub

Private Sub BuscaVendedores()
    Dim Rs                                                  As DAO.Recordset

1   If CboVendedor.ListIndex = -1 Then Exit Sub

2   Sql = "Select * From Vendedor Where Id =" & CboVendedor.ItemData(CboVendedor.ListIndex)
3   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
4   If Rs.EOF = False Then
5       LblComissao.Caption = Rs!Comissao
6   Else
7       LblComissao.Caption = "0%"
8   End If
9   Rs.Close


End Sub


Private Sub BuscaCond()
    Dim Rs                                                  As DAO.Recordset

1   Sql = "Select * From CodPag"
2   CboCond.Clear

3   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

4   If Rs.EOF = False Then
5       Do While Not Rs.EOF
6           CboCond.AddItem Rs!Descricao
7           CboCond.ItemData(CboCond.ListCount - 1) = Rs!Id
8           Rs.MoveNext
9       Loop
10  End If

11  Rs.Close

End Sub

Private Sub Form_Load()
    Dim Desconto                                            As Double
1   If Trim(FrmBalcao.TxtDesconto.Text) = "" Then
2       Desconto = 0
3   Else
4       Desconto = FrmBalcao.TxtDesconto.Text
5   End If
6   LblValorPagar.Caption = Format(CCur(FrmBalcao.lblTotal.Caption) - Desconto, "###,###,##0.00")

End Sub
Private Sub TxtValorRecebido_Change()
    Dim ValorPagar                                          As Double
    Dim ValorRecebido                                       As Double

1   If Trim(TxtValorRecebido.Text) <> "" Then
2       ValorRecebido = CCur(TxtValorRecebido.Text)
3   Else
4       ValorRecebido = 0
5   End If

6   If Trim(LblValorPagar.Caption) <> "" Then
7       ValorPagar = LblValorPagar.Caption
8   Else
9       ValorPagar = 0
10  End If



11  Troco.Caption = Format(ValorRecebido - ValorPagar, "###,###,##0.00")

End Sub

Private Sub TxtValorRecebido_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub
