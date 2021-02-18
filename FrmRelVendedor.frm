VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRelVendedor 
   Caption         =   "Relatorio de Vendedor"
   ClientHeight    =   6540
   ClientLeft      =   735
   ClientTop       =   705
   ClientWidth     =   10425
   Icon            =   "FrmRelVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   6375
      Left            =   60
      TabIndex        =   8
      Top             =   810
      Width           =   11865
      ExtentX         =   20929
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox CboVend 
      Height          =   315
      Left            =   2490
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   2985
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   10350
      TabIndex        =   4
      Top             =   7230
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   5580
      TabIndex        =   3
      Top             =   420
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MskInicial 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskFinal 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   420
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "Vendedor"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   2520
      TabIndex        =   7
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRelVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Localiza()
1   On Error GoTo Trata_Erro

    Dim Rs                                                  As DAO.Recordset
    Dim Rs1                                                 As DAO.Recordset

2   MskInicial.Text = Valida(MskInicial)
3   MskFinal.Text = Valida(MskFinal)

4   If CboVend.ListIndex = -1 Then
5       MsgBox "Vendedor invalido!", vbCritical, App.Title
6       CboVend.SetFocus
7       Exit Sub
8   End If

9   Sql = "Select *, Vendedor.id as IdVend From Vendas left join Vendedor on vendas.IdVendedor = Vendedor.Id  Where data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# "
10  Sql = Sql & " And Data <=#" & Format(MskFinal.Text, "MM/DD/YYYY") & "#"
11  Sql = Sql & " And IdVendedor = " & CboVend.ItemData(CboVend.ListIndex)
12  Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)


13  If Rs.EOF = True Then
14      MsgBox "Nâo há vendas neste periodo para este vendedor!", vbCritical, "Atenção"
15      CboVend.SetFocus
16      Rs.Close
17      Exit Sub
18  End If

19  Open "C:\Temp.Neo" For Output As #1
20  Print #1, "<html>"

21  Print #1, "<head>"
22  Print #1, "<meta http-equiv='Content-Language' content='pt-br'>"
23  Print #1, "<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>"
24  Print #1, "<meta name='GENERATOR' content='Microsoft FrontPage 4.0'>"
25  Print #1, "<meta name='ProgId' content='FrontPage.Editor.Document'>"
26  Print #1, "<title>Listagem de Vendas por Vendedor </title>"
27  Print #1, "</head>"

28  Print #1, "<body>"
29  Print #1, "<table border='0' width='90%'>"
30  Print #1, "<tr>"
31  Print #1, "<td width='9%'><p align='center'><font size='1' face='Arial'>" & Time & "</font></td>"
32  Print #1, "<td width='77%'><p align='center'><b><font size='3' face='Arial'>Listagem de Vendedores</font></b></td>"
33  Print #1, "<td width='14%'><p align='right'><font size='1' face='Arial'>" & Date & "</font></td>"
34  Print #1, "</tr>"
35  Print #1, "</Table><Table width='682' border='0'>"
36  Print #1, "<tr>"
37  Print #1, "<td width='319'><font size='1' face='Arial'>Vendedore : " & Rs!Razao & "</font></td>"
38  Print #1, "<td width='38'><font size='1' face='Arial'>Total :</font></td>"
39  TotalVend = BuscaTotalVend(Rs!IdVend)
40  ValorVend = (TotalVend / 100) * Rs!Comissao
41  Print #1, "<td width='106'><p align='right'><font size='1' face='Arial'>" & Format(TotalVend, "###,###,##0.00") & "</font></td>"
42  Print #1, "<td width='56'><font size='1' face='Arial'>Comissão :</font></td>"
43  Print #1, "<td width='41'><font size='1' face='Arial'>" & Rs!Comissao & "%</font></td>"
44  Print #1, "<td width='92'><p align='right'><font size='1' face='Arial'>" & Format(ValorVend, "###,###,##0.00") & "</font></td>"
45  Print #1, "</tr>"
46  Print #1, "</table>"

    'Do While Not Rs.EOF
47  Print #1, "<table border='1' width='90%' cellspacing='0' cellpadding='0'>"
48  Print #1, "<tr>"
49  Print #1, "<td width='59%'><font face='Arial' size='1'>Produto</font></td>"
50  Print #1, "<td width='17%'><font face='Arial' size='1'>Qtd</font></td>"
51  Print #1, "<td width='24%'><font face='Arial' size='1'>Valor</font></td>"
52  Print #1, "</tr>"

53  Sql = "Select * from ItenVenda left join Vendas on ItenVenda.IdVendas = Vendas.Id Where ItenVenda.data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# "
54  Sql = Sql & " And ItenVenda.Data <=#" & Format(MskFinal.Text, "MM/DD/YYYY") & "#"
55  Sql = Sql & " And Vendas.IdVendedor = " & Rs!IdVend & " Order By ItenVenda.Codigo"
56  Set Rs1 = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)


57  If Rs1.EOF = False Then
58      Qtd = 0
59      Tt = 0
60      Do While Not Rs1.EOF
61          CodigoAtual = Rs1!Codigo

62          Do While CodigoAtual = Rs1!Codigo
63              Qtd = Qtd + Rs1!Qtd
64              Tt = Tt + Rs1!Valor
65              Rs1.MoveNext
66              If Rs1.EOF = True Then Exit Do
67          Loop
68          Print #1, "<tr>"
69          If Rs1.EOF = True Then
70              If Rs1.RecordCount <> 0 Then
71                  Rs1.MovePrevious
72              End If
73          End If
74          Print #1, "<td width='59%'><font face='Arial' size='1'>" & Rs1!Descricao & "</font></td>"

75          Print #1, "<td width='17%'><p align='right'><font face='Arial' size='1'>" & Qtd & "</font></td>"
76          Print #1, "<td width='24%'><p align='right'><font face='Arial' size='1'>" & Format(Tt, "###,###,##0.00") & "</font></td>"
77          Print #1, "</tr>"
78          Rs1.MoveNext
79      Loop
80  End If
81  Print #1, "</table>"
82  Print #1, "<br>"
83  Print #1, "</body>"
84  Print #1, "</html>"
85  Rs.MoveNext
    'Loop
86  Close #1

87  Web.Navigate "C:\Temp.Neo"
Trata_Erro:
88  E

End Sub

Private Sub Command1_Click()
1   Localiza
End Sub

Private Sub Command2_Click()
1   Web.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{Tab}"
2   If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()

'Me.Width = 10590
'Me.Height = 6945
'Centra Me
1   On Error GoTo Trata_Erro

2   Sql = "Select * From Vendedor Order By Razao"
3   CboVend.Clear

    Dim Rs                                                  As Recordset


4   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

5   Do While Not Rs.EOF
6       CboVend.AddItem Rs!Razao
7       CboVend.ItemData(CboVend.ListCount - 1) = Rs!Id
8       Rs.MoveNext
9   Loop

10  Rs.Close

11  Open "C:\Temp.Neo" For Output As #1
12  Print #1, "<html></Html>"
13  Close #1

14  Web.Navigate "C:\Temp.neo"



Trata_Erro:
15  E
End Sub


Private Function BuscaTotalVend(IdVend As String)
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   Sql = "Select Sum(TotalNota) As TT from Vendas Where data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# "
3   Sql = Sql & " And Data <=#" & Format(MskFinal.Text, "MM/DD/YYYY") & "#"
4   Sql = Sql & " And IdVendedor = " & IdVend

5   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

6   If Rs.EOF = False Then
7       BuscaTotalVend = 0
        'Do While Not Rs.EOF
8       BuscaTotalVend = Rs!Tt    ' BuscaTotalVend + Rs!Total
9   Else
10      BuscaTotalVend = 0
11  End If
12  Rs.Close

Trata_Erro:
13  E
End Function

Private Sub MskFinal_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then MskFinal.Text = Valida(MskFinal)
End Sub

Private Sub MskInicial_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then MskInicial.Text = Valida(MskInicial)
End Sub
