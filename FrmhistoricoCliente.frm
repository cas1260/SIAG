VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmhistoricoCliente 
   Caption         =   "Relatorio de Historico de Cliente"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FrmhistoricoCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11880
   Begin MSComctlLib.ProgressBar Barra 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   10
      Top             =   6450
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   8760
      TabIndex        =   9
      Top             =   300
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   300
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   10320
      TabIndex        =   5
      Top             =   7110
      Width           =   1575
   End
   Begin VB.ComboBox CboCliente 
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   4545
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   5805
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   11805
      ExtentX         =   20823
      ExtentY         =   10239
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
   Begin MSMask.MaskEdBox MskInicial 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   300
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
      Left            =   1230
      TabIndex        =   1
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   1290
      TabIndex        =   8
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "Data inicial:"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   2490
      TabIndex        =   6
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "FrmhistoricoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimir_Click()
1   Web.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command1_Click()
1   PesquisaRel
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   Me.Height = 7065
2   Me.Width = Screen.Width
3   Web.Left = 0
4   Web.Width = Me.Width - 10
5   Centra Me
6   AddCboCliente
7   Web.Navigate "about:blank"
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

Private Sub PesquisaRel()
    Dim Rs                                                  As Recordset
    Dim Sql                                                 As String
    Dim Total                                               As Double
    Dim Tipo                                                As String

1   If CboCliente.ListIndex = -1 Then
2       MsgBox "Cliente Invalido!", vbCritical, "Atenção"
3       CboCliente.SetFocus
4       Exit Sub
5   End If

6   If Not IsDate(MskInicial.Text) Then
7       MsgBox "Data Inicial invalida!", vbCritical, "Atenção"
8       MskInicial.SetFocus
9       Exit Sub
10  End If

11  If Not IsDate(MskFinal.Text) Then
12      MsgBox "Data Inicial invalida!", vbCritical, "Atenção"
13      MskFinal.SetFocus
14      Exit Sub
15  End If


16  Sql = "SELECT vendas.id as IdVenda,* FROM Vendas Left JOIN  Vendedor on  Vendas.IdVendedor = Vendedor.Id Where Vendas.IdCliente = " & CboCliente.ItemData(CboCliente.ListIndex) & " And Data >=#" & Format(MskInicial.Text, "MM/DD/YYYY") & "# and Data <= #" & Format(MskFinal.Text, "MM/DD/YYYY") & "#"
17  Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)

18  If Rs.EOF = True Then
19      Web.Navigate "about:blank"
20      MsgBox "Caro Usuario, não há relatario para ser exibidos neste periodo!", vbCritical, "Atenção"
21      Rs.Close
22      Exit Sub
23  End If

24  Barra.Min = 0
25  Barra.Value = 0
26  Barra.Visible = True
27  Rs.MoveLast
28  Rs.MoveFirst
29  Barra.Max = Rs.RecordCount
30  DoEvents
31  DoEvents
32  DoEvents
33  Open "C:\Temp.Neo" For Output As #1
34  Print #1, "<title>Relatorio de Historico de Cliente</title>"
35  Print #1, "<body>"
36  Print #1, "<table border='0' width='750' cellspacing='0' cellpadding='0' height='112'>"
37  Print #1, "<tr>"
38  Print #1, "<td width='14%' height='18'>"
39  Print #1, "<p align='center'><font size='2' face='Arial'><b>" & Date & "</b></font></td>"
40  Print #1, "<td width='71%' height='18'>"
41  Print #1, "<p align='center'><font size='2' face='Arial'><b>Relatório de Historio de"
42  Print #1, "Cliente</b></font></td>"
43  Print #1, "<td width='15%' height='18'>"
44  Print #1, "<p align='center'><font size='2' face='Arial'><b>" & Time & "</b></font></td>"
45  Print #1, "</tr>"
46  Print #1, "<tr>"
47  Print #1, "<td width='14%' height='21'>"
48  Print #1, "<p align='right'><font size='2' face='Arial'><b>Cliente :</b></font></td>"
49  Print #1, "<td width='71%' height='21'>" & CboCliente.Text & "</td>"
50  Print #1, "<td width='15%' height='21'></td>"
51  Print #1, "</tr>"
52  Print #1, "<tr>"
53  Print #1, "<td width='14%' height='30'><font size='2' face='Arial'><b>&nbsp;</b></font></td>"
54  Print #1, "<td width='71%' height='30'></td>"
55  Print #1, "<td width='15%' height='30'><font size='2' face='Arial'><b>&nbsp;</b></font></td>"
56  Print #1, "</tr>"
57  Print #1, "<tr>"
58  Print #1, "<td width='100%' colspan='3' height='43'>"
59  Print #1, "<table border='1' width='100%' cellspacing='0' cellpadding='0' bordercolorlight='#000000' bordercolordark='#FFFFFF' height='32'>"
60  Print #1, "<tr>"
61  Print #1, "<td width='12%' height='19' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Data da Compra</font></td>"
62  Print #1, "<td width='28%' height='19' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Vendedor</font></td>"
63  Print #1, "<td width='14%' height='19' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Condição de Pag.</font></td>"
64  Print #1, "<td width='11%' height='19' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Desconto</font></td>"
    'Print #1, "<td width='20%' height='19' colspan='2' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Valor</font></td>"
65  Print #1, "<td width='20%' height='19' bgcolor='#808080'><font size='2' face='Arial' color='#FFFFFF'>Valor</font></td>"
66  Print #1, "</tr>"
67  Total = 0
68  Do While Not Rs.EOF
69      Tipo = "Indefinido"
70      Select Case Rs("IdCondPag")
        Case "0"
71          Tipo = "A Vista"
72      Case "1"
73          Tipo = "15 Dias"
74      Case "2"
75          Tipo = "30 Dias"
76      Case "3"
77          Tipo = "45 Dias"
78      Case "4"
79          Tipo = "60 Dias"
80      Case "5"
81          Tipo = "75 Dias"
82      Case "6"
83          Tipo = "90 Dias"
84      Case "7"
85          Tipo = "120 Dias"
86      End Select
87      Print #1, "<tr>"
88      Print #1, "<td width='12%' height='9'><p align='center'><font face='Arial' size='2'>" & Rs("Data") & "</font></td>"
89      Print #1, "<td width='28%' height='9'><font face='Arial' size='2'>" & Rs("Razao") & "</font></td>"
90      Print #1, "<td width='14%' height='9'><p align='center'><font face='Arial' size='2'>" & Tipo & "</font></td>"
91      Print #1, "<td width='11%' height='9'><p align='Right'><font face='Arial' size='2'>" & Format(Rs("Desconto"), "###,###,##0.00") & "</font></td>"
92      Print #1, "<td width='11%' height='9'><p align='right'><font face='Arial' size='2'>" & Format(Rs("TotalNota"), "###,###,##0.00") & "</font></td>"
        'Print #1, "<td width='9%' height='9'><p align='center'><b><font size='2' face='Arial'><a href='http://" & Rs("IdVenda") & "#'><font color='#008000'>Ver Itens</font></a></font></b></td>"
93      Print #1, "</tr>"
94      Total = Total + (CCur(Rs("TotalNota")) - CCur(Rs("Desconto")))
95      Rs.MoveNext
96      Barra.Value = Barra.Value + 1
97  Loop
98  Print #1, "</table>"
99  Print #1, "</td>"
100 Print #1, "</tr>"
101 Print #1, "</table>"
102 Print #1, ""
103 Print #1, "</body>"
104 Print #1, ""
105 Print #1, "</html>"
106 Print #1, "<p align='right'><font size='2' face='Arial'>Valor Total : " & Format(Total, "###,###,##0.00") & "</font>"
107 Close #1
108 Barra.Visible = False
109 Web.Navigate "C:\Temp.Neo"
End Sub

Private Sub Web_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim IdVenda                                             As String
1   If Right(URL, 1) = "#" Then
2       Cancel = True
3       IdVenda = Right(URL, Len(URL) - 7)
4       IdVenda = Left(IdVenda, Len(IdVenda) - 2)



5   End If
End Sub

