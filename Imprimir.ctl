VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl Imp 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Imprimir.ctx":0000
   Begin MSComDlg.CommonDialog Com 
      Left            =   720
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame F 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   570
      Begin VB.Image Image1 
         Height          =   240
         Left            =   45
         Picture         =   "Imprimir.ctx":0312
         Top             =   90
         Width           =   240
      End
   End
End
Attribute VB_Name = "Imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NumeroColunas                                           As Long    '1= 2 'PERSONALIZE
Dim LarguraColuna(1 To 20)                                  As Long
Dim TituloColuna(1 To 20)                                   As String
Dim CentralizaColuna(1 To 20)                               As Boolean
Dim TextoColuna(1 To 20)                                    As Long
Dim TipoLetraColuna(1 To 20)                                As String
Dim TamanhoLetraColuna(1 To 20)                             As Long
Dim LetraNegritoColuna(1 To 20)                             As Boolean
Dim QuebraLinhaColuna(1 To 20)                              As Boolean
Dim Papel1                                                  As TipoDePagina
Dim Titulo_1                                                As String
Dim SubTitulo_1                                             As String
Dim Rodape_1                                                As String
Public GridImpir                                            As MSFlexGrid

Public Enum TipoDePagina
    Retrato = 0    '"Retrato"
    Paisagem = 1    '"Paisagem"
End Enum

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
1   Papel1 = PropBag.ReadProperty("Papel", Retrato)
2   Titulo_1 = PropBag.ReadProperty("Titulo", "")
3   SubTitulo_1 = PropBag.ReadProperty("SubTitulo", "")
4   Rodape_1 = PropBag.ReadProperty("Rodape", "")
End Sub

Private Sub UserControl_Resize()
1   UserControl.Height = F.Height + F.Top
2   UserControl.Width = F.Width
End Sub

Private Sub ImprimeCabecalho(Titulo As String, LinhaAuxiliar As String)

    Dim Texto                                               As String

1   Texto = Titulo
2   Printer.FontSize = 17
3   Printer.FontBold = True
4   Printer.CurrentX = Int((Printer.Width - Printer.TextWidth(Texto)) / 2)
5   Printer.CurrentY = 450
6   Printer.Print Texto

    '------------------------------
    'imprimindo a 2ª linha

7   Texto = LinhaAuxiliar
8   Printer.FontSize = 12
9   Printer.FontBold = True
10  Printer.CurrentX = Int((Printer.Width - Printer.TextWidth(Texto)) / 2)
11  Printer.CurrentY = 1000
12  Printer.Print Texto

End Sub

Private Sub ImprimeRodape(Rodape As String, NumeroPagina As Long, AlturaRodape As Long, MargemDireita As Long)

'------------------------------
'imprimindo a 1ª linha

1   Texto = Rodape
2   Printer.FontSize = 8
3   Printer.FontName = "Arial"
4   Printer.FontBold = False
5   Printer.CurrentX = Int((Printer.Width - Printer.TextWidth(Texto)) / 2)
6   Printer.CurrentY = Printer.Height - AlturaRodape
7   Printer.Print Texto

    '------------------------------
    'imprimindo a 2ª linha

8   Texto = "Neo SoftWare {wwww.neobh.com.br} Data: " & Format(Now, "dd/mm/yyyy - hh:mm:ss") & "  -  Pág. " & Trim$(Str$(NumeroPagina))

9   Printer.FontSize = 8
10  Printer.FontBold = False
11  Printer.FontName = "Arial"
12  Printer.CurrentX = Printer.Width - Printer.TextWidth(Texto) - MargemDireita - 700
13  Printer.CurrentY = Printer.Height - AlturaRodape + 300
14  Printer.Print Texto

End Sub

Private Function VerificaLarguraTexto(Texto As String, LarguraColuna As Long) As Long

    Dim LarguraLinha                                        As Long
    Dim LarguraLinhaMaxima                                  As Long
    Dim zz                                                  As Long

1   LarguraLinha = 0
2   LarguraLinhaMaxima = 0

3   Do While Printer.TextWidth(Texto) > LarguraColuna - 100
4       For zz = 1 To Len(Texto)
5           If Printer.TextWidth(Left(Texto, zz)) >= LarguraColuna - 100 Then
6               Do While Mid(Texto, zz, 1) <> " "
7                   zz = zz - 1
8                   If zz < 1 Then GoTo salto
9               Loop
10              LarguraLinha = LarguraLinha + 1
11              Texto = Right(Texto, Len(Texto) - zz)
12              Exit For
13          End If
14      Next
15  Loop

salto:
16  If Len(Texto) > 0 Then LarguraLinha = LarguraLinha + 1

17  VerificaLarguraTexto = LarguraLinha

End Function

Public Function ImprimirRelatorios()
1   On Error GoTo Trata_Erro
    Dim Papel                                               As TipoDePagina
    Dim Titulo1                                             As String
    Dim SubTitulo1                                          As String
    Dim RodapeRel                                           As String
    Dim Grid                                                As MSFlexGrid
    'Setar as Variaves que foram escolhidas na propriedade.

2   Titulo1 = Titulo_1
3   SubTitulo1 = SubTitulo_1
4   RodapeRel = Rodape_1
5   Papel = Papel1
6   Set Grid = GridImpir
    'solicitando confirmação do usuário para impressão de relatório

    'If MsgBox("    Caro(a) usuário(a), confirma a impressão do relatório?", vbQuestion Or vbYesNo Or vbDefaultButton2, "CONFIRMAÇÃO DE IMPRESSÃO") = vbNo Then Exit Function
    Dim ComDialogo                                          As CommonDialog
    'Definindo a impressora a ser utilizada

7   Set ComDialogo = Com
8   ComDialogo.Max = 10000
9   ComDialogo.CancelError = True
10  ComDialogo.Action = 5

    Dim ImprimePaginaEspecifica                             As Boolean

11  NumPagImpIni = ComDialogo.FromPage
12  NumPagImpFin = ComDialogo.ToPage
13  If ComDialogo.Flags = 2 Then
14      ImprimePaginaEspecifica = True
15  Else
16      ImprimePaginaEspecifica = False
17  End If             'rotina de impresso de escolha de pagina

    Dim Titulo                                              As String
    Dim SubTitulo                                           As String
    Dim LinhaAuxiliar                                       As String
    Dim Rodape                                              As String
    Dim CentralizaColunas                                   As Boolean
    Dim NumeroItens                                         As Long
    Dim LarguraLinha                                        As Long
    Dim OrientacaoPapel                                     As String

    'orientação do papel: "Retrato" ou "Paisagem"
18  If Papel = Paisagem Then
19      OrientacaoPapel = "Paisagem"    'PERSONALIZE
20  Else
21      OrientacaoPapel = "Retrato"
22  End If

    'largura de cada coluna (em milímetros)
    'definição das margens do papel (em milímetros)
23  MargemEsquerda = 5  'PERSONALIZE
24  MargemDireita = 5    'PERSONALIZE
25  MargemSuperior = 6    'PERSONALIZE
26  MargemInferior = 5    'PERSONALIZE

    'definição da largura das linhas (em milímetros)
27  LarguraLinha = 5    'PERSONALIZE

    'título, sub-título e rodapé do relatório
28  Titulo = Titulo1    'PERSONALIZE
29  LinhaAuxiliar = SubTitulo1
30  Rodape = RodapeRel

    'número de itens a serem impressos
31  NumeroItens = Grid.Rows - 1    'PERSONALIZE

    'centralizar as colunas no papel
32  CentralizaColunas = True    'PERSONALIZE

    '----------------------------------------------------------------
    'rotinas auxiliares
    '----------------------------------------------------------------

    'definição de variáveis
    Dim Escala                                              As Double
    Dim LarguraTotal                                        As Long

    Dim X                                                   As Long
    Dim y                                                   As Long
    Dim XInicial                                            As Long
    Dim YInicial                                            As Long
    Dim XFinal                                              As Long
    Dim YFinal                                              As Long
    Dim AlturaCabecalho                                     As Long
    Dim AlturaRodape                                        As Long
    Dim NumeroMaximoLinhas                                  As Long
    Dim QualLinha                                           As Long
    Dim dY                                                  As Long
    Dim NumeroPagina                                        As Long
    Dim index                                               As Long


    'setando o valor da escala
33  Escala = 56.67

    'setando a altura do cabeçalho e rodapé
34  AlturaCabecalho = Int(Escala * 30)
35  AlturaRodape = Int(Escala * 30)

    'convertendo a largura das colunas para pixels
36  For index = 1 To NumeroColunas - 1
37      LarguraColuna(index) = Int(LarguraColuna(index) * Escala)
38  Next index

    'convertendo as margens para pixels
39  MargemEsquerda = Int(MargemEsquerda * Escala)
40  MargemDireita = Int(MargemDireita * Escala)
41  MargemSuperior = Int(MargemSuperior * Escala)
42  MargemInferior = Int(MargemInferior * Escala)

    'convertendo a largura da linha para pixels
43  dY = Int(Escala * LarguraLinha)

44  LarguraTotal = 0
45  For index = 1 To NumeroColunas - 1
46      LarguraTotal = LarguraTotal + LarguraColuna(index)
47  Next index

48  YInicial = MargemSuperior + AlturaCabecalho
49  y = YInicial

    '----------------------------------------------------------------

    Dim J                                                   As Long
    Dim XAux                                                As Long
    Dim QuebraLinha                                         As Boolean
    Dim NumeroLinhasParaItem                                As Long
    Dim Texto                                               As String

    Dim X1                                                  As Double
    Dim X2                                                  As Double
    Dim Y1                                                  As Double
    Dim Y2                                                  As Double
    Dim VarLarg                                             As Long
    'Dim MargemDireita As long
    'setando o tamanho do papel para A4
50  Printer.PaperSize = 9

    'setando a escala a ser utilizada na impressão
51  Printer.ScaleMode = 1

    'setando a orientação do papel para retrato
52  If OrientacaoPapel = "Retrato" Then
53      Printer.Orientation = 1
54  Else
55      Printer.Orientation = 2
56  End If

    'setando a espessura das linhas
57  Printer.DrawWidth = 2
58  Printer.DrawStyle = 0

    'calculando o número máximo de linhas a imprimir por página
59  NumeroMaximoLinhas = Int((Printer.Height - MargemSuperior - MargemInferior - AlturaCabecalho - AlturaRodape) / dY)

    'calculando o XInicial
60  If CentralizaColunas Then
61      XInicial = Int((Printer.Width - LarguraTotal) / 2)
62  Else
63      XInicial = MargemEsquerda
64  End If
65  X = XInicial

    'verificando se haverá quebra de linha em alguma coluna
66  QuebraLinha = False
67  For J = 1 To NumeroColunas - 1
68      If QuebraLinhaColuna(J) Then
69          QuebraLinha = True
70          Exit For
71      End If
72  Next J


    '----------------------------------------------------------------
    'efetuando a impressão dos itens
    '----------------------------------------------------------------

73  NumeroPagina = 1

74  ImprimeCabecalho Titulo, LinhaAuxiliar

    'impressão de linha inicial de coluna
75  y = y - dY
76  For J = 1 To NumeroColunas - 1
77      XAux = 0
78      If J > 1 Then
79          For aa = 1 To J - 1
80              XAux = XAux + LarguraColuna(aa)
81          Next aa
82      End If
83      X1 = XInicial + XAux
84      X2 = X1 + LarguraColuna(J)
85      Y1 = y
86      Y2 = y + dY
        'impressão das linhas de contorno
87      Printer.Line (X1, Y1)-(X2, Y2), RGB(192, 192, 192), BF
88      Printer.Line (X1, Y1)-(X2, Y1)
89      Printer.Line (X1, Y2)-(X2, Y2)
90      Printer.Line (X1, Y1)-(X1, Y2)
91      Printer.Line (X2, Y1)-(X2, Y2)

        'propriedades de impressão do texto
92      Printer.FontName = "Arial"
93      Printer.FontSize = 8
94      Printer.FontBold = True
        'texto a ser impresso
95      Texto = TituloColuna(J)
96      Printer.CurrentY = y + Int((dY - Printer.TextHeight(Texto)) / 2)
97      XAux = 0
98      If J > 1 Then
99          For aa = 1 To J - 1
100             XAux = XAux + LarguraColuna(aa)
101         Next aa
102     End If
        'centralização do texto
103     Printer.CurrentX = XInicial + XAux + Int((LarguraColuna(J) - Int(Printer.TextWidth(Texto))) / 2)
        'imprimindo o texto
104     Printer.Print Texto
105 Next J
106 y = y + dY


107 For index = 1 To NumeroItens

108     NumeroLinhasParaItem = 0

        'rotina para verificar o número de linhas necessárias para impressão do item
        'caso esteja definido a opção de quebra de linha em alguma coluna
109     If QuebraLinha Then
110         For J = 1 To NumeroColunas - 1
                'texto a ser impresso
111             Texto = Grid.TextMatrix(index, TextoColuna(J))
112             VarLarg = VerificaLarguraTexto(Texto, LarguraColuna(J))
113             If VarLarg > NumeroLinhasParaItem Then NumeroLinhasParaItem = VarLarg
114         Next J
115     Else
116         NumeroLinhasParaItem = 1
117     End If

118     If NumeroLinhasParaItem + QualLinha > NumeroMaximoLinhas Then
119         QualLinha = 1
120         ImprimeRodape Rodape, NumeroPagina, AlturaRodape, Val(MargemDireita)

121         If ImprimePaginaEspecifica Then   'rotina de impressao de escolha de pagina
122             If NumPagImpIni <= NumeroPagina And NumPagImpFin >= NumeroPagina Then Printer.EndDoc Else Printer.KillDoc
123         Else
124             Printer.NewPage
125         End If  'rotina de impressao de escolha de pagina

126         ImprimeCabecalho Titulo, LinhaAuxiliar
127         y = YInicial

            'impressão de linha inicial de coluna
128         y = y - dY
129         For J = 1 To NumeroColunas - 1
130             XAux = 0
131             If J > 1 Then
132                 For aa = 1 To J - 1
133                     XAux = XAux + LarguraColuna(aa)
134                 Next aa
135             End If
136             X1 = XInicial + XAux
137             X2 = X1 + LarguraColuna(J)
138             Y1 = y
139             Y2 = y + dY
                '  impressão das linhas de contorno
140             Printer.Line (X1, Y1)-(X2, Y2), RGB(192, 192, 192), BF
141             Printer.Line (X1, Y1)-(X2, Y1)
142             Printer.Line (X1, Y2)-(X2, Y2)
143             Printer.Line (X1, Y1)-(X1, Y2)
144             Printer.Line (X2, Y1)-(X2, Y2)

                'propriedades de impressão do texto
145             Printer.FontName = "Arial"
146             Printer.FontSize = 8
147             Printer.FontBold = True
                'texto a ser impresso
148             Texto = TituloColuna(J)
149             Printer.CurrentY = y + Int((dY - Printer.TextHeight(Texto)) / 2)
150             XAux = 0
151             If J > 1 Then
152                 For aa = 1 To J - 1
153                     XAux = XAux + LarguraColuna(aa)
154                 Next aa
155             End If
                'centralização do texto
156             Printer.CurrentX = XInicial + XAux + Int((LarguraColuna(J) - Int(Printer.TextWidth(Texto))) / 2)
                'imprimindo o texto
157             Printer.Print Texto
158         Next J
159         y = y + dY

160         NumeroPagina = NumeroPagina + 1
161     Else
162         QualLinha = QualLinha + NumeroLinhasParaItem
163     End If

164     For J = 1 To NumeroColunas - 1

165         XAux = 0
166         If J > 1 Then
167             For aa = 1 To J - 1
168                 XAux = XAux + LarguraColuna(aa)
169             Next aa
170         End If
171         X1 = XInicial + XAux
172         X2 = X1 + LarguraColuna(J)
173         Y1 = y
174         Y2 = y + (NumeroLinhasParaItem * dY)

            'impressão das linhas de contorno
175         Printer.Line (X1, Y1)-(X2, Y1)
176         Printer.Line (X1, Y2)-(X2, Y2)
177         Printer.Line (X1, Y1)-(X1, Y2)
178         Printer.Line (X2, Y1)-(X2, Y2)

            'propriedades de impressão do texto
179         Printer.FontName = TipoLetraColuna(J)
180         Printer.FontSize = TamanhoLetraColuna(J)
181         Printer.FontBold = LetraNegritoColuna(J)

            'texto a ser impresso
182         Texto = Grid.TextMatrix(index, TextoColuna(J))

            'caso o texto seja mais largo que a coluna...
183         If Printer.TextWidth(Texto) > LarguraColuna(J) - 100 Then

184             NumeroPassadas = 0

185             Do While Printer.TextWidth(Texto) > LarguraColuna(J) - 100
186                 For zz = 1 To Len(Texto)
187                     If Printer.TextWidth(Left(Texto, zz)) >= LarguraColuna(J) - 100 Then
188                         Do While Mid(Texto, zz, 1) <> " "
189                             zz = zz - 1
190                             If zz < 1 Then GoTo salto
191                         Loop
192                         NumeroPassadas = NumeroPassadas + 1
193                         TextoAux = Right(Texto, Len(Texto) - zz)
194                         Texto = Left(Texto, Len(Texto) - Len(TextoAux))

195                         Printer.CurrentY = y + ((NumeroPassadas - 1) * dY) + Int((dY - Printer.TextHeight(Texto)) / 2)

196                         XAux = 0
197                         If J > 1 Then
198                             For aa = 1 To J - 1
199                                 XAux = XAux + LarguraColuna(aa)
200                             Next aa
201                         End If

                            'centralização do texto
202                         If CentralizaColuna(J) Then
203                             Printer.CurrentX = XInicial + XAux + Int((LarguraColuna(J) - Int(Printer.TextWidth(Texto))) / 2)
204                         Else
205                             Printer.CurrentX = XInicial + XAux + 100
206                         End If

                            'imprimindo o texto
207                         Printer.Print Texto

208                         Texto = TextoAux

209                         Exit For
210                     End If
211                 Next
212             Loop

salto:
213             NumeroPassadas = NumeroPassadas + 1

214             Printer.CurrentY = y + ((NumeroPassadas - 1) * dY) + Int((dY - Printer.TextHeight(Texto)) / 2)

215             XAux = 0
216             If J > 1 Then
217                 For aa = 1 To J - 1
218                     XAux = XAux + LarguraColuna(aa)
219                 Next aa
220             End If

                'centralização do texto
221             If CentralizaColuna(J) Then
222                 Printer.CurrentX = XInicial + XAux + Int((LarguraColuna(J) - Int(Printer.TextWidth(Texto))) / 2)
223             Else
224                 Printer.CurrentX = XInicial + XAux + 100
225             End If

                'imprimindo o texto
226             Printer.Print Texto

227         Else

228             Printer.CurrentY = y + Int(((dY * NumeroLinhasParaItem) - Printer.TextHeight(Texto)) / 2)

229             XAux = 0
230             If J > 1 Then
231                 For aa = 1 To J - 1
232                     XAux = XAux + LarguraColuna(aa)
233                 Next aa
234             End If

                'centralização do texto
235             If CentralizaColuna(J) Then
236                 Printer.CurrentX = XInicial + XAux + Int((LarguraColuna(J) - Int(Printer.TextWidth(Texto))) / 2)
237             Else
238                 Printer.CurrentX = XInicial + XAux + 100
239             End If

                'imprimindo o texto
240             Printer.Print Texto

241         End If


242     Next J

243     y = y + (dY * NumeroLinhasParaItem)

244 Next index

    '----------------------------------------------------------------
    '----------------------------------------------------------------

245 ImprimeRodape Rodape, NumeroPagina, AlturaRodape, Val(MargemDireita)

246 If ImprimePaginaEspecifica Then     'rotina de escolha de pagina
247     If NumPagImpIni <= NumeroPagina And NumPagImpFin >= NumeroPagina Then Printer.EndDoc Else Printer.KillDoc
248 Else
249     Printer.EndDoc
250 End If      'rotina de escolha de pagina

251 MsgBox msg & " relatório enviado para a impressora!", vbExclamation, "ATENÇÃO"

252 On Error GoTo 0
253 Exit Function

Imprimir_Error:

254 If Err.Number = 32755 Then Exit Function

255 MsgBox "Erro: " & Err.Number & " (" & Err.Description & ") na procedure Imprimir no Formulário frm_Relatorio_TipoFios", vbCritical, "ATENÇÃO"

Trata_Erro:
256 If Err.Number = 32755 Then Exit Function
257 Erros "ImprimirRelatorios"
End Function
Public Function DefineCampos(TituloCol As String, TextoCol As Long, _
                             LarguraCol As Long, Optional QuebraLinhaCol As Boolean, Optional CentralizaCol As Boolean, Optional TipoLetraCol As String, _
                             Optional TamanhoLetraCol As Long, Optional LetraNegritoCol As Boolean)

1   TituloColuna(NumeroColunas) = TituloCol
2   TextoColuna(NumeroColunas) = TextoCol
3   LarguraColuna(NumeroColunas) = LarguraCol
4   TipoLetraColuna(NumeroColunas) = IIf(TipoLetraCol = "", "Arial", TipoLetraCol)
5   CentralizaColuna(NumeroColunas) = CentralizaCol
6   TamanhoLetraColuna(NumeroColunas) = IIf(TamanhoLetraCol = 0, 7, TamanhoLetraCol)
7   LetraNegritoColuna(NumeroColunas) = LetraNegritoCol
8   QuebraLinhaColuna(NumeroColunas) = QuebraLinhaCol
9   NumeroColunas = NumeroColunas + 1
End Function

Public Sub NovoRelatorio()
1   NumeroColunas = 1
End Sub

Public Property Get Papel() As TipoDePagina
1   Papel = Papel1
End Property

Public Property Let Papel(ByVal vNewValue As TipoDePagina)
1   Papel1 = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
1   PropBag.WriteProperty "Papel", Papel1, Retrato
2   PropBag.WriteProperty "Titulo", Titulo_1, ""
3   PropBag.WriteProperty "SubTitulo", SubTitulo_1, ""
4   PropBag.WriteProperty "Rodape", Rodape_1, ""
End Sub

Public Property Get Titulo() As String
1   Titulo = Titulo_1
End Property

Public Property Let Titulo(ByVal vNewValue As String)
1   Titulo_1 = vNewValue
End Property

Public Property Get SubTitulo() As String
1   SubTitulo = SubTitulo_1
End Property

Public Property Let SubTitulo(ByVal vNewValue As String)
1   SubTitulo_1 = vNewValue
End Property
Public Property Get Rodape() As String
1   Rodape = Rodape_1
End Property

Public Property Let Rodape(ByVal vNewValue As String)
1   Rodape_1 = vNewValue
End Property


Private Function Erros(Prog As String)
1   If Err Then
2       MsgBox "Ocorreu um Erro: " & Err.Description & " n.º" & Err.Number & " Rotina :" & Prog, vbInformation, App.Title
3       Err.Number = 0
4   End If
End Function

'Public Property Get Grid() As MSHFlexGrid
'    Grid = GridImpir
'End Property

'Public Property Let Grid(ByVal vNewValue As MSHFlexGrid)
'    GridImpir = vNewValue
'End Property
