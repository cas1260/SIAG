VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCheques 
   Caption         =   "Cheques"
   ClientHeight    =   6060
   ClientLeft      =   1620
   ClientTop       =   1815
   ClientWidth     =   9750
   Icon            =   "FrmCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab Tab 
      Height          =   5415
      Left            =   60
      TabIndex        =   10
      Top             =   600
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados do Cheques"
      TabPicture(0)   =   "FrmCheques.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "kgsd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MskVencimento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "MskData"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CboNome"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtNumero"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtBanco"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtAgencia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtConta"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtValor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtObs"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Consulta de Cheques"
      TabPicture(1)   =   "FrmCheques.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdPesq"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Grid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MskIni"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MskFinal"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LblTotal"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Labelddd"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton CmdPesq 
         Caption         =   "Pesquisar"
         Height          =   315
         Left            =   -71370
         TabIndex        =   25
         Top             =   720
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3945
         Left            =   -74910
         TabIndex        =   24
         Top             =   1080
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"FrmCheques.frx":0342
      End
      Begin MSMask.MaskEdBox MskIni 
         Height          =   285
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtObs 
         Height          =   3255
         Left            =   120
         TabIndex        =   8
         Top             =   2070
         Width           =   9465
      End
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   6510
         TabIndex        =   6
         Top             =   1470
         Width           =   1755
      End
      Begin VB.TextBox TxtConta 
         Height          =   285
         Left            =   4680
         TabIndex        =   5
         Top             =   1470
         Width           =   1785
      End
      Begin VB.TextBox TxtAgencia 
         Height          =   285
         Left            =   2850
         TabIndex        =   4
         Top             =   1470
         Width           =   1785
      End
      Begin VB.TextBox TxtBanco 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1470
         Width           =   1485
      End
      Begin VB.TextBox TxtNumero 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   930
         Width           =   1605
      End
      Begin VB.ComboBox CboNome 
         Height          =   315
         Left            =   1770
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   900
         Width           =   7755
      End
      Begin MSMask.MaskEdBox MskData 
         Height          =   270
         Left            =   150
         TabIndex        =   2
         Top             =   1470
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskVencimento 
         Height          =   300
         Left            =   8310
         TabIndex        =   7
         Top             =   1470
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFinal 
         Height          =   285
         Left            =   -73110
         TabIndex        =   23
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -67200
         TabIndex        =   26
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "Data de Vencimento Final"
         Height          =   225
         Left            =   -73260
         TabIndex        =   22
         Top             =   510
         Width           =   1965
      End
      Begin VB.Label Labelddd 
         Caption         =   "Data de Vencimento Inicial"
         Height          =   225
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Obs.:"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Vencimento"
         Height          =   195
         Left            =   8340
         TabIndex        =   18
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Valor"
         Height          =   195
         Left            =   6510
         TabIndex        =   17
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Conta"
         Height          =   195
         Left            =   4680
         TabIndex        =   16
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Agência"
         Height          =   195
         Left            =   2850
         TabIndex        =   15
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Banco"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Numero do Cheque"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Cliente"
         Height          =   195
         Left            =   1770
         TabIndex        =   12
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label kgsd 
         Caption         =   "Data"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1290
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":03F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":050A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":061E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0732
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0846
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCheques.frx":0A6E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   953
      ButtonWidth     =   1005
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "C"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "D"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "G"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdPesq_Click()
1   MskIni.Text = Valida(MskIni)
2   MskFinal.Text = Valida(MskFinal)

    Dim Rs                                                  As Recordset
3   Sql = "Select * from Cheques Where Vencimento >=#" & Format(MskIni.Text, "MM/DD/YYYY") & "# And Vencimento <= #" & Format(MskFinal.Text, "MM/DD/YYYY") & "# Order By Vencimento"

4   Grid.Clear
5   Grid.Rows = 2
6   Grid.FormatString = "Vencimento        |Numero            |Cliente                                                                                                    |Valor                         "

7   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
8   LblTotal.Caption = "0"
9   If Rs.EOF = False Then
10      Rs.MoveLast
11      Rs.MoveFirst
12      Grid.Rows = Rs.RecordCount + 1
13      X = 1
14      LblTotal.Caption = "0"
15      Do While Not Rs.EOF
16          Grid.TextMatrix(X, 0) = Rs!Vencimento
17          Grid.TextMatrix(X, 1) = Rs!Numero
18          Grid.TextMatrix(X, 2) = Rs!Cliente
19          Grid.TextMatrix(X, 3) = Format(Rs!Valor, "###,###,##0.00")
20          LblTotal.Caption = CCur(LblTotal.Caption) + Rs!Valor
21          Rs.MoveNext
22          X = X + 1
23      Loop
24      Rs.Close
25  End If
26  LblTotal.Caption = Format(LblTotal.Caption, "###,###,##0.00")
End Sub

Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   Select Case Button.index
    Case 1
2       Salvar
3   Case 2
4       Excluir
5   Case 4
6       Unload Me
7   End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{tab}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   Sql = "Select * From Cliente order By Razao"
3   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)

4   CboNome.Clear
5   If Rs.EOF = False Then
6       Do While Not Rs.EOF
7           CboNome.AddItem Rs!Razao
8           Rs.MoveNext
9       Loop
10  End If

11  Rs.Close

Trata_Erro:
12  E
End Sub

Public Sub BuscaCheque()
1   On Error GoTo Trata_Erro
    Dim Rs                                                  As Recordset

2   If Trim(TxtNumero.Text) = "" Then
3       MsgBox "Numero do Cheque invalido!", vbCritical, "Atenção"
4       TxtNumero.SetFocus
5       Exit Sub
6   End If

7   Sql = "Select * From Cheques Where Numero = " & TxtNumero.Text
8   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenSnapshot)
9   If Rs.EOF = False Then
10      TxtAgencia.Text = Rs!Agencia
11      TxtBanco.Text = Rs!Banco
12      TxtConta.Text = Rs!Conta
13      TxtNumero.Text = Rs!Numero
14      TxtObs.Text = Rs!Obs
15      TxtValor.Text = Rs!Valor
16      MskData.Text = Rs!Data
17      MskVencimento.Text = Rs!Vencimento
18      CboNome.Text = Rs!Cliente
        'MskIni.Text = "__/__/____"
        'MskFinal.Text = "__/__/____"
19  Else
20      TxtAgencia.Text = ""
21      TxtBanco.Text = ""
22      TxtConta.Text = ""
        'TxtNumero.Text = ""
23      TxtObs.Text = ""
24      TxtValor.Text = ""
25      MskData.Text = "__/__/____"
26      MskVencimento.Text = "__/__/____"
27      CboNome.Text = ""
        'MskIni.Text = "__/__/____"
        'MskFinal.Text = "__/__/____"
28  End If



Trata_Erro:
29  E
End Sub

Private Sub Salvar()
1   On Error GoTo Trata_Erro

2   If Trim(TxtNumero.Text) = "" Then
3       MsgBox "Numero do Cheque invalido!", vbCritical, "Atenção"
4       TxtNumero.SetFocus
5       Exit Sub
6   End If

7   If Trim(CboNome.Text) = "" Then
8       MsgBox "Nome do cliente invalido!", vbCritical, "Atenção"
9       CboNome.SetFocus
10      Exit Sub
11  End If

12  If Trim(MskData.ClipText) = "" Then
13      MsgBox "Data invalida!", vbCritical, "Atenção"
14      MskData.SetFocus
15      Exit Sub
16  End If

17  If Trim(TxtBanco.Text) = "" Then
18      MsgBox "Banco invalido!", vbCritical, App.Title
19      TxtBanco.SetFocus
20      Exit Sub
21  End If

22  If Trim(TxtConta.Text) = "" Then
23      MsgBox "Conta Invalida!", vbCritical, "Atenção"
24      TxtConta.SetFocus
25      Exit Sub
26  End If

27  If Trim(TxtAgencia.Text) = "" Then
28      MsgBox "Agência invalida!", vbCritical, "Atenção"
29      TxtAgencia.SetFocus
30      Exit Sub
31  End If

32  If Trim(MskVencimento.ClipText) = "" Then
33      MsgBox "Vencimento invalido!", vbCritical, "Atenção"
34      MskVencimento.SetFocus
35      Exit Sub
36  End If

37  If Trim(TxtValor.Text) = "" Then
38      MsgBox "Valor invalido!", vbCritical, "Atenção"
39      TxtValor.SetFocus
40      Exit Sub
41  End If

42  Sql = "Select * From Cheques Where Numero = " & TxtNumero.Text
    Dim Rs                                                  As Recordset
43  Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)

44  If Rs.EOF = False Then
45      Rs.Edit
46  Else
47      Rs.AddNew
48  End If

49  Rs!Agencia = TxtAgencia.Text
50  Rs!Banco = TxtBanco.Text
51  Rs!Conta = TxtConta.Text
52  Rs!Numero = TxtNumero.Text
53  Rs!Obs = IIf(Trim(TxtObs.Text) = "", " ", TxtObs.Text)
54  Rs!Valor = TxtValor.Text
55  Rs!Data = MskData.Text
56  Rs!Vencimento = MskVencimento.Text
57  Rs!Cliente = CboNome.Text
58  Rs.Update
59  Rs.Close
60  MsgBox "Operação realizada com sucesso!", vbInformation, "Ok"

61  TxtAgencia.Text = ""
62  TxtBanco.Text = ""
63  TxtConta.Text = ""
64  TxtObs.Text = ""
65  TxtValor.Text = ""
66  MskData.Text = "__/__/____"
67  MskVencimento.Text = "__/__/____"
68  TxtNumero.Text = ""
69  TxtNumero.SetFocus

70  CboNome.Text = ""

Trata_Erro:
71  E
End Sub


Private Sub Excluir()
1   On Error GoTo Trata_Erro

2   If Trim(TxtNumero.Text) = "" Then
3       MsgBox "numero invalido!", vbCritical, "Atenção"
4       TxtNumero.SetFocus
5       Exit Sub
6   End If

7   Sql = "Select * From Cheques Where Numero = " & TxtNumero.Text
    Dim Rs                                                  As Recordset

8   Set Rs = BancoDeDados.OpenRecordset(Sql, dbOpenDynaset)
9   If Rs.EOF = False Then
10      If MsgBox("Confirma a exclusão deste cheque?", vbYesNo + vbQuestion + vbDefaultButton2, "Atenção") = vbNo Then Exit Sub
11      Rs.Delete
12  Else
13      MsgBox "Impossivel excluir este cheque!", vbCritical, "Atenção"
14  End If
15  TxtAgencia.Text = ""
16  TxtBanco.Text = ""
17  TxtConta.Text = ""
18  TxtObs.Text = ""
19  TxtValor.Text = ""
20  MskData.Text = "__/__/____"
21  MskVencimento.Text = "__/__/____"
22  TxtNumero.Text = ""
23  CboNome.Text = ""
24  TxtNumero.SetFocus

25  Rs.Close

Trata_Erro:
26  E
End Sub

Private Sub MskData_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskData.Text = Valida(MskData)
3   End If
End Sub

Private Sub MskFinal_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then MskFinal.Text = Valida(MskFinal)
End Sub

Private Sub MskIni_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then MskIni.Text = Valida(MskIni)
End Sub

Private Sub MskVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskVencimento.Text = Valida(MskVencimento)
3   End If
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then BuscaCheque
End Sub
