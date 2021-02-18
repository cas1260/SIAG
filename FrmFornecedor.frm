VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFornecedor 
   Caption         =   "Cadastro de Fornecedor"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   Icon            =   "FrmFornecedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   6825
   Begin TabDlg.SSTab SSTab1 
      Height          =   3315
      Left            =   30
      TabIndex        =   49
      Top             =   570
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5847
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483637
      ForeColor       =   8388608
      TabCaption(0)   =   "Cadastros (F2)"
      TabPicture(0)   =   "FrmFornecedor.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Endereço (F3)"
      TabPicture(1)   =   "FrmFornecedor.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dados Cadastrais"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   90
         TabIndex        =   45
         Top             =   1830
         Width           =   6645
         Begin VB.TextBox TxtFant 
            Height          =   285
            Left            =   780
            TabIndex        =   5
            Top             =   240
            Width           =   5775
         End
         Begin VB.TextBox TxtInscEst 
            Height          =   285
            Left            =   780
            TabIndex        =   6
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox TxtInsMun 
            Height          =   285
            Left            =   4170
            TabIndex        =   7
            Top             =   600
            Width           =   2385
         End
         Begin VB.Label Label4 
            Caption         =   "Fantasia :"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   60
            TabIndex        =   48
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Insc.Est.:"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   90
            TabIndex        =   47
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Insc.Mun.:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3330
            TabIndex        =   46
            Top             =   690
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cadastro"
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   60
         TabIndex        =   40
         Top             =   540
         Width           =   6645
         Begin VB.TextBox TxtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   870
            TabIndex        =   0
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox TxtNome 
            Height          =   285
            Left            =   3780
            TabIndex        =   4
            Top             =   600
            Width           =   2745
         End
         Begin VB.OptionButton OPF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Fisica"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2820
            TabIndex        =   1
            Top             =   300
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton OPJ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Juridica"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   4260
            TabIndex        =   2
            Top             =   270
            Width           =   1545
         End
         Begin MSMask.MaskEdBox MskCpfCnpj 
            Height          =   285
            Left            =   870
            TabIndex        =   3
            Top             =   630
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   90
            TabIndex        =   44
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Cpf/Cnpj"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   90
            TabIndex        =   43
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Razao / Nome"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2640
            TabIndex        =   42
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label33 
            Caption         =   "Tipo"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   2310
            TabIndex        =   41
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Diversos"
         ForeColor       =   &H8000000D&
         Height          =   645
         Left            =   -74940
         TabIndex        =   37
         Top             =   2610
         Width           =   6645
         Begin VB.TextBox TxtMat 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1530
            TabIndex        =   20
            Top             =   210
            Width           =   1365
         End
         Begin VB.TextBox TxtNumeroDep 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4710
            TabIndex        =   21
            Top             =   210
            Width           =   1785
         End
         Begin VB.Label Label11 
            Caption         =   "Matricula de INSS:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   90
            TabIndex        =   39
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Numero Dependentes:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2970
            TabIndex        =   38
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Endereço"
         ForeColor       =   &H8000000D&
         Height          =   2205
         Left            =   -74940
         TabIndex        =   23
         Top             =   390
         Width           =   6645
         Begin VB.TextBox txtLim 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3330
            TabIndex        =   19
            Top             =   1800
            Width           =   3165
         End
         Begin VB.TextBox TxTCont 
            Height          =   285
            Left            =   4890
            TabIndex        =   15
            Top             =   1020
            Width           =   1605
         End
         Begin VB.TextBox TxtCxp 
            Height          =   285
            Left            =   900
            TabIndex        =   13
            Top             =   1020
            Width           =   1155
         End
         Begin VB.TextBox TxTCep 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3990
            TabIndex        =   12
            Top             =   630
            Width           =   2505
         End
         Begin VB.TextBox TxtEstado 
            Height          =   285
            Left            =   2850
            MaxLength       =   2
            TabIndex        =   11
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtCidade 
            Height          =   285
            Left            =   900
            TabIndex        =   10
            Top             =   630
            Width           =   1155
         End
         Begin VB.TextBox TxtBairro 
            Height          =   285
            Left            =   3990
            TabIndex        =   9
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txtEnd 
            Height          =   285
            Left            =   900
            TabIndex        =   8
            Top             =   240
            Width           =   2385
         End
         Begin MSMask.MaskEdBox MskDataCompra 
            Height          =   285
            Left            =   900
            TabIndex        =   18
            Top             =   1800
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelefone 
            Height          =   285
            Left            =   2850
            TabIndex        =   14
            Top             =   1020
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelFax 
            Height          =   285
            Left            =   900
            TabIndex        =   16
            Top             =   1410
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFax 
            Height          =   285
            Left            =   2850
            TabIndex        =   17
            Top             =   1410
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "1a.compra"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   60
            TabIndex        =   36
            Top             =   1860
            Width           =   765
         End
         Begin VB.Label Label10 
            Caption         =   "Lim.Credito:"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2400
            TabIndex        =   35
            Top             =   1860
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "Cont"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4350
            TabIndex        =   34
            Top             =   1110
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2130
            TabIndex        =   33
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label Label16 
            Caption         =   "Cx. P"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Cep"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3420
            TabIndex        =   31
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Estado"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   2190
            TabIndex        =   30
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   150
            TabIndex        =   29
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Bairro"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3390
            TabIndex        =   28
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Endereço"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   27
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Tel. Fax"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   26
            Top             =   1470
            Width           =   675
         End
         Begin VB.Label Label20 
            Caption         =   "Fax"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2430
            TabIndex        =   25
            Top             =   1470
            Width           =   435
         End
         Begin VB.Label Label21 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4290
            TabIndex        =   24
            Top             =   1410
            Width           =   2205
         End
      End
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   953
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Novo"
            Key             =   "A"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Abrir"
            Key             =   "B"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "C"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "D"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Localizar"
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "F"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "G"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4320
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
            Picture         =   "FrmFornecedor.frx":047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":058E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":06A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFornecedor.frx":0AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsFor                                               As Recordset
Private Sub Form_Load()
1   Me.Width = 6945
2   Me.Height = 4320
3   Centra Me
End Sub
Private Sub OPF_Click()
1   On Error GoTo Trata_Erro
2   MskCpfCnpj.Mask = ""
3   MskCpfCnpj.Text = ""
4   MskCpfCnpj.Mask = "###.###.###-##"
Trata_Erro:
5   E
End Sub

Private Sub OPJ_Click()
1   On Error GoTo Trata_Erro
2   MskCpfCnpj.Mask = ""
3   MskCpfCnpj.Text = ""
4   MskCpfCnpj.Mask = "########/####-##"
Trata_Erro:
5   E
End Sub
Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   Select Case Button.Key
    Case "A"
2       Novo
3       txtCodigo.SetFocus
4   Case "B"
5       Abrir
6   Case "C"
7       Salvar
8       txtCodigo.SetFocus
9   Case "D"
10      Excluir

11  Case "E"
12      Localizar
13  Case "G"
14      Unload Me
15  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       SendKeys "{TAB}"
3   ElseIf KeyCode = 27 Then
4       Unload Me
5   ElseIf KeyCode = vbKeyF2 Then
6       SSTab1.Tab = 0
7   ElseIf KeyCode = vbKeyF3 Then
8       SSTab1.Tab = 1
9   End If
End Sub
Private Sub MskCpfCnpj_KeyDown(KeyCode As Integer, Shift As Integer)
1   On Error GoTo Trata_Erro
2   If KeyCode = 13 Then
3       If OPF.Value = True Then
4           If Calc_CPF(MskCpfCnpj.ClipText) = False Then
5               Resp 30, ""
6               MskCpfCnpj.SetFocus
7           Else
8               If Trim(MskCpfCnpj.ClipText) = "" Then
9                   Resp 30, ""
10                  MskCpfCnpj.SetFocus
11              End If
12          End If
13      Else
14          If Calc_CGC(MskCpfCnpj.ClipText) = False Then
15              Resp 29, ""
16              MskCpfCnpj.SetFocus
17          Else
18              If Trim(MskCpfCnpj.ClipText) = "" Then
19                  Resp 29, ""
20                  MskCpfCnpj.SetFocus
21              End If
22          End If
23      End If
24  End If
Trata_Erro:
25  E
End Sub
Private Sub MskDataCompra_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskDataCompra.Text = Valida(MskDataCompra)
3   End If
End Sub

Private Sub TxtCodigo_GotFocus()
1   SSTab1.Tab = 0
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       Abrir
3   End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub txtEnd_GotFocus()
1   SSTab1.Tab = 1
End Sub

Private Sub txtLim_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtNome_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       If Trim(TxtNome.Text) = "" Then
3           MsgBox "Entrada inconsistente", vbCritical, App.Title
4           TxtNome.SetFocus
5       End If
6   End If
End Sub

Private Sub Novo()
1   On Error GoTo Trata_Erro
2   TxtEstado.Text = ""
3   txtCodigo.Text = ""
4   OPF.Value = True
5   OPJ.Value = False
6   MskCpfCnpj.Mask = ""
7   MskCpfCnpj.Text = ""
8   MskCpfCnpj.Mask = "###.###.###-##"
    'MskCpfCnpj.Text = ""
9   TxtNome.Text = ""
10  TxtFant.Text = ""
11  TxtInscEst.Text = ""
12  TxtInsMun.Text = ""
13  MskDataCompra.Text = "__/__/____"
14  txtLim.Text = ""
15  TxtEnd.Text = ""
16  TxtBairro.Text = ""
17  TxtCidade.Text = ""
18  TxtCep.Text = ""
19  TxtCxp.Text = ""
20  MskTelefone.Mask = ""
21  MskTelefone.Text = ""
22  MskTelefone.Mask = "(##) ####-####"
23  TxTCont.Text = ""
24  MskTelFax.Mask = ""
25  MskTelFax.Text = ""
26  MskTelFax.Mask = "(##) ####-####"

27  MskFax.Mask = ""
28  MskFax.Text = ""
29  MskFax.Mask = "(##) ####-####"
30  TxtMat.Text = ""
31  TxtNumeroDep.Text = ""
32  SSTab1.Tab = 0
33  txtCodigo.SetFocus
Trata_Erro:
34  E
End Sub

Private Sub Abrir()
1   On Error GoTo Trata_Erro
    Dim VCodigo                                             As String
    Dim Tipo                                                As Long
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Fornecedor", vbCritical, App.Title
4       Exit Sub
5   End If

6   Comando = "Select * from Fornecedor Where Codigo =" & txtCodigo.Text & ""
7   Set RsFor = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

8   If RsFor.RecordCount = 0 Then
9       VCodigo = txtCodigo.Text
10      Novo
11      txtCodigo.Text = VCodigo
12  Else
13      txtCodigo.Text = RsFor!Codigo
14      If UCase(RsFor!pessoa) = "F" Then
15          OPF.Value = True
16          OPJ.Value = False
17      Else
18          OPF.Value = False
19          OPJ.Value = True
20      End If
21      MskCpfCnpj.Text = RsFor!CNPJ
22      TxtNome.Text = RsFor!Razao
23      TxtFant.Text = RsFor!Fantasia
24      TxtInscEst.Text = RsFor!Estadual
25      TxtInsMun.Text = RsFor!Municipal
26      MskDataCompra.Text = RsFor!PCompra
27      txtLim.Text = RsFor!Limite
28      MskTelFax.Text = RsFor!TelFax
29      MskFax.Text = RsFor!Fax
30      TxtMat.Text = RsFor!Mat
31      TxtNumeroDep.Text = RsFor!Mat
32      MskCpfCnpj.Text = RsFor!CNPJ
33      TxtEnd.Text = RsFor!Endereco
34      TxtBairro.Text = RsFor!Bairro
35      TxtCidade.Text = RsFor!Cidade
36      TxtCep.Text = RsFor!Cep
37      TxtCxp.Text = RsFor!Caixa
38      MskTelefone.Text = RsFor!Telefone
39      TxTCont.Text = RsFor!Cont
40      MskTelFax.Text = RsFor!TelFax
41      MskFax.Text = RsFor!Fax
42      TxtMat.Text = RsFor!Mat
43      TxtNumeroDep.Text = RsFor!Dep
44      TxtEstado.Text = RsFor!Estado
45  End If
46  RsFor.Close
Trata_Erro:
47  E
End Sub
Private Sub Salvar()
1   On Error GoTo Trata_Erro
    Dim VCodigo                                             As String
    Dim Tipo                                                As Long
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Fornecedor"
4       Exit Sub
5   End If
6   MskDataCompra.Text = Valida(MskDataCompra)
7   Comando = "Select * from Fornecedor Where Codigo =" & txtCodigo.Text & ""
8   Set RsFor = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

9   If RsFor.RecordCount = 0 Then
10      RsFor.AddNew
11  Else
12      RsFor.Edit
13  End If

14  RsFor!Codigo = txtCodigo.Text
15  If OPF.Value = True Then
16      RsFor!pessoa = "F"
17  Else
18      RsFor!pessoa = "J"
19  End If
20  RsFor!Estado = TxtEstado.Text
21  RsFor!CNPJ = MskCpfCnpj.Text
22  RsFor!Razao = TxtNome.Text
23  RsFor!Fantasia = TxtFant.Text
24  RsFor!Estadual = TxtInscEst.Text
25  RsFor!Municipal = TxtInsMun.Text
26  RsFor!PCompra = MskDataCompra.Text
27  RsFor!Limite = txtLim.Text
28  RsFor!Endereco = TxtEnd.Text
29  RsFor!Bairro = TxtBairro.Text
30  RsFor!Cidade = TxtCidade.Text
31  RsFor!Cep = TxtCep.Text
32  RsFor!Caixa = TxtCxp.Text
33  RsFor!Telefone = MskTelefone.Text
34  RsFor!Cont = TxTCont.Text
35  RsFor!TelFax = MskTelFax.Text
36  RsFor!Fax = MskFax.Text
37  RsFor!Mat = TxtMat.Text
38  RsFor!Dep = TxtNumeroDep.Text
39  RsFor.Update
40  RsFor.Close
41  Novo
42  SSTab1.Tab = 0
43  SSTab1.SetFocus
    'TxtCodigo.SetFocus
Trata_Erro:
44  E
End Sub

Public Sub Excluir()
1   On Error GoTo Trata_Erro
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Fornecedor", vbCritical, App.Title
4       Exit Sub
5   End If

6   Comando = "Select * from Fornecedor Where Codigo =" & txtCodigo.Text & ""
7   Set RsFor = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

8   If RsFor.RecordCount <> 0 Then
9       If MsgBox("Confirma Exclusão ?", vbCritical + vbYesNo + vbDefaultButton2 + vbSystemModal, App.Title) = vbYes Then
10          RsFor.Delete
11          RsFor.Close
12          MsgBox "Fornecedor Excluido com Sucesso ! ! !"
13          Novo
14      End If
15  Else
16      MsgBox "Impossivel Excluir"
17  End If
Trata_Erro:
18  E
End Sub
Private Sub Localizar()
1   On Error GoTo Trata_Erro
2   ShowFornecedor = 0
3   FrmPesqFornecedor.Show 1
4   If ShowFornecedor <> 0 Then
5       txtCodigo.Text = ShowFornecedor
6       Abrir
7   End If
Trata_Erro:
8   E
End Sub
Private Sub TxtNumeroDep_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Salvar
End Sub
