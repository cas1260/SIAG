VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmVendedor 
   Caption         =   "Vendedores"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "FrmVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   7245
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Empresa"
      ForeColor       =   &H80000002&
      Height          =   2325
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   7215
      Begin VB.TextBox TxtDoc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5940
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1080
         Width           =   1185
      End
      Begin VB.TextBox TxtCom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6090
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox TxtCom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox TxtCom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3990
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1680
         Width           =   1035
      End
      Begin MSMask.MaskEdBox MskTel 
         Height          =   315
         Left            =   2370
         TabIndex        =   8
         Top             =   1680
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         Mask            =   "(##) #### - ####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   50
         TabIndex        =   0
         Top             =   510
         Width           =   915
      End
      Begin VB.TextBox TxtRazao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   510
         Width           =   6075
      End
      Begin VB.TextBox TxtEnd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Txtbairro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2625
         TabIndex        =   3
         Top             =   1080
         Width           =   1605
      End
      Begin VB.TextBox TxtCidade 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4230
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtEstado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1680
         Width           =   945
      End
      Begin VB.TextBox TxtCEp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Doc. (s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   5940
         TabIndex        =   25
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   6090
         TabIndex        =   24
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gerente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   5040
         TabIndex        =   23
         Top             =   1470
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4020
         TabIndex        =   22
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1080
         TabIndex        =   19
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   870
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2640
         TabIndex        =   17
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4260
         TabIndex        =   16
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1050
         TabIndex        =   14
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2370
         TabIndex        =   13
         Top             =   1470
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   690
      Top             =   810
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
            Picture         =   "FrmVendedor.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVendedor.frx":0ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
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
End
Attribute VB_Name = "FrmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsVendedor                                          As Recordset

Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   Select Case UCase(Button.Key)
    Case "C"
2       Salvar
3   Case "D"
4       Excluir
5   Case "E"
6       ShowVendedor = "0"
7       FrmLocalizarVendedor.Show 1
8       If ShowVendedor <> "0" Then
9           txtCodigo.Text = ShowVendedor
10          Abrir
11      End If
12  Case "G"
13      Unload Me
14  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{TAB}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   Me.Height = 3360
2   Me.Width = 7365
3   Centra Me
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Abrir
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtCom_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       If index = 1 Then
3           If Trim(TxtCom(1).Text) = "" Then
4               Salvar
5               SendKeys "{TAB}"
6           End If
7       End If
8       If index = 2 Then
9           Salvar
10      End If
11  End If
End Sub

Private Sub TxtCom_KeyPress(index As Integer, KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub Novo()
'On Error GoTo Trata_Erro
1   On Error Resume Next

2   txtCodigo.Text = ""
3   TxtRazao.Text = ""
4   TxtEnd.Text = ""
5   TxtBairro.Text = ""
6   TxtCidade.Text = ""
7   TxtEstado.Text = ""
8   TxtCep.Text = ""
9   MskTel.Text = "(__) ____ - ____"
10  TxtCom(0).Text = ""
11  TxtCom(1).Text = ""
12  TxtCom(2).Text = ""
13  TxtDoc.Text = ""
14  E
End Sub

Private Sub Abrir()
1   On Error Resume Next
    Dim CodigoVend                                          As String
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Vendedor com Codigo Invalido", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If

7   Comando = "Select * From Vendedor Where Codigo = " & txtCodigo & ""
8   Set RsVendedor = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

9   If RsVendedor.RecordCount = 0 Then
10      CodigoVend = txtCodigo.Text
11      Novo
12      txtCodigo.Text = CodigoVend
13  Else
14      With RsVendedor
15          TxtRazao.Text = !Razao
16          TxtEnd.Text = !Endereco
17          TxtBairro.Text = !Bairro
18          TxtCidade.Text = !Cidade
19          TxtEstado.Text = !Estado
20          TxtCep.Text = !Cep
21          MskTel.Text = !Telefone
22          TxtCom(0).Text = !Comissao
23          TxtCom(1).Text = !Gerente
24          TxtCom(2).Text = !ComissaoGer
25          TxtDoc.Text = !Doc
26      End With
27  End If
28  RsVendedor.Close
29  E
End Sub

Private Sub Salvar()
1   On Error Resume Next
    Dim CodigoVend                                          As String
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Vendedor com Codigo Invalido", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If

7   Comando = "Select * From Vendedor Where Codigo = " & txtCodigo & ""
8   Set RsVendedor = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
9   If Trim(TxtCom(0).Text) = "" Then
10      TxtCom(0).Text = 0
11  End If
12  If RsVendedor.RecordCount = 0 Then
13      RsVendedor.AddNew
14  Else
15      RsVendedor.Edit
16  End If
17  With RsVendedor
18      !Codigo = txtCodigo.Text
19      !Razao = TxtRazao.Text
20      !Endereco = TxtEnd.Text
21      !Bairro = TxtBairro.Text
22      !Cidade = TxtCidade.Text
23      !Estado = TxtEstado.Text
24      !Cep = TxtCep.Text
25      !Telefone = MskTel.Text
26      !Comissao = TxtCom(0).Text
27      !Gerente = TxtCom(1).Text
28      !ComissaoGer = TxtCom(2).Text
29      !Doc = TxtDoc.Text
30      .Update
31  End With
32  RsVendedor.Close
33  Novo
    'TxtCodigo.SetFocus
34  E
End Sub

Private Sub Excluir()
1   On Error Resume Next
    Dim CodigoVend                                          As String
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Vendedor com Codigo Invalido", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If

7   Comando = "Select * From Vendedor Where Codigo = " & txtCodigo & ""
8   Set RsVendedor = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

9   If RsVendedor.RecordCount = 0 Then
10      MsgBox "Não a vendedor com este codigo", vbCritical, App.Title
11      Exit Sub
12  Else
13      If MsgBox("Deseja Excluir este Vendedor ?", vbCritical + vbYesNo + vbDefaultButton2 + vbSystemModal, App.Title) = vbYes Then
14          RsVendedor.Delete
15          Novo
16          MsgBox "Vendedor excluido com Sucesso !", vbInformation, App.Title
17          txtCodigo.SetFocus
18      End If
19      RsVendedor.Close
20  End If
21  E
End Sub


