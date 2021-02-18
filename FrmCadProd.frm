VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCadProd 
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "FrmCadProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8145
   Begin MSComDlg.CommonDialog Com 
      Left            =   3210
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtICMS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4380
      MaxLength       =   6
      TabIndex        =   30
      Top             =   4770
      Width           =   1125
   End
   Begin VB.TextBox TxtComissao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6870
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4770
      Width           =   1185
   End
   Begin VB.TextBox TxtSit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   14
      Top             =   3120
      Width           =   1275
   End
   Begin VB.TextBox TxtClassFiscal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4380
      MaxLength       =   6
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TxtIpi 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6870
      MaxLength       =   6
      TabIndex        =   16
      Top             =   3120
      Width           =   1185
   End
   Begin VB.TextBox TxtFabricante 
      Height          =   285
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3450
      Width           =   1300
   End
   Begin VB.TextBox txtEmbalagem 
      Height          =   285
      Left            =   4380
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3450
      Width           =   1575
   End
   Begin VB.TextBox TxtMarca 
      Height          =   285
      Left            =   6870
      MaxLength       =   20
      TabIndex        =   19
      Top             =   3450
      Width           =   1185
   End
   Begin VB.TextBox TxtModelo 
      Height          =   285
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   20
      Top             =   3780
      Width           =   1300
   End
   Begin VB.TextBox TxtNumero 
      Height          =   285
      Left            =   4380
      MaxLength       =   50
      TabIndex        =   21
      Top             =   3780
      Width           =   1575
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6870
      MaxLength       =   6
      TabIndex        =   22
      Top             =   3780
      Width           =   1185
   End
   Begin VB.TextBox TxtEspecie 
      Height          =   285
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   23
      Top             =   4110
      Width           =   1300
   End
   Begin VB.TextBox TxtBruto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6870
      MaxLength       =   6
      TabIndex        =   28
      Top             =   4440
      Width           =   1185
   End
   Begin VB.TextBox TxtLiq 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   29
      Top             =   4770
      Width           =   1095
   End
   Begin VB.TextBox TxtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   810
      MaxLength       =   6
      TabIndex        =   0
      Top             =   690
      Width           =   975
   End
   Begin VB.TextBox TxtClas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2790
      MaxLength       =   15
      TabIndex        =   1
      Top             =   690
      Width           =   1425
   End
   Begin VB.TextBox TxtMedida 
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   2
      Top             =   690
      Width           =   1245
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   810
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Width           =   5475
   End
   Begin VB.TextBox TxtAtual 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   810
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1410
      Width           =   1185
   End
   Begin VB.TextBox Txtmax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3030
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1410
      Width           =   1305
   End
   Begin VB.TextBox TxtCompra 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   810
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1860
      Width           =   1155
   End
   Begin VB.TextBox TxtMedio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5190
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1410
      Width           =   1095
   End
   Begin VB.TextBox TxtVenda 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3030
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1860
      Width           =   1305
   End
   Begin VB.TextBox TxtPrazo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1860
      Width           =   1125
   End
   Begin VB.ComboBox CboGrupo 
      Height          =   315
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.ComboBox CboFornecedor 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   2355
   End
   Begin VB.ComboBox CboLocal 
      Height          =   315
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2700
      Width           =   2205
   End
   Begin VB.ComboBox CboCusto 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2670
      Width           =   2355
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   -120
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
            Picture         =   "FrmCadProd.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadProd.frx":0ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
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
   Begin MSMask.MaskEdBox MskEntrada 
      Height          =   285
      Left            =   4380
      TabIndex        =   24
      Top             =   4110
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskSaida 
      Height          =   285
      Left            =   6870
      TabIndex        =   25
      Top             =   4110
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFabricacao 
      Height          =   285
      Left            =   1410
      TabIndex        =   26
      Top             =   4440
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskValidade 
      Height          =   285
      Left            =   4380
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Image I 
      Height          =   2160
      Left            =   90
      Picture         =   "FrmCadProd.frx":0F0E
      Top             =   3930
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Prod 
      BorderStyle     =   1  'Fixed Single
      Height          =   2235
      Left            =   6390
      Stretch         =   -1  'True
      Top             =   690
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ICMS"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   20
      Left            =   3810
      TabIndex        =   64
      Top             =   4860
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comissão"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   19
      Left            =   6090
      TabIndex        =   63
      Top             =   4890
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Situação Tibutaria"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   11
      Left            =   30
      TabIndex        =   62
      Top             =   3210
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Classificação Fiscal"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   12
      Left            =   2940
      TabIndex        =   61
      Top             =   3210
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ipi %"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   13
      Left            =   6150
      TabIndex        =   60
      Top             =   3180
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   570
      TabIndex        =   59
      Top             =   3510
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Embalagem"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3390
      TabIndex        =   58
      Top             =   3540
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   14
      Left            =   6120
      TabIndex        =   57
      Top             =   3540
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Modelo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   780
      TabIndex        =   56
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Numero"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3660
      TabIndex        =   55
      Top             =   3900
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serie"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   15
      Left            =   6150
      TabIndex        =   54
      Top             =   3840
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Especie"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   750
      TabIndex        =   53
      Top             =   4170
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Ultima Entrada"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3180
      TabIndex        =   52
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label Label11 
      Caption         =   "Ultima Saida"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   6090
      TabIndex        =   51
      Top             =   4080
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Fabricação"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   390
      TabIndex        =   50
      Top             =   4530
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Validade"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3600
      TabIndex        =   49
      Top             =   4530
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Peso Bruto"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   16
      Left            =   6000
      TabIndex        =   48
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Peso Liquido"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   17
      Left            =   360
      TabIndex        =   47
      Top             =   4860
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "U. Medida"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4260
      TabIndex        =   46
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C. Medio"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   4
      Left            =   4470
      TabIndex        =   45
      Top             =   1500
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Preço a Prazo"
      ForeColor       =   &H8000000D&
      Height          =   405
      Index           =   6
      Left            =   4470
      TabIndex        =   44
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   43
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Classificação"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1830
      TabIndex        =   42
      Top             =   780
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   41
      Top             =   1140
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Est.Atual"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   1470
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Est. Minimo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   2130
      TabIndex        =   39
      Top             =   1470
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Preço de Compra"
      ForeColor       =   &H8000000D&
      Height          =   435
      Index           =   3
      Left            =   60
      TabIndex        =   38
      Top             =   1770
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Preço de Venda"
      ForeColor       =   &H8000000D&
      Height          =   405
      Index           =   5
      Left            =   2190
      TabIndex        =   37
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   7
      Left            =   270
      TabIndex        =   36
      Top             =   2400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   8
      Left            =   3150
      TabIndex        =   35
      Top             =   2370
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   9
      Left            =   300
      TabIndex        =   34
      Top             =   2820
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C. de Custo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   10
      Left            =   3090
      TabIndex        =   33
      Top             =   2790
      Width           =   825
   End
End
Attribute VB_Name = "FrmCadProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp                                                  As Recordset
Dim NomeFig                                                 As String
Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "G"
        Unload Me
    Case "C"
        Salvar
        txtCodigo.SetFocus
    Case "D"
        Excluir
    Case "E"
        ShowProduto = ""
        FrmPesqProd.Show 1
        If Trim(ShowProduto) <> "" Then
            txtCodigo.Text = ShowProduto
            TxtClas.Text = ""
            Abrir False
        End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Fornecedor
    Grupo
    Centro
    'LocalObra
    Prod.Picture = i.Picture
    'Novo False
    Me.Height = 5595
    Me.Width = 8265
    Centra Me
End Sub

Private Sub Prod_DblClick()
    On Error Resume Next
    Com.DialogTitle = "Todos os arquivo de Figuras"
    Com.Filter = "*.Gif/*.Bmp/*.Jpg/*.Jpeg|*.Gif;*.Bmp;*.Jpg;*.Jpeg|Todos os Arquivos |*.*"
    Com.ShowOpen
    If Trim(Com.FileName) <> "" Then
        If Dir(Com.FileName) <> "" Then
            Prod.Picture = LoadPicture(Com.FileName)
        End If
    End If
End Sub

Private Sub TxtAtual_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtClas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Abrir True
End Sub

Private Sub TxtClas_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtClassFiscal_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtCodigo.Text) <> "" Then
            Abrir False
        End If
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub
Private Sub Novo(Tipo As Boolean)
    On Error Resume Next
    TxtAtual.Text = 0
    TxtBruto.Text = 0
    If Tipo = False Then
        txtCodigo.Text = ""
        TxtClas.Text = ""
    End If
    TxtClassFiscal.Text = 0
    TxtCompra.Text = 0
    TxtDescricao.Text = ""
    txtEmbalagem.Text = ""
    TxtEspecie.Text = ""
    TxtFabricante.Text = ""
    TxtIpi.Text = 0
    TxtLiq.Text = 0
    TxtMarca.Text = ""
    Txtmax.Text = 0
    TxtMedida.Text = ""
    TxtMedio.Text = 0
    TxtModelo.Text = ""
    TxtNumero.Text = ""
    TxtPrazo.Text = 0
    TxtSerie.Text = 0
    TxtSit.Text = ""
    TxtVenda.Text = 0
    mskFabricacao.Text = "__/__/____"
    MskSaida.Text = "__/__/____"
    MskEntrada.Text = "__/__/____"
    MskValidade.Text = "__/__/____"
    CboCusto.ListIndex = -1
    CboFornecedor.ListIndex = -1
    CboGrupo.ListIndex = -1
    CboLocal.ListIndex = -1
    Prod.Picture = i.Picture
End Sub

Private Sub TxtComissao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Salvar
    End If
End Sub

Private Sub TxtComissao_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtCompra_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(TxtDescricao.Text) = "" Then
            MsgBox "A Descrição e Obrigatorio", vbInformation, App.Title
            TxtDescricao.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub TxtICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtIpi_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub Txtmax_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub
Private Sub TxtMedio_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub
Private Sub TxtPrazo_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtSit_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtVenda_KeyPress(KeyAscii As Integer)
    KeyAscii = Num(KeyAscii)
End Sub

Private Sub Fornecedor()
    On Error Resume Next
    Comando = "Select Razao From Fornecedor Order By Razao"

    Set RsTemp = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

    CboFornecedor.Clear
    CboFornecedor.AddItem ""
    If RsTemp.RecordCount <> 0 Then
        Do While Not RsTemp.EOF
            CboFornecedor.AddItem RsTemp!Razao
            RsTemp.MoveNext
        Loop
    End If
    RsTemp.Close
    E
End Sub
Private Sub Centro()
    On Error Resume Next
    Comando = "Select * From Centro Order By Descricao"

    Set RsTemp = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

    CboCusto.Clear
    CboCusto.AddItem ""
    If RsTemp.RecordCount <> 0 Then
        Do While Not RsTemp.EOF
            CboCusto.AddItem RsTemp!Descricao
            RsTemp.MoveNext
        Loop
    End If
    RsTemp.Close
    E
End Sub
Private Sub LocalObra()
    On Error Resume Next
    Comando = "Select * From Obra Order By Razao"

    Set RsTemp = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

    CboLocal.Clear
    CboLocal.AddItem ""
    If RsTemp.RecordCount <> 0 Then
        Do While Not RsTemp.EOF
            CboLocal.AddItem RsTemp!Razao
            RsTemp.MoveNext
        Loop
    End If
    RsTemp.Close
    E
End Sub

Private Sub Grupo()
    On Error Resume Next

    Comando = "Select * From GrupoProd Order By Descricao"    ' Order By Razao"

    Set RsTemp = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

    CboGrupo.Clear
    CboGrupo.AddItem ""
    If RsTemp.RecordCount <> 0 Then
        Do While Not RsTemp.EOF
            CboGrupo.AddItem RsTemp!Descricao
            RsTemp.MoveNext
        Loop
    End If
    RsTemp.Close
    E
End Sub



Private Sub Abrir(Tipo As Boolean)
    On Error Resume Next
    Dim RsProduto As Recordset, Passa                       As Boolean

    If Tipo = True Then
        If Trim(txtCodigo.Text) = "" And Trim(TxtClas.Text) = "" Then
            MsgBox "O Codigo e Obrigatorio", vbCritical, App.Title
            txtCodigo.SetFocus
            Exit Sub
        End If
    End If
    Passa = False
    If Tipo = False Then
        Comando = "Select * From Produtos where Codigo = '" & txtCodigo.Text & "' Order By Codigo"
    Else
        If Trim(txtCodigo.Text) = "" Then
            Comando = "Select * From Produtos Where Class = '" & TxtClas.Text & "' Order By Codigo"
        Else
            Comando = "Select * From Produtos where Codigo = '" & txtCodigo.Text & "' Order By Codigo"
        End If
    End If

    Set RsProduto = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)

    If Tipo = True Then
        If RsProduto.RecordCount = 0 And Trim(txtCodigo.Text) = "" Then
            MsgBox "Registro inexistente", vbCritical, App.Title
            Novo False
            txtCodigo.SetFocus
            Exit Sub
        Else
            Comando = "Select * From Produtos Where Class = '" & TxtClas.Text & "' Order By Codigo"

            Set RsTemp = BancoDeDados.OpenRecordset(Comando, dbOpenSnapshot)
            If RsTemp.RecordCount <> 0 Then
                If txtCodigo <> RsTemp!Codigo Then
                    MsgBox "A Classificação já esta cadastro com o Codigo " & RsTemp!Codigo & " ", vbInformation, App.Title
                    TxtClas.Text = ""
                    Exit Sub
                End If
            End If
        End If
    End If
    If RsProduto.RecordCount = 0 Then
        Novo True
    Else
        With RsProduto
            TxtAtual.Text = !Atual
            TxtBruto.Text = !Bruto
            TxtClassFiscal.Text = !Fiscal
            TxtCompra.Text = !Compra
            TxtDescricao.Text = !Descricao
            txtEmbalagem.Text = !Embalagem
            TxtEspecie.Text = !Especie
            TxtFabricante.Text = !Fabricante
            TxtIpi.Text = !Ipi
            TxtLiq.Text = !Liq
            TxtMarca.Text = !Marca
            Txtmax.Text = !Max
            TxtMedida.Text = !Medida
            TxtMedio.Text = !Medio
            TxtModelo.Text = !Modelo
            TxtNumero.Text = !Numero
            TxtPrazo.Text = !Prazo
            TxtSerie.Text = !Serie
            TxtSit.Text = !Sit
            TxtVenda.Text = !Venda
            mskFabricacao.Text = !Fabricacao
            MskSaida.Text = !Saida
            MskEntrada.Text = !Entrada
            MskValidade.Text = !Validade
            txtCodigo.Text = !Codigo
            TxtClas.Text = !Class
            TxtICMS.Text = IIf(IsNull(!ICms), "", !ICms)
            TxtComissao.Text = IIf(IsNull(!Comissao), "", !Comissao)
            If Trim(!Fornecedor) = "" Or IsNull(!Fornecedor) Then
                CboFornecedor.ListIndex = -1
            Else
                CboFornecedor.Text = !Fornecedor
            End If
            If Trim(!Grupo) = "" Or IsNull(!Grupo) Then
                CboGrupo.ListIndex = -1
            Else
                CboGrupo.Text = !Grupo
            End If
            '        If Trim(!Local) = "" Or IsNull(!Local) Then
            '            CboLocal.ListIndex = -1
            '        Else
            '            CboLocal.Text = !Local
            '       End If
            If Trim(!Centro) = "" Or IsNull(!Centro) Then
                CboCusto.ListIndex = -1
            Else
                CboCusto.Text = !Centro
            End If
            NomeFig = DiretorioDados + "FigEst" + Format(txtCodigo.Text, "0000") + ".Fig"
            If Dir(NomeFig) = "" Then
                Prod.Picture = i.Picture
            Else
                Prod.Picture = LoadPicture(NomeFig)
            End If
        End With
    End If
    RsProduto.Close
    'E
End Sub


Private Sub Salvar()
    On Error Resume Next
    Dim RsProduto                                           As Recordset

    If Trim(txtCodigo.Text) = "" Then
        MsgBox "O Codigo e Obrigatorio", vbCritical, App.Title
        txtCodigo.SetFocus
        Exit Sub
    End If
    If Trim(TxtClas.Text) = "" Then
        MsgBox "A Classificação e Obrigatorio", vbCritical, App.Title
        TxtClas.SetFocus
        Exit Sub
    End If

    If Trim(TxtDescricao.Text) = "" Then
        MsgBox "A Descrição e Obrigatorio", vbInformation, App.Title
        TxtDescricao.SetFocus
        Exit Sub
    End If

    Comando = "Select * From Produtos where Codigo = '" & txtCodigo.Text & "' And Class = '" & TxtClas.Text & "' Order By Codigo"

    Set RsProduto = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

    If RsProduto.RecordCount = 0 Then
        RsProduto.AddNew
    Else
        RsProduto.Edit
    End If

    With RsProduto
        !Atual = TxtAtual.Text
        !Bruto = TxtBruto.Text
        !Fiscal = TxtClassFiscal.Text
        !Compra = TxtCompra.Text
        !Descricao = TxtDescricao.Text
        !Embalagem = txtEmbalagem.Text
        !Especie = TxtEspecie.Text
        !Fabricante = TxtFabricante.Text
        !Ipi = TxtIpi.Text
        !Liq = TxtLiq.Text
        !Marca = TxtMarca.Text
        !Max = Txtmax.Text
        !Medida = TxtMedida.Text
        !Medio = TxtMedio.Text
        !Modelo = TxtModelo.Text
        !Numero = TxtNumero.Text
        !Prazo = TxtPrazo.Text
        !Serie = TxtSerie.Text
        !Sit = TxtSit.Text
        !Venda = TxtVenda.Text
        !Fabricacao = mskFabricacao.Text
        !Saida = MskSaida.Text
        !Entrada = MskEntrada.Text
        !Validade = MskValidade.Text
        !Codigo = txtCodigo.Text
        !Class = TxtClas.Text
        !Fornecedor = CboFornecedor.Text
        !Grupo = CboGrupo.Text
        !Local = CboLocal.Text
        !Centro = CboCusto.Text
        !ICms = TxtICMS.Text
        !Comissao = TxtComissao.Text
        .Update
    End With
    NomeFig = DiretorioDados + "FigEst" + Format(txtCodigo.Text, "0000") + ".Fig"
    If Dir(NomeFig) <> "" Then Kill NomeFig
    SavePicture Prod, NomeFig
    Novo False
    RsProduto.Close
    E
End Sub


Private Sub Excluir()
    On Error Resume Next
    Dim RsProduto                                           As Recordset

    If Trim(txtCodigo.Text) = "" Then
        MsgBox "O Codigo e Obrigatorio", vbCritical, App.Title
        txtCodigo.SetFocus
        Exit Sub
    End If

    If Trim(TxtClas.Text) = "" Then
        MsgBox "A Classificação e Obrigatorio", vbCritical, App.Title
        TxtClas.SetFocus
        Exit Sub
    End If

    Comando = "Select * From Produtos where Codigo = '" & txtCodigo.Text & "' And Class = '" & TxtClas.Text & "' Order By Codigo"

    Set RsProduto = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

    If RsProduto.RecordCount = 0 Then
        MsgBox "Registro inexistente", vbInformation, App.Title
        Exit Sub
    Else
        If MsgBox("Confirma Exclusão ?", vbCritical + vbYesNo + vbDefaultButton2 + vbSystemModal, App.Title) = vbYes Then
            RsProduto.Delete
            NomeFig = DiretorioDados + "FigEst" + Format(txtCodigo.Text, "0000") + ".Fig"
            If Dir(NomeFig) <> "" Then
                Kill NomeFig
            End If
            Novo False
            MsgBox "Registro Excluido com Sucesso !", vbCritical, App.Title
            txtCodigo.SetFocus
        End If
    End If
    E

End Sub
