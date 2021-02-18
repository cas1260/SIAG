VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCadCLiente 
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "FrmCadCLiente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   6840
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   30
      TabIndex        =   42
      Top             =   570
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados Pessoais (F1)"
      TabPicture(0)   =   "FrmCadCLiente.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Endereços (F2)"
      TabPicture(1)   =   "FrmCadCLiente.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Dados Complementares (F3)"
      TabPicture(2)   =   "FrmCadCLiente.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Diversos (F7)"
      TabPicture(3)   =   "FrmCadCLiente.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1"
      Tab(3).Control(1)=   "Tipo"
      Tab(3).Control(2)=   "MskVenc"
      Tab(3).Control(3)=   "Label34"
      Tab(3).ControlCount=   4
      Begin VB.PictureBox Picture1 
         Height          =   1065
         Left            =   -74910
         ScaleHeight     =   1005
         ScaleWidth      =   6495
         TabIndex        =   88
         Top             =   2340
         Width           =   6555
      End
      Begin VB.Frame Tipo 
         Caption         =   "Tipo de Correcao"
         ForeColor       =   &H8000000D&
         Height          =   1485
         Left            =   -74910
         TabIndex        =   85
         Top             =   810
         Width           =   6555
         Begin VB.TextBox PorCe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3090
            TabIndex        =   41
            Top             =   1020
            Width           =   1305
         End
         Begin VB.OptionButton OpS4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Inf.Anual"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   5370
            TabIndex        =   40
            Top             =   690
            Width           =   1005
         End
         Begin VB.OptionButton OpS3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Val.Fixo"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3780
            TabIndex        =   39
            Top             =   690
            Width           =   1245
         End
         Begin VB.OptionButton OpS2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Inflacao Mensal"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1590
            TabIndex        =   38
            Top             =   690
            Width           =   1425
         End
         Begin VB.OptionButton OpS1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Sal.Min."
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   150
            TabIndex        =   37
            Top             =   630
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.TextBox TxtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3090
            TabIndex        =   36
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label Label36 
            Caption         =   "Porcentagem Sobre o Salario Minimo :"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   150
            TabIndex        =   87
            Top             =   1140
            Width           =   2715
         End
         Begin VB.Label Label35 
            Caption         =   "Valor dos Honorarios/Manutencao. . . :"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   210
            TabIndex        =   86
            Top             =   300
            Width           =   2775
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Referencias"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   -74940
         TabIndex        =   78
         Top             =   1950
         Width           =   6615
         Begin VB.TextBox TxtNomeRef2 
            Height          =   285
            Left            =   660
            TabIndex        =   33
            Top             =   600
            Width           =   3645
         End
         Begin VB.TextBox TxtNomeRef1 
            Height          =   285
            Left            =   660
            TabIndex        =   31
            Top             =   240
            Width           =   3645
         End
         Begin MSMask.MaskEdBox MskTefRef1 
            Height          =   285
            Left            =   5100
            TabIndex        =   32
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTefRef2 
            Height          =   285
            Left            =   5100
            TabIndex        =   34
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label32 
            Caption         =   "Telefone"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4380
            TabIndex        =   82
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label31 
            Caption         =   "Nome :"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   60
            TabIndex        =   81
            Top             =   660
            Width           =   525
         End
         Begin VB.Label Label30 
            Caption         =   "Telefone"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4380
            TabIndex        =   80
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label29 
            Caption         =   "Nome :"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   60
            TabIndex        =   79
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados Complementares"
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   -74940
         TabIndex        =   75
         Top             =   450
         Width           =   6615
         Begin VB.TextBox TxtFil2 
            Height          =   285
            Left            =   1620
            TabIndex        =   30
            Top             =   900
            Width           =   4905
         End
         Begin VB.TextBox TxtFil1 
            Height          =   285
            Left            =   1620
            TabIndex        =   29
            Top             =   570
            Width           =   4905
         End
         Begin MSMask.MaskEdBox MskDataNasc 
            Height          =   285
            Left            =   1620
            TabIndex        =   28
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label28 
            Caption         =   "Filiação"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   960
            TabIndex        =   77
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "Data de nascimento:"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   76
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Endereco de Cobrança"
         ForeColor       =   &H8000000D&
         Height          =   1425
         Left            =   -74910
         TabIndex        =   66
         Top             =   1890
         Width           =   6585
         Begin VB.TextBox txtEndCob 
            Height          =   285
            Left            =   870
            TabIndex        =   20
            Top             =   240
            Width           =   2385
         End
         Begin VB.TextBox TxtBairroCob 
            Height          =   285
            Left            =   3960
            TabIndex        =   21
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox TxtCidadeCob 
            Height          =   285
            Left            =   870
            TabIndex        =   22
            Top             =   630
            Width           =   1155
         End
         Begin VB.TextBox TxtEstadoCob 
            Height          =   285
            Left            =   2820
            MaxLength       =   2
            TabIndex        =   23
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxTCepCob 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   24
            Top             =   630
            Width           =   2505
         End
         Begin VB.TextBox TxtCxpCob 
            Height          =   285
            Left            =   870
            TabIndex        =   25
            Top             =   1020
            Width           =   1155
         End
         Begin VB.TextBox TxTContCob 
            Height          =   285
            Left            =   4860
            TabIndex        =   27
            Top             =   1020
            Width           =   1605
         End
         Begin MSMask.MaskEdBox MskTelefoneCob 
            Height          =   285
            Left            =   2820
            TabIndex        =   26
            Top             =   1020
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            Caption         =   "Endereço"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   74
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "Bairro"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3360
            TabIndex        =   73
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Cidade"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Estado"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   2160
            TabIndex        =   71
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label22 
            Caption         =   "Cep"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3390
            TabIndex        =   70
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Cx. P"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   69
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Telefone"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2100
            TabIndex        =   68
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label Label19 
            Caption         =   "Cont"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4320
            TabIndex        =   67
            Top             =   1110
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Endereço de Entrega"
         ForeColor       =   &H8000000D&
         Height          =   1425
         Left            =   -74910
         TabIndex        =   57
         Top             =   390
         Width           =   6585
         Begin VB.TextBox TxTContEnt 
            Height          =   285
            Left            =   4860
            TabIndex        =   19
            Top             =   1020
            Width           =   1605
         End
         Begin MSMask.MaskEdBox MskTelefoneEnt 
            Height          =   285
            Left            =   2820
            TabIndex        =   18
            Top             =   1020
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "(##) ####-####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCxpEnt 
            Height          =   285
            Left            =   870
            TabIndex        =   17
            Top             =   1020
            Width           =   1155
         End
         Begin VB.TextBox TxTCepEnt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   16
            Top             =   630
            Width           =   2505
         End
         Begin VB.TextBox TxtEstadoEst 
            Height          =   285
            Left            =   2820
            MaxLength       =   2
            TabIndex        =   15
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtCidadeEst 
            Height          =   285
            Left            =   870
            TabIndex        =   14
            Top             =   630
            Width           =   1155
         End
         Begin VB.TextBox TxtBairroEnt 
            Height          =   285
            Left            =   3960
            TabIndex        =   13
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txtEndEnt 
            Height          =   285
            Left            =   870
            TabIndex        =   12
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label18 
            Caption         =   "Cont"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4320
            TabIndex        =   65
            Top             =   1110
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2100
            TabIndex        =   64
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label Label16 
            Caption         =   "Cx. P"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   63
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Cep"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3390
            TabIndex        =   62
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Estado"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   2160
            TabIndex        =   61
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Bairro"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   3360
            TabIndex        =   59
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Endereço"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   58
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados Financeiros"
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   90
         TabIndex        =   52
         Top             =   2400
         Width           =   6615
         Begin VB.TextBox txtLim 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3630
            TabIndex        =   11
            Top             =   600
            Width           =   2865
         End
         Begin MSMask.MaskEdBox MskDataCompra 
            Height          =   285
            Left            =   1410
            TabIndex        =   10
            Top             =   600
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboVendedor 
            Height          =   315
            Left            =   3630
            TabIndex        =   9
            Top             =   210
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.ComboBox CboTipoCli 
            Height          =   315
            Left            =   750
            TabIndex        =   8
            Top             =   210
            Width           =   1995
         End
         Begin VB.Label Label10 
            Caption         =   "Lim.Credito:"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2760
            TabIndex        =   56
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Label9 
            Caption         =   "1a.compra"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   90
            TabIndex        =   55
            Top             =   690
            Width           =   765
         End
         Begin VB.Label Label8 
            Caption         =   "Vendedor.:"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   2790
            TabIndex        =   54
            Top             =   300
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Cli.:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   90
            TabIndex        =   53
            Top             =   300
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Cadastrais"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   90
         TabIndex        =   48
         Top             =   1350
         Width           =   6645
         Begin VB.TextBox TxtInsMun 
            Height          =   285
            Left            =   4170
            TabIndex        =   7
            Top             =   600
            Width           =   2385
         End
         Begin VB.TextBox TxtInscEst 
            Height          =   285
            Left            =   780
            TabIndex        =   6
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox TxtFant 
            Height          =   285
            Left            =   780
            TabIndex        =   5
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label Label6 
            Caption         =   "Insc.Mun.:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3330
            TabIndex        =   51
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Insc.Est.:"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   90
            TabIndex        =   50
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Fantasia :"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   60
            TabIndex        =   49
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cadastro"
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   90
         TabIndex        =   44
         Top             =   330
         Width           =   6645
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
         Begin VB.TextBox TxtNome 
            Height          =   285
            Left            =   3780
            TabIndex        =   4
            Top             =   600
            Width           =   2745
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
         Begin VB.TextBox TxtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   870
            TabIndex        =   0
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label33 
            Caption         =   "Tipo"
            ForeColor       =   &H8000000D&
            Height          =   165
            Left            =   2310
            TabIndex        =   83
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Razao / Nome"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2640
            TabIndex        =   47
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Cpf/Cnpj"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   90
            TabIndex        =   46
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   90
            TabIndex        =   45
            Top             =   300
            Width           =   555
         End
      End
      Begin MSMask.MaskEdBox MskVenc 
         Height          =   285
         Left            =   -71820
         TabIndex        =   35
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label34 
         Caption         =   "Vencimento dos Honorarios/Manutencao :"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74910
         TabIndex        =   84
         Top             =   510
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   360
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
            Picture         =   "FrmCadCLiente.frx":04B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":05C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":06DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":07EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":0902
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":0A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCLiente.frx":0B2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
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
Attribute VB_Name = "FrmCadCLiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsCliente                                           As Recordset

Private Sub CboTipoCli_GotFocus()
1   SendKeys "{F4}"
End Sub
Private Sub CboVendedor_GotFocus()
1   SendKeys "{TAB}"
End Sub

Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   Select Case UCase(Button.Key)
    Case "A"
2       Novo
3       txtCodigo.SetFocus
4   Case "B"
5       Abrir
6   Case "C"
7       Salvar
8   Case "D"
9       Excluir
10  Case "E"
11      Localizar
12  Case "F"
13      Relatorio
14  Case "G"
15      Unload Me
16  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       SendKeys "{TAB}"
3   ElseIf KeyCode = 27 Then
4       Unload Me
5   ElseIf KeyCode = vbKeyF1 Then
6       SSTab1.Tab = 0
7       txtCodigo.SetFocus
8   ElseIf KeyCode = vbKeyF2 Then
9       SSTab1.Tab = 1
10      txtEndEnt.SetFocus
11  ElseIf KeyCode = vbKeyF3 Then
12      SSTab1.Tab = 2
13      MskDataNasc.SetFocus
14  ElseIf KeyCode = vbKeyF7 Then
15      SSTab1.Tab = 3
16      MskVenc.SetFocus
17  End If
End Sub

Private Sub Form_Load()
1   AbrirTipo
2   Me.Width = 6960
3   Me.Height = 4470
4   Centra Me
5   SSTab1.TabVisible(3) = False
End Sub

Private Sub Image1_Click()

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

Private Sub MskDataNasc_GotFocus()
1   SSTab1.Tab = 2
End Sub

Private Sub MskDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskDataNasc.Text = Valida(MskDataNasc)
3   End If
End Sub

Private Sub MskVenc_GotFocus()
1   SSTab1.Tab = 3
End Sub

Private Sub MskVenc_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskVenc.Text = Valida(MskVenc)
3   End If
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

Private Sub PorCe_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       Salvar
        'SendKeys "+{TAB}"
3       SSTab1.SetFocus
4   End If
End Sub

Private Sub PorCe_KeyPress(KeyAscii As Integer)
1   If KeyAscii = 9 Then
2       Salvar
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

Private Sub txtEndEnt_GotFocus()
1   SSTab1.Tab = 1
End Sub

Private Sub txtLim_GotFocus()
1   SSTab1.Tab = 0
End Sub

Private Sub txtLim_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SSTab1.Tab = 1
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
2   txtCodigo.Text = ""
3   OPF.Value = False
4   MskCpfCnpj.Mask = ""
5   MskCpfCnpj.Text = ""
6   MskCpfCnpj.Mask = "###.###.###-##"
7   TxtNome.Text = ""
8   TxtFant.Text = ""
9   TxtInscEst.Text = ""
10  TxtInsMun.Text = ""
11  MskDataCompra.Mask = ""
12  MskDataCompra.Text = ""
13  MskDataCompra.Mask = "##/##/####"
14  txtLim.Text = ""
15  txtEndCob.Text = ""
16  txtEndCob.Text = ""
17  TxtEstadoCob.Text = ""
18  TxtEstadoEst.Text = ""
19  TxtCidadeCob.Text = ""
20  TxtCidadeEst.Text = ""
21  TxtBairroCob.Text = ""
22  TxtBairroEnt.Text = ""
23  TxTCepCob.Text = ""
24  TxTCepEnt.Text = ""
25  TxtCxpCob.Text = ""
26  TxtCxpEnt.Text = ""
27  MskTefRef1.Mask = ""
28  MskTefRef1.Text = ""
29  MskTefRef1.Mask = "(##) ####-####"
30  MskTefRef2.Mask = ""
31  MskTefRef2.Text = ""
32  MskTefRef2.Mask = "(##) ####-####"
33  TxTContCob.Text = ""
34  TxTContEnt.Text = ""
35  MskDataNasc.Mask = ""
36  MskDataNasc.Text = ""
37  MskDataNasc.Mask = "##/##/####"
38  TxtFil1.Text = ""
39  TxtFil2.Text = ""
40  MskTelefoneCob.Mask = ""
41  MskTelefoneCob.Text = ""
42  MskTelefoneCob.Mask = "(##) ####-####"
43  MskTelefoneEnt.Mask = ""
44  MskTelefoneEnt.Text = ""
45  MskTelefoneEnt.Mask = "(##) ####-####"
46  TxtNomeRef1.Text = ""
47  TxtNomeRef2.Text = ""
48  MskVenc.Mask = ""
49  MskVenc.Text = ""
50  MskVenc.Mask = "##/##/####"
51  TxtValor.Text = ""
52  OpS1.Value = True
53  PorCe.Text = ""
54  txtEndEnt.Text = ""
55  OPF.Value = True
56  CboTipoCli.Text = ""
57  CboVendedor.Text = ""
58  AbrirTipo
Trata_Erro:
59  E
End Sub

Private Sub Abrir()
1   On Error GoTo Trata_Erro
    Dim VCodigo                                             As String
    Dim Tipo                                                As Long
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Cliente", vbCritical, App.Title
4       Exit Sub
5   End If

6   Comando = "Select * from Cliente Where Codigo =" & txtCodigo.Text & ""
7   Set RsCliente = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

8   If RsCliente.RecordCount = 0 Then
9       VCodigo = txtCodigo.Text
10      Novo
11      txtCodigo.Text = VCodigo
12  Else
13      txtCodigo.Text = RsCliente!Codigo
14      If UCase(RsCliente!pessoa) = "F" Then
15          OPF.Value = True
16          OPJ.Value = False
17      Else
18          OPF.Value = False
19          OPJ.Value = True
20      End If
21      MskCpfCnpj.Text = RsCliente!CNPJ
22      TxtNome.Text = RsCliente!Razao
23      TxtFant.Text = RsCliente!Fantasia
24      TxtInscEst.Text = RsCliente!Estadual
25      TxtInsMun.Text = RsCliente!Municipal
26      CboTipoCli.Text = RsCliente!TipoCliente
27      CboVendedor.Text = RsCliente!Vendedor
28      MskDataCompra.Text = RsCliente!PCompra
29      txtLim.Text = RsCliente!Limite
30      txtEndEnt.Text = RsCliente!EnderecoE
31      TxtBairroEnt.Text = RsCliente!BairroE
32      TxtCidadeEst.Text = RsCliente!CidadeE
33      TxTCepEnt.Text = RsCliente!CepE
34      TxtCxpEnt.Text = RsCliente!CaixaE
35      MskTelefoneEnt.Text = RsCliente!TelefoneE
36      TxTContEnt.Text = RsCliente!ContE
37      txtEndCob.Text = RsCliente!EnderecoC
38      TxtBairroCob.Text = RsCliente!BairroC
39      TxtCidadeCob.Text = RsCliente!CidadeC
40      TxTCepCob.Text = RsCliente!CepC
41      TxtCxpCob.Text = RsCliente!CaixaC
42      MskTelefoneCob.Text = RsCliente!TelefoneC
43      TxTContCob.Text = RsCliente!ContC
44      MskDataNasc.Text = RsCliente!DataN
45      TxtFil1.Text = RsCliente!Filiacao1
46      TxtFil2.Text = RsCliente!Filiacao2
47      TxtNomeRef1.Text = RsCliente!Ref1
48      TxtNomeRef2.Text = RsCliente!Ref2
49      MskTefRef1.Text = RsCliente!Tel1
50      MskTefRef2.Text = RsCliente!Tel2
51      MskVenc.Text = RsCliente!Vencimento
52      TxtValor.Text = RsCliente!Valor
53      PorCe.Text = RsCliente!Porc
54      Tipo = RsCliente!Op
55      TxtEstadoCob.Text = RsCliente!UFC
56      TxtEstadoEst.Text = RsCliente!UFE
57      If Tipo = 1 Then
58          OpS1.Value = True
59      ElseIf Tipo = 2 Then
60          OpS2.Value = True
61      ElseIf Tipo = 3 Then
62          OpS3.Value = True
63      Else
64          OpS4.Value = True
65      End If
66  End If
67  RsCliente.Close
Trata_Erro:
68  E
End Sub
Private Sub Salvar()
1   On Error GoTo Trata_Erro
    Dim VCodigo                                             As String
    Dim Tipo                                                As Long
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Cliente"
4       Exit Sub
5   End If
6   MskDataCompra.Text = Valida(MskDataCompra)
7   MskDataNasc.Text = Valida(MskDataNasc)
8   Comando = "Select * from Cliente Where Codigo =" & txtCodigo.Text & ""
9   Set RsCliente = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

10  If RsCliente.RecordCount = 0 Then
11      RsCliente.AddNew
12  Else
13      RsCliente.Edit
14  End If

15  RsCliente!Codigo = txtCodigo.Text
16  If OPF.Value = True Then
17      RsCliente!pessoa = "F"
18  Else
19      RsCliente!pessoa = "J"
20  End If
21  RsCliente!CNPJ = MskCpfCnpj.Text
22  RsCliente!Razao = TxtNome.Text
23  RsCliente!Fantasia = TxtFant.Text
24  RsCliente!Estadual = TxtInscEst.Text
25  RsCliente!Municipal = TxtInsMun.Text
26  RsCliente!TipoCliente = CboTipoCli.Text
27  RsCliente!Vendedor = CboVendedor.Text
28  RsCliente!PCompra = MskDataCompra.Text
29  RsCliente!Limite = txtLim.Text
30  RsCliente!EnderecoE = txtEndEnt.Text
31  RsCliente!BairroE = TxtBairroEnt.Text
32  RsCliente!CidadeE = TxtCidadeEst.Text
33  RsCliente!CepE = TxTCepEnt.Text
34  RsCliente!CaixaE = TxtCxpEnt.Text
35  RsCliente!TelefoneE = MskTelefoneEnt.Text
36  RsCliente!ContE = TxTContEnt.Text
37  RsCliente!EnderecoC = txtEndCob.Text
38  RsCliente!BairroC = TxtBairroCob.Text
39  RsCliente!CidadeC = TxtCidadeCob.Text
40  RsCliente!CepC = TxTCepCob.Text
41  RsCliente!CaixaC = TxtCxpCob.Text
42  RsCliente!TelefoneC = MskTelefoneCob.Text
43  RsCliente!ContC = TxTContCob.Text
44  RsCliente!DataN = MskDataNasc.Text
45  RsCliente!Filiacao1 = TxtFil1.Text
46  RsCliente!Filiacao2 = TxtFil2.Text
47  RsCliente!Ref1 = TxtNomeRef1.Text
48  RsCliente!Ref2 = TxtNomeRef2.Text
49  RsCliente!Tel1 = MskTefRef1.Text
50  RsCliente!Tel2 = MskTefRef2.Text
51  RsCliente!Vencimento = IIf(MskVenc.Text = "__/__/____", Date, MskVenc.Text)
52  If Trim(TxtValor.Text) = "" Then
53      TxtValor.Text = 0
54  End If
55  RsCliente!Valor = TxtValor.Text
56  RsCliente!Porc = PorCe.Text
57  RsCliente!UFC = TxtEstadoCob.Text
58  RsCliente!UFE = TxtEstadoEst.Text
59  If OpS1.Value = True Then
60      Tipo = 1
61  ElseIf OpS2.Value = True Then
62      Tipo = 2
63  ElseIf OpS3.Value = True Then
64      Tipo = 3
65  Else
66      Tipo = 4
67  End If
68  RsCliente!Op = Tipo
69  RsCliente.Update
70  RsCliente.Close
71  Novo
72  SSTab1.Tab = 0
73  txtCodigo.SetFocus
Trata_Erro:
74  E
End Sub

Public Sub Excluir()
1   On Error GoTo Trata_Erro
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "E Preciso digitar o Codigo do Cliente", vbCritical, App.Title
4       Exit Sub
5   End If

6   Comando = "Select * from Cliente Where Codigo =" & txtCodigo.Text & ""
7   Set RsCliente = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

8   If RsCliente.RecordCount <> 0 Then
9       If MsgBox("Confirma Exclusão ?", vbCritical + vbYesNo + vbDefaultButton2 + vbSystemModal, App.Title) = vbYes Then
10          RsCliente.Delete
11          RsCliente.Close
12          MsgBox "Cliente Excluido com Sucesso!", vbInformation, App.Title
13          Novo
14      End If
15  Else
16      MsgBox "Impossivel Excluir"
17  End If
Trata_Erro:
18  E
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
1   kleyascii = Num(KeyAscii)
End Sub

Private Sub Localizar()
1   On Error GoTo Trata_Erro
2   ShowCliente = 0
3   FrmPesqCliente.Show 1
4   If ShowCliente <> 0 Then
5       txtCodigo.Text = ShowCliente
6       Abrir
7   End If
Trata_Erro:
8   E
End Sub

Private Sub AbrirTipo()
1   On Error GoTo Trata_Erro
    Dim RsTipo                                              As Recordset
2   Comando = "Select * from TipoCli Order By Descricao"
3   Set RsTipo = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
4   CboTipoCli.Clear
5   If RsTipo.RecordCount <> 0 Then
6       Do While Not RsTipo.EOF
7           CboTipoCli.AddItem RsTipo!Descricao
8           RsTipo.MoveNext
9       Loop
10  End If
11  RsTipo.Close
Trata_Erro:
12  E
End Sub

Private Sub Relatorio()
1   On Error Resume Next
    Dim RsCli                                               As Recordset
    Dim X                                                   As Long
2   Comando = "Select * from Cliente Order By Codigo"
3   Set RsCli = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)

4   Set RsCliente = BancoRel.OpenRecordset("RCliente", dbOpenDynaset)
5   Do While Not RsCliente.EOF
6       RsCliente.Delete
7       RsCliente.MoveNext
8   Loop
9   Me.MousePointer = 11
10  Do While Not RsCli.EOF
11      RsCliente.AddNew
12      RsCliente!Codigo = RsCli!Codigo
13      RsCliente!pessoa = IIf(UCase(RsCli!pessoa) = "J", "Juritica", "Fisica")
14      RsCliente!CNPJ = RsCli!CNPJ
15      RsCliente!Razao = RsCli!Razao
16      RsCliente!Fantasia = RsCli!Fantasia
17      RsCliente!Estadual = RsCli!Estadual
18      RsCliente!Municipal = RsCli!Municipal
19      RsCliente!TipoCliente = RsCli!TipoCliente
20      RsCliente!Vendedor = RsCli!Vendedor
21      RsCliente!PCompra = RsCli!PCompra
22      RsCliente!Limite = Format(RsCli!Limite, "##,##0.00")
23      RsCliente!EnderecoE = RsCli!EnderecoE
24      RsCliente!BairroE = RsCli!EnderecoE
25      RsCliente!CidadeE = RsCli!CidadeE
26      RsCliente!CepE = RsCli!CepE
27      RsCliente!CaixaE = RsCli!CaixaE
28      RsCliente!TelefoneE = RsCli!TelefoneE
29      RsCliente!ContE = RsCli!ContE
30      RsCliente!EnderecoC = RsCli!EnderecoC
31      RsCliente!BairroC = RsCli!BairroC
32      RsCliente!CidadeC = RsCli!CidadeC
33      RsCliente!CepC = RsCli!CepC
34      RsCliente!CaixaC = RsCli!CaixaC
35      RsCliente!TelefoneC = RsCli!TelefoneC
36      RsCliente!ContC = RsCli!ContC
37      RsCliente!DataN = RsCli!DataN
38      RsCliente!Filiacao1 = RsCli!Filiacao1
39      RsCliente!Filiacao2 = RsCli!Filiacao2
40      RsCliente!Ref1 = RsCli!Ref1
41      RsCliente!Ref2 = RsCli!Ref2
42      RsCliente!Tel1 = RsCli!Tel1
43      RsCliente!Tel2 = RsCli!Tel2
44      RsCliente!Vencimento = RsCli!Vencimento
45      RsCliente!Valor = Format(RsCli!Valor, "##,##0.00")
46      RsCliente!Porc = RsCli!Porc
47      RsCliente!UFC = RsCli!UFC
48      RsCliente!UFE = RsCli!UFE
49      RsCliente!Op = RsCli!Op
50      RsCliente!Emp = NomeEmpresa
51      RsCliente!Hora = Str(Time)
52      RsCliente!Data = Str(Date)
53      RsCliente.Update
54      RsCli.MoveNext
55  Loop
56  Rel.ReportFileName = DiretorioRel + "LCliente.Rpt"
57  Rel.DataFiles(0) = LocalRel
58  For X = 0 To 1000: Next X

59  Rel.Action = 1
60  Me.MousePointer = 0
61  E
End Sub
