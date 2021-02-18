VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   2820
         Top             =   2970
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright : Cleber de Almeida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa  : Neo SoftWare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção: Neo SoftWare Tem todos os direitos autorais sobre este SoftWare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6045
         TabIndex        =   5
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 9x, Nt, 2000, XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         TabIndex        =   6
         Top             =   2340
         Width           =   3825
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIAG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizando Para:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neo SoftWare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   7
         Top             =   705
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Passo                                                   As Boolean
Option Explicit

Private Sub Form_Load()
1   On Error Resume Next
2   If Dir("C:\Windows\", vbDirectory) = "" Then
3       MkDir "C:\Windows"
4   End If
5   'If Date >= CDate("01/07/2004") Then
6   '     Open "C:\Windows\logMetalNeo.ini" For Output As #1
7   '         Print #1, Date
8   '     Close #1
9   '     Timer.Enabled = False
10  '     MsgBox "Run-Time : is not create form in dll line 685 form frmprincipal", vbCritical, "Count find dll"
11  '     End
12  ' End If

13  'If Dir("C:\Windows\logMetalNeo.ini") <> "" Then
14  '    MsgBox "Run-Time : is not create form in dll line 685 form frmprincipal", vbCritical, "Count find dll"
15  '    End
16  'End If

17  lblVersion.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
18  lblProductName.Caption = App.Title
19  Passo = True
20  FrmTelaLogin = False
End Sub

Private Sub Timer_Timer()
    If Passo = False Then
        Unload Me
        FrmPrincipal.Show
    Else
        If Right(App.Path, 1) = "\" Then
            CaminhoBanco = Ler("Banco", "Arquivo", "", App.Path & "Siag.ini")
            Aminacao = IIf(Ler("Tela", "Aminação", "1", App.Path & "Siag.ini") = "1", True, False)
            LocalImagem = Ler("Tela", "Log", "", App.Path & "Siag.ini")
        Else
            CaminhoBanco = Ler("Banco", "Arquivo", "", App.Path & "\Siag.ini")
            Aminacao = IIf(Ler("Tela", "Aminação", "1", App.Path & "\Siag.ini") = "1", True, False)
            LocalImagem = Ler("Tela", "Log", "", App.Path & "\Siag.ini")
        End If

        If CaminhoBanco = "" Or Dir(CaminhoBanco) = "" Then
            If InStr(1, CaminhoBanco, ":") = 0 Then
                CaminhoBanco = Left(App.Path, 2) & CaminhoBanco
                If CaminhoBanco = "" Or Dir(CaminhoBanco) = "" Then
                    MsgBox "Caro Usuario, Impossivel localizar o banco de dados", vbCritical, App.Title
                    End
                End If
            Else
                MsgBox "Caro Usuario, Impossivel localizar o banco de dados", vbCritical, App.Title
                End

            End If
        End If
        SenhaSistema = "1906bili"
        Set BancoDeDados = DAO.OpenDatabase(CaminhoBanco, dbDriverComplete, False, ";PWD=" & SenhaSistema)
        ''Set BancoDeDados = dao.OpenDatabase(CaminhoBanco)
        Passo = False
        LocalBanco = CaminhoBanco
    End If
End Sub
