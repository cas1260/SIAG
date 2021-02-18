VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFechaCaixa 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6525
   ClientLeft      =   1575
   ClientTop       =   1875
   ClientWidth     =   9555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6510
      TabIndex        =   12
      Top             =   360
      Width           =   1545
   End
   Begin MSComctlLib.ProgressBar Barra 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5235
      Left            =   60
      TabIndex        =   6
      Top             =   810
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   $"FrmFechaCaixa.frx":0000
   End
   Begin VB.CommandButton CmdPesq 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   4980
      TabIndex        =   5
      Top             =   360
      Width           =   1545
   End
   Begin MSMask.MaskEdBox MskIni 
      Height          =   285
      Left            =   930
      TabIndex        =   2
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskFinal 
      Height          =   285
      Left            =   3390
      TabIndex        =   4
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   6120
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Total R$"
      Height          =   225
      Left            =   7320
      TabIndex        =   10
      Top             =   6150
      Width           =   1095
   End
   Begin VB.Label LblTotaliten 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6090
      TabIndex        =   9
      Top             =   6120
      Width           =   1125
   End
   Begin VB.Label lblT 
      Caption         =   "Total de Itens :"
      Height          =   225
      Left            =   5010
      TabIndex        =   7
      Top             =   6150
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Data Final"
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Labelrete 
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Relatorio de Caixa"
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
      TabIndex        =   0
      Top             =   0
      Width           =   9525
   End
End
Attribute VB_Name = "FrmFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPesq_Click()
1   On Error GoTo Trata_Erro



Trata_Erro:
2   E
End Sub

Private Sub MskFinal_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskFinal.Text = Valida(MskIni)
3   End If
End Sub
Private Sub MskIni_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       MskIni.Text = Valida(MskIni)
3   End If
End Sub
