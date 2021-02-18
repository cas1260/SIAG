VERSION 5.00
Begin VB.Form FrmTela 
   BorderStyle     =   0  'None
   ClientHeight    =   945
   ClientLeft      =   1650
   ClientTop       =   1590
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmTela.frx":0000
   ScaleHeight     =   945
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.Image II 
      Height          =   435
      Left            =   120
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "FrmTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
1   SendKeys "^{F6}"
End Sub

Private Sub Form_Load()
1   Me.Height = 1350
2   Me.Width = 4110
    'Centra Me
3   If LocalImagem <> "" Then
4       If Dir(LocalImagem) <> "" Then
5           II.Picture = LoadPicture(LocalImagem)
6           Me.Height = II.Height
7           Me.Width = II.Width
8           II.Visible = False
9           Me.Picture = II.Picture
10      End If
11  End If
End Sub
