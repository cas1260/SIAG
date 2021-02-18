VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTipoClie 
   Caption         =   "Cadastro de Tipo de Cliente"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   Icon            =   "FrmTipoClie.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   5010
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1050
      Width           =   3795
   End
   Begin VB.TextBox TxtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   690
      Width           =   1035
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
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
            Picture         =   "FrmTipoClie.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoClie.frx":0ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   953
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Key             =   "aaaa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "E"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Descrição"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   780
      Width           =   555
   End
End
Attribute VB_Name = "FrmTipoClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsTipo                                              As Recordset

Private Sub Fer_ButtonClick(ByVal Button As MSComctlLib.Button)
1   On Error GoTo Trata_Erro
2   Select Case Button.Key
    Case "A"
3       Novo
4   Case "B"
5       Abrir
6   Case "C"
7       Salvar
8   Case "D"
9       Excluir
10  Case "E"
11      Unload Me
12  Case "aaaa"
13      Localizar
14  End Select
Trata_Erro:
15  E
End Sub
Private Sub Novo()
1   On Error GoTo Trata_Erro
2   txtCodigo.Text = ""
3   TxtDescricao.Text = ""
Trata_Erro:
4   E
End Sub
Private Sub Abrir()
1   On Error GoTo Trata_Erro
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido !", vbExclamation, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If
7   Comando = "Select * from TipoCli Where Codigo = " & txtCodigo.Text & ""
8   Set RsTipo = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
9   If RsTipo.EOF = True Then
10      TxtDescricao.Text = ""
11  Else
12      TxtDescricao.Text = RsTipo!Descricao
13  End If
14  RsTipo.Close
Trata_Erro:
15  E
End Sub
Private Sub Salvar()
1   On Error GoTo Trata_Erro
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido !", vbExclamation, App.Title
4       txtCodigo.SetFocus
5       txtCodigo.SetFocus
6   End If
7   If Trim(TxtDescricao.Text) = "" Then
8       MsgBox "Descricao Invalida ! ! !", vbExclamation, App.Title
9       TxtDescricao.SetFocus
10      Exit Sub
11  End If

12  Comando = "Select * from TipoCli Where Codigo = " & txtCodigo.Text & ""
13  Set RsTipo = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
14  If RsTipo.EOF = True Then
15      RsTipo.AddNew
16  Else
17      RsTipo.Edit
18  End If
19  RsTipo!Codigo = txtCodigo.Text
20  RsTipo!Descricao = TxtDescricao.Text
21  RsTipo.Update
22  RsTipo.Close
23  Novo
    'TxtCodigo.SetFocus
Trata_Erro:
24  E
End Sub


Private Sub Excluir()
1   On Error GoTo Trata_Erro

2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido !", vbExclamation, App.Title
4       txtCodigo.SetFocus
5   End If
6   Comando = "Select * from TipoCli Where Codigo = " & txtCodigo.Text & ""
7   Set RsTipo = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
8   If RsTipo.EOF = False Then
9       If MsgBox("Deseja Excluir ?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
10          RsTipo.Delete
11          MsgBox "Tipo de cliente excluido com sucesso", vbInformation, App.Title
12          Novo
13          txtCodigo.SetFocus
14      End If
15  Else
16      MsgBox "Impossivel Excluir ! ! !", vbExclamation, App.Title
17  End If
Trata_Erro:
18  E
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then SendKeys "{TAB}"
2   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
1   Me.Height = 1875
2   Me.Width = 5130
3   Centra Me
End Sub
Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Abrir
End Sub
Private Sub TxtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Salvar
End Sub
Private Sub Localizar()
1   ShowTipoCliente = "0"
2   FrmLocalTipoCliente.Show 1
3   If ShowTipoCliente <> "0" Then
4       txtCodigo.Text = ShowTipoCliente
5       Abrir
6   End If
End Sub
