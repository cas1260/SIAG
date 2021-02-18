VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCondPag 
   Caption         =   "Cadastro de Condições de Pagamento"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "FrmCondPag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   6735
   Begin VB.TextBox TxtJuro 
      Height          =   285
      Left            =   1290
      TabIndex        =   9
      Top             =   1470
      Width           =   1995
   End
   Begin VB.TextBox TxtDia 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   1140
      Width           =   2145
   End
   Begin VB.TextBox TxtDuPli 
      Height          =   285
      Left            =   1290
      TabIndex        =   6
      Top             =   1140
      Width           =   1995
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   3330
      TabIndex        =   4
      Top             =   750
      Width           =   3375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1290
      TabIndex        =   2
      Top             =   750
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   120
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
            Picture         =   "FrmCondPag.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCondPag.frx":0ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Fer 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   953
      ButtonWidth     =   1005
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "E"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Label L 
      Alignment       =   1  'Right Justify
      Caption         =   "Juros :    "
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Intervalo Dias"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3360
      TabIndex        =   7
      Top             =   1230
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "N. de Parcelas :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1230
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Descrição"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   690
      TabIndex        =   1
      Top             =   840
      Width           =   525
   End
End
Attribute VB_Name = "FrmCondPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsTipo As Recordset, RsPag                          As Recordset
Attribute RsPag.VB_VarUserMemId = 1073938432
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
11      Unload Me
12  End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1   On Error GoTo Trata_Erro
2   If KeyCode = 13 Then
3       SendKeys "{TAB}"
4   ElseIf KeyCode = 27 Then
5       Unload Me
6   End If
Trata_Erro:
7   E
End Sub

Private Sub Form_Load()
1   On Error GoTo Trata_Erro
2   Me.Height = 2280
3   Me.Width = 6855
4   Centra Me
Trata_Erro:
5   E
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then Abrir
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtDia_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub TxtDuPli_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub

Private Sub Novo()
1   On eroror GoTo Trata_Erro
2   txtCodigo.Text = ""
3   TxtDescricao.Text = ""
4   TxtDia.Text = ""
5   TxtDuPli.Text = ""
6   TxtJuro.Text = ""
Trata_Erro:
7   E
End Sub

Private Sub Abrir()
1   On Error Resume Next
    Dim CodAux                                              As String
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido ! ! !", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If
7   Comando = "Select * from CodPag Where Codigo = " & txtCodigo.Text & ""
8   Set RsPag = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
9   If RsPag.RecordCount = 0 Then
10      CodAux = txtCodigo.Text
11      Novo
12      txtCodigo.Text = CodAux
13  Else
14      TxtDescricao.Text = RsPag!Descricao
15      TxtDuPli.Text = RsPag!Dup
16      TxtDia.Text = RsPag!Dias
17      TxtJuro.Text = RsPag!Juros
18  End If
19  RsPag.Close
Trata_Erro:
20  E
End Sub

Private Sub Salvar()
1   On Error GoTo Trata_Erro
    Dim CodAux                                              As String
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido ! ! !", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If
7   Comando = "Select * from CodPag Where Codigo = " & txtCodigo.Text & ""
8   Set RsPag = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
9   If RsPag.RecordCount = 0 Then
10      RsPag.AddNew
11  Else
12      RsPag.Edit
13  End If
14  RsPag!Codigo = txtCodigo.Text
15  RsPag!Descricao = TxtDescricao.Text
16  RsPag!Dup = TxtDuPli.Text
17  RsPag!Dias = TxtDia.Text
18  RsPag!Juros = TxtJuro.Text
19  RsPag.Update
20  Novo

Trata_Erro:
21  E
End Sub
Private Sub Excluir()
1   On Error GoTo Trata_Erro
    Dim RsPag1                                              As Recordset
    Dim RsPagAux1                                           As Recordset
    Dim XX1                                                 As Boolean
    Dim XX2                                                 As Boolean
2   If Trim(txtCodigo.Text) = "" Then
3       MsgBox "Codigo Invalido ! ! ! ", vbCritical, App.Title
4       txtCodigo.SetFocus
5       Exit Sub
6   End If
7   If MsgBox("Confirma Excluir ? ", vbYesNo + vbCritical + vbDefaultButton2, App.Title) = vbYes Then
8       Comando = "Select * From CodPag Where Codigo =" & txtCodigo.Text & ""
9       Set RsPagAux1 = BancoDeDados.OpenRecordset(Comando, dbOpenDynaset)
10      If RsPagAux1.RecordCount <> 0 Then
11          RsPagAux1.Delete
12      Else
13          MsgBox "Impossivel Excluir ! ! !", vbCritical, App.Title
14      End If
15      RsPagAux1.Close
16      If XX1 = False And XX2 = False Then

17          Exit Sub
18      End If
19      MsgBox "Resgistro Excluir ! ! !", vbCritical, App.Title
20      Novo
21      txtCodigo.SetFocus
22  End If
Trata_Erro:
23  E
End Sub
Private Sub TxtJuro_KeyDown(KeyCode As Integer, Shift As Integer)
1   If KeyCode = 13 Then
2       Salvar
3   End If
End Sub
Private Sub TxtJuro_KeyPress(KeyAscii As Integer)
1   KeyAscii = Num(KeyAscii)
End Sub
