Attribute VB_Name = "Ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
                                                 Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                   ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
                                                                                   ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
                                                   Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                       ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function Escreva(sSection As String, _
                        sKey As String, ByVal sValue As String, Arquivo As String) As Boolean

    Dim lR                                                  As Long
    Dim sItemValue                                          As String

1   sValue = Trim$(sValue)
2   sItemValue = Trim$(sItemValue) & vbNullChar
3   lR = WritePrivateProfileString(sSection, sKey, _
                                   sValue, Arquivo)

4   If lR = 0 Then
5       Escreva = False
6   Else
7       Escreva = True
8   End If
End Function


Public Function Ler(sSection As String, _
                    sKey As String, sDefault As String, Arquivo As String) As String

    Dim lR                                                  As Long
    Dim sReturnedValue                                      As String

1   sReturnedValue = Space$(512)
2   lR = GetPrivateProfileString(sSection, sKey, sDefault, _
                                 sReturnedValue, 512, Arquivo)
3   If lR = 0 Then
4       Ler = vbNullString
5   Else
6       Ler = Left$(sReturnedValue, lR)
7   End If
End Function


Sub Main()
    Dim sR                                                  As String
    Dim lR                                                  As String
1   lR = Escreva("Windows", "Sistema", "Windows 98", "D:\Windows\Win.ini")
2   sR = Ler("Windows", "Sistema", "", "D:\Windows\Win.Ini")
3   MsgBox "INI Value: " & sR
End Sub



Public Function EscrevaVar(sSection As String, _
                           sKey As String, ByVal sValue As String) As Boolean

    Dim lR                                                  As Long
    Dim sItemValue As String, Arquivo                       As String

1   Arquivo = "C:\TmpVar.win"
2   sValue = sValue
3   sItemValue = Trim$(sItemValue) & vbNullChar
4   lR = WritePrivateProfileString(sSection, sKey, _
                                   sValue + ".", Arquivo)

5   If lR = 0 Then
6       EscrevaVar = False
7   Else
8       EscrevaVar = True
9   End If
End Function


Public Function LerVar(sSection As String, _
                       sKey As String, sDefault As String) As String

    Dim lR                                                  As Long
    Dim sReturnedValue As String, Arquivo                   As String

1   Arquivo = "C:\TmpVar.win"
2   sReturnedValue = Space$(512)
3   lR = GetPrivateProfileString(sSection, sKey, sDefault, _
                                 sReturnedValue, 512, Arquivo)
4   If lR = 0 Then
5       LerVar = vbNullString
6   Else
7       LerVar = Left$(sReturnedValue, lR)
8       If LerVar <> sDefault Then
9           LerVar = Left(LerVar, Len(LerVar) - 1)
10      End If
11  End If
End Function


