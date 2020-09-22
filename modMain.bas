Attribute VB_Name = "modMain"
'Declare any globals
Option Explicit

#If Win16 Then
    Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal Filename As String) As Integer
    Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal Filename As String) As Integer
#Else
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Private Crypt As clsCryptAPI
Dim Key As String

'-----------------------------------------------------------
'Our startup procedure
'-----------------------------------------------------------
Sub Main()
    Key = "cds rocks"
    frmLogin.Show
End Sub

'-----------------------------------------------------------
'All of our INI functions
'-----------------------------------------------------------
Function ReadINI(Section, KeyName, Filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), Filename))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    WritePrivateProfileString sSection, sKeyName, sNewString, sFileName
End Function

'---------------------------------------------------------
'All other functions
'---------------------------------------------------------
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        Open FullFileName For Input As #1
        Close #1
        FileExists = True
    Exit Function
MakeF:
        FileExists = False
    Exit Function
End Function

'-------------------------------------------------------
'Functions for encryption
'-------------------------------------------------------
Public Function EncryptSecure()
    Set Crypt = New clsCryptAPI
    Crypt.EncryptFile App.Path & "\Security\CDS.secure", App.Path & "\Security\CDS.secure", Key
End Function

Public Sub DecryptSecure()
    Set Crypt = New clsCryptAPI
    Crypt.DecryptFile App.Path & "\Security\CDS.secure", App.Path & "\Security\CDS.secure", Key
End Sub
