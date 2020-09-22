VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDS PowerCrypt"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4575
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "/"
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Pass:"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "User:"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Please login with the admin username and password access and use CDS PowerCrypt."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Perfect Dark (BRK)"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Set our variables
Dim strUser, strPass As String

'Get the username and pass in the security file
DecryptSecure
strUser = ReadINI("Security", "User", App.Path & "\Security\CDS.secure")
strPass = ReadINI("Security", "Pass", App.Path & "\Security\CDS.secure")
EncryptSecure

'Compare username and password with one another
If txtUser.Text = strUser Then
    If txtPass.Text = strPass Then
        Unload Me
        frmMain.Show
    Else
        MsgBox "Invalid username or password!", vbCritical, "CDS"
        Exit Sub
    End If
Else
    MsgBox "Invalid username or password!", vbCritical, "CDS"
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
txtUser.Text = ""
txtPass.Text = ""
Unload Me
End
End Sub

Private Sub Form_Load()
'Set our variables
Dim strTemp As String
Dim strUser As String
Dim strPass As String
Dim strTimes As String

'Get the amount of times the security file has been created
strTimes = GetSetting("CDS PowerCrypt", "Security", "SFC", "0")

'Check if security file exists, if it doesnt then make it ONE time..
'If the security file disapears after its created, then we assume the program has been tampered with..
If FileExists(App.Path & "\Security\CDS.secure") = False Then
    If strTimes = "1" Then
        MsgBox "The security file has been deleted after it was created. For security reasons CDS PowerCrypt can not allow you to access the program. Please contact CDS tech support immediately if you require something to be decrypted.", vbCritical, "CDS"
        Unload Me
        End
    Else
        MkDir App.Path & "\Security"
        WriteINI "Security", "User", "", App.Path & "\Security\CDS.secure"
        WriteINI "Security", "Pass", "", App.Path & "\Security\CDS.secure"
        EncryptSecure
        strTimes = strTimes + 1
        SaveSetting "CDS PowerCrypt", "Security", "SFC", strTimes
    End If
End If

'Check if this is the first time the user has run PowerCrypt, if so then give temp access
DecryptSecure
strTemp = ReadINI("Security", "Pass", App.Path & "\Security\CDS.secure")
If strTemp = "" Then
    MsgBox "This is your first time running PowerCrypt. Please be enter a username and password at the following prompts.", vbInformation, "CDS"
    strUser = InputBox("Please choose a username:", "CDS")
    strPass = InputBox("Please choose a password:", "CDS")
    WriteINI "Security", "User", strUser, App.Path & "\Security\CDS.secure"
    WriteINI "Security", "Pass", strPass, App.Path & "\Security\CDS.secure"
    MsgBox "Your username and password have been saved. You may now login to PowerCrypt.", vbInformation, "CDS"
End If
EncryptSecure
End Sub
