VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDS PowerCrypt"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "PowerCrypt"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "&Decrypt"
         Height          =   375
         Left            =   3360
         TabIndex        =   19
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Encrypt"
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Text            =   "PowerCrypt"
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Text            =   "C:\Text.dec"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Text            =   "C:\Text.txt"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   600
         Width           =   4335
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4440
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label2 
         Caption         =   "<N/A>"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "<N/A>"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "<N/A>"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   15
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Progess:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Time used:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Information:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4440
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label7 
         Caption         =   "Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Decrypt to File/Decrypted Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Output of encrypted File or Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "File or Text to encrypt:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Choose encryption method:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EncryptCryptAPI As clsCryptAPI
Private WithEvents EncryptTEA As clsTEA
Attribute EncryptTEA.VB_VarHelpID = -1
Private WithEvents EncryptGost As clsGost
Attribute EncryptGost.VB_VarHelpID = -1
Private WithEvents EncryptSkipJack As clsSkipjack
Attribute EncryptSkipJack.VB_VarHelpID = -1
Private WithEvents EncryptTwofish As clsTwofish
Attribute EncryptTwofish.VB_VarHelpID = -1
Private WithEvents EncryptBlowfish As clsBlowfish
Attribute EncryptBlowfish.VB_VarHelpID = -1
Private WithEvents EncryptXOR As clsSimpleXOR
Attribute EncryptXOR.VB_VarHelpID = -1
Private WithEvents EncryptRC4 As clsRC4
Attribute EncryptRC4.VB_VarHelpID = -1
Private WithEvents EncryptDES As clsDES
Attribute EncryptDES.VB_VarHelpID = -1

Private EncryptObject As Object

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub AddEncryption(Object As Object, Name As String, Optional Homepage As String)
  'Add encryption to internal array
  ReDim Preserve EncryptObjects(EncryptObjectsCount)
  With EncryptObjects(EncryptObjectsCount)
    Set .Object = Object
    .Name = Name
    .Homepage = Homepage
  End With
  EncryptObjectsCount = EncryptObjectsCount + 1
  
  'Add encryption to combobox
  Call Combo1.AddItem(Name)
  Combo1.ItemData(Combo1.NewIndex) = (EncryptObjectsCount - 1)
End Sub

Private Function CmpFile(File1 As String, File2 As String)
  Dim a As Long
  Dim S1 As String
  Dim S2 As String
  
  Open File1 For Binary As #1
  S1 = Space$(LOF(1))
  Get #1, , S1
  Close #1
  
  Open File2 For Binary As #2
  S2 = Space$(LOF(2))
  Get #2, , S2
  Close #2
  
  CmpFile = (S1 = S2)
End Function

Private Sub Combo1_Click()
  With EncryptObjects(Combo1.ItemData(Combo1.ListIndex))
    Set EncryptObject = .Object
  End With
End Sub

Private Sub Command1_Click()
  Dim OldTimer As Single
  
'  On Error GoTo ErrorHandler
  
  'Reset the labels
  Label2(0).Caption = "<unknown>"
  Label2(1).Caption = "<unknown>"
  Label2(2).Caption = "<unknown>"
  
  'If the text fields contain filenames we
  'want to encrypt the file given
  If (Mid$(Text1(0).Text, 2, 2) = ":\") Then
    If (Mid$(Text1(1).Text, 2, 2) = ":\") Then
      Label2(0).Caption = FileLen(Text1(0).Text) & " bytes"
      OldTimer = Timer
      Call EncryptObject.EncryptFile(Text1(0).Text, Text1(1).Text, Text1(3).Text)
      Label2(1).Caption = Timer - OldTimer
      Call MsgBox("File Encryption successful.")
      Exit Sub
    End If
  End If

  'Encrypt the content of the first textbox
  'and store the *hexadecimal* value in the
  'second textbox
  OldTimer = Timer
  Text1(1).Text = StrToHex(EncryptObject.EncryptString(Text1(0).Text, Text1(3).Text))
  Label2(1).Caption = Timer - OldTimer
  Exit Sub
  
Finished:
  Call MsgBox("Encryption/Decryption successful.", vbExclamation)
  Exit Sub
  
ErrorHandler:
  Call MsgBox("Warning: A major error occured:" & vbCrLf & vbCrLf & Err.Description, vbExclamation)
End Sub

Private Sub Command2_Click()
  Dim OldTimer As Single

  'On Error GoTo ErrorHandler
  
  'Reset the labels
  Label2(0).Caption = "<unknown>"
  Label2(1).Caption = "<unknown>"
  Label2(2).Caption = "<unknown>"
  
  'If the text fields contain filenames we
  'want to encrypt the file given
  If (Mid$(Text1(0).Text, 2, 2) = ":\") Then
    If (Mid$(Text1(1).Text, 2, 2) = ":\") Then
      Label2(0).Caption = FileLen(Text1(1).Text) & " bytes"
      OldTimer = Timer
      Call EncryptObject.DecryptFile(Text1(1).Text, Text1(2).Text, Text1(3).Text)
      Label2(1).Caption = Timer - OldTimer
      Call MsgBox("File Decryption successful.")
      Exit Sub
    End If
  End If

  'Decrypt the content of the second textbox
  'making sure to use the value from the Tag
  'property instead of the Text property
  Text1(2).Text = EncryptObject.DecryptString(HexToStr(Text1(1).Text), Text1(3).Text)
    
  Exit Sub
  
ErrorHandler:
  Call MsgBox("Warning: A major error occured:" & vbCrLf & vbCrLf & Err.Description, vbExclamation)
End Sub

Private Sub Command4_Click()
  On Error Resume Next
  
  Label2(0).Caption = BENCHMARKSIZE & " bytes"
  Label2(1).Caption = "<unknown>"
  Label2(2).Caption = "<unknown>"
  
  Call frmBenchmark.Show(vbModal, Me)
End Sub

Private Sub EncryptBlowfish_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptDES_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptGost_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptRC4_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptSkipJack_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptTEA_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptTwofish_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub EncryptXOR_Progress(Percent As Long)
  'Update the progress label
  Label2(2).Caption = Percent & "%"
  DoEvents
End Sub

Private Sub Form_Load()
  'Create instances of encryption classes
  Set EncryptSkipJack = New clsSkipjack
  Set EncryptBlowfish = New clsBlowfish
  Set EncryptCryptAPI = New clsCryptAPI
  Set EncryptTwofish = New clsTwofish
  Set EncryptXOR = New clsSimpleXOR
  Set EncryptGost = New clsGost
  Set EncryptTEA = New clsTEA
  Set EncryptRC4 = New clsRC4
  Set EncryptDES = New clsDES
  
  'Add all encryption classes to an
  'internal array for easier access
  Call AddEncryption(EncryptBlowfish, "Blowfish", "http://www.counterpane.com/blowfish.html")
  Call AddEncryption(EncryptCryptAPI, "CryptAPI")
  Call AddEncryption(EncryptDES, "DES (Data Encryption Standard)", "http://csrc.nist.gov/fips/fips46-3.pdf")
  Call AddEncryption(EncryptGost, "Gost", "http://www.jetico.sci.fi/index.htm#/gost.htm")
  Call AddEncryption(EncryptXOR, "Simple XOR", "http://tuath.pair.com/docs/xorencrypt.html")
  Call AddEncryption(EncryptRC4, "RC4", "http://www.rsasecurity.com/rsalabs/faq/3-6-3.html")
  Call AddEncryption(EncryptSkipJack, "Skipjack", "http://csrc.nist.gov/encryption/skipjack-kea.htm")
  Call AddEncryption(EncryptTEA, "TEA, A Tiny Encryption Algorithm", "http://www.cl.cam.ac.uk/Research/Papers/djw-rmn/djw-rmn-tea.html")
  Call AddEncryption(EncryptTwofish, "Twofish", "http://www.counterpane.com/twofish.html")
  
  'Pre-select the first item in the list
  Combo1.ListIndex = 0
End Sub

Function Run(strFilePath As String, Optional strParms As String, Optional strDir As String) As String
  Const SW_SHOW = 5
  
  'Run the Program and Evaluate errors
  Select Case ShellExecute(0, "Open", strFilePath, strParms, strDir, SW_SHOW)
  Case 0
    Run = "Insufficent system memory or corrupt program file"
  Case 2
    Run = "File not found"
  Case 3
    Run = "Invalid path"
  Case 5
    Run = "Sharing or Protection Error"
  Case 6
    Run = "Seperate data segments are required for each task"
  Case 8
    Run = "Insufficient memory to run the program"
  Case 10
    Run = "Incorrect Windows version"
  Case 11
    Run = "Invalid program file"
  Case 12
    Run = "Program file requires a different operating system"
  Case 13
    Run = "Program requires MS-DOS 4.0"
  Case 14
    Run = "Unknown program file type"
  Case 15
    Run = "Windows program does not support protected memory mode"
  Case 16
    Run = "Invalid use of data segments when loading a second instance of a program"
  Case 19
    Run = "Attempt to run a compressed program file"
  Case 20
    Run = "Invalid dynamic link library"
  Case 21
    Run = "Program requires Windows 32-bit extensions"
  Case Else
    Run = ""
  End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub
