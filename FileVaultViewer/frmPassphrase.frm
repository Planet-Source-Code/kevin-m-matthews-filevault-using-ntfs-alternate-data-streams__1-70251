VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Enter VaultFile Parameters"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtPassphrase 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.ComboBox cboKeySize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmPassphrase.frx":0000
      Left            =   3960
      List            =   "frmPassphrase.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1000
      Width           =   1215
   End
   Begin VB.ComboBox cboBlockSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmPassphrase.frx":0004
      Left            =   1560
      List            =   "frmPassphrase.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1000
      Width           =   1215
   End
   Begin VB.Label lblPassphrase 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Passphrase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblBlockSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " BlockSize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Key Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   1050
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmLoginRet = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim tmpStr As String
    ' Set the passphrase, key and block size to
    ' whatever values the user provided
    If txtPassphrase.Text <> "" Then
        PassPhrase = txtPassphrase.Text
        KeySize = cboKeySize.ItemData(cboKeySize.ListIndex)
        BlockSize = cboBlockSize.ItemData(cboBlockSize.ListIndex)
    Else
        MsgBox "You must enter a passphrase to continue.", vbInformation, "Error"
        Exit Sub
    End If
    If Not FirstRun Then
        frmLoginRet = CheckPassphrase(PassPhrase, KeySize, BlockSize)
        Unload Me
    Else
        Open App.Path & "\tmp1.tmp" For Output As #1
            tmpStr = ">" & PassPhrase
            Print #1, tmpStr
        Close #1
        Dim Pass() As Byte
        ' Setup the cipher key
        Pass = GetPassword(PassPhrase)
        m_AES.SetCipherKey Pass, KeySize, BlockSize
        ' Encrypt the file
        m_AES.FileEncrypt App.Path & "\tmp1.tmp", App.Path & "\tmp2.tmp", BlockSize
        CopyFromFileIntoADS App.Path & "\tmp2.tmp", VaultPath, "FStruct.VLT", True
        KillIt App.Path & "\tmp1.tmp": KillIt App.Path & "\tmp2.tmp"
        frmLoginRet = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ' This just fills the combo boxes
    ' with valid data
    cboBlockSize.AddItem "128 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 128
    cboBlockSize.AddItem "160 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 160
    cboBlockSize.AddItem "192 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 192
    cboBlockSize.AddItem "224 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 224
    cboBlockSize.AddItem "256 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 256
    cboKeySize.AddItem "128 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 128
    cboKeySize.AddItem "160 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 160
    cboKeySize.AddItem "192 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 192
    cboKeySize.AddItem "224 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 224
    cboKeySize.AddItem "256 Bit"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 256
    cboBlockSize.ListIndex = 0
    cboKeySize.ListIndex = 0
End Sub

Private Function CheckPassphrase(PassPhrase As String, KeyBits As Long, BlockBits As Long) As Boolean
    Dim Pass() As Byte
    ' Setup the cipher key
    Pass = GetPassword(PassPhrase)
    m_AES.SetCipherKey Pass, KeyBits, BlockBits
    ' Extract the FStruct.VLT file from the ADS
    CopyFromADSToFile App.Path & "\tmp1.tmp", "FStruct.VLT"
    ' If it decrypts correctly
    If m_AES.FileDecrypt(App.Path & "\tmp2.tmp", App.Path & "\tmp1.tmp", BlockBits) = 0 Then
        ' Then the return value is true
        CheckPassphrase = True
        ' Secure delete the temp files
        KillIt App.Path & "\tmp1.tmp": KillIt App.Path & "\tmp2.tmp"
    Else
        ' Return value gets set to false
        CheckPassphrase = False
        ' Kill the one tempfile
        KillIt App.Path & "\tmp1.tmp"
    End If
End Function
