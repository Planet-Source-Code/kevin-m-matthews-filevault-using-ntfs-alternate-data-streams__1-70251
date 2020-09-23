VERSION 5.00
Begin VB.Form frmChangeSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change VaultFile Settings"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
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
      ItemData        =   "frmChangeSettings.frx":0000
      Left            =   1560
      List            =   "frmChangeSettings.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1005
      Width           =   1215
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
      ItemData        =   "frmChangeSettings.frx":0004
      Left            =   3960
      List            =   "frmChangeSettings.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1005
      Width           =   1215
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
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
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
      TabIndex        =   7
      Top             =   1080
      Width           =   1050
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
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmChangeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmChangeSettingsRet = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ' If the user entered a passphrase
    If txtPassphrase.Text <> "" Then
        ' Set the NewPassphrase, NewKeySize,
        ' and NewByteSize values accordingling, these will
        ' values will be transfered over to the Passphrase,
        ' KeySize, and ByteSize variables later on
        NewPassPhrase = txtPassphrase.Text
        NewKeySize = cboKeySize.ItemData(cboKeySize.ListIndex)
        NewBlockSize = cboBlockSize.ItemData(cboBlockSize.ListIndex)
        ' Set return value to true
        frmChangeSettingsRet = True
        ' Unload this form
        Unload Me
    Else
        ' Tell the user what they did wrong
        MsgBox "Please enter a passhrase.", vbExclamation, "Error"
    End If
End Sub

Private Sub Form_Load()
    ' This just loads the combo boxes with
    ' the correct values
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
