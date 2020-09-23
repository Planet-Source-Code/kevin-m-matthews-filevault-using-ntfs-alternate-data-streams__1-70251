VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "File Vault Viewer"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeSettings 
      Caption         =   "Change VaultFile Parameters"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   3375
   End
   Begin VB.PictureBox picProgContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   3345
      TabIndex        =   3
      Top             =   5520
      Width           =   3375
      Begin VB.PictureBox picProgress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   4
         Top             =   0
         Width           =   15
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   60
            Width           =   3375
         End
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   60
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ListBox lstFiles 
      Height          =   4350
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangeSettings_Click()
    Dim x As Integer
    Dim Pass() As Byte
    Dim newPass() As Byte
    Dim tmpStr As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim FileBrk() As String
    
    ' Set the return value of the change
    ' setting form to false
    frmChangeSettingsRet = False
    ' Show the Change Settings Form as modal
    ' so we wait for the user to click OK or
    ' cancel on that form before continuing
    frmChangeSettings.Show vbModal
    ' If the user enter valid info, then the
    ' return value is true
    If frmChangeSettingsRet Then
        ' For each file listed in the main listbox
        For x = 0 To lstFiles.ListCount - 1
            ' Setup the cipher key with the previous
            ' password, key and block sizes
            Pass = GetPassword(PassPhrase)
            m_AES.SetCipherKey Pass, KeySize, BlockSize
            ' Set tmpSTR to the filename for the
            ' current listbox item
            tmpStr = lstFiles.List(x)
            tmpStr = Trim(tmpStr)
            
            ' Split the tmpSTR string using . as a delimiter
            ' FileBrk(0) will hold the name, and FileBrk(1)
            ' will hold the extension
            FileBrk() = Split(tmpStr, ".")
            
            ' Set the strings Temp1 and Temp2 to tmp1 and tmp2
            ' using the FileBrk(1) as the extension
            Temp1 = App.Path & "\tmp1" & "." & FileBrk(1)
            Temp2 = App.Path & "\tmp2" & "." & FileBrk(1)
            
            ' Copy the file out of the ADS to the Temp1 file
            CopyFromADSToFile Temp1, tmpStr
            ' Decrypt temp1 and save enecrypted text as Temp2
            m_AES.FileDecrypt Temp2, Temp1, BlockSize
            
            ' Setup the cipher key with the new
            ' password, key and block sizes
            newPass = GetPassword(NewPassPhrase)
            m_AES.SetCipherKey newPass, NewKeySize, NewBlockSize
            ' Secure delete the Temp1 file. May not be neccessary,
            ' but I want to make sure that there is now way any data
            ' that is currently held in that file could be left in place
            ' when I overwrite it, since it would through off our encryption
            KillIt Temp1
            ' Now encrypt the unencrypted Temp2 file and save
            ' the result as temp1
            m_AES.FileEncrypt Temp2, Temp1, NewBlockSize
            ' Copy Temp1 back into the ADS using it's previous name
            CopyFromFileIntoADS Temp1, VaultPath, tmpStr, True, lstFiles
            ' Secure delete both tempfiles
            KillIt Temp1: KillIt Temp2
        ' Continue the for next loop for all items in the listbox
        Next x
        
        ' Set the strings holding the previous
        ' password, Key and block sizes to the
        ' values the user just set
        PassPhrase = NewPassPhrase
        KeySize = NewKeySize
        BlockSize = NewBlockSize
        ' Setup the cipher key with the new
        ' password, key and block sizes
        Pass = GetPassword(PassPhrase)
        m_AES.SetCipherKey Pass, KeySize, BlockSize
        ' Rewrite the Filestruct using the new
        ' cipher key
        WriteFileStruct lstFiles, VaultPath
    End If
End Sub

Private Sub cmdDelete_Click()
    ' If a file is selected in listbox
    If lstFiles.SelCount > 0 Then
        ' then delete the current file
        DeleteStream VaultPath, lstFiles.List(lstFiles.ListIndex), lstFiles
    ' Otherwise
    Else
        ' Tell the user the error
        MsgBox "Please select a file to delete", vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim openHandle As Long
    Dim lngReturn As Long
    Dim strFile As String
    Dim StrFiles() As String
    Dim FileBrk() As String
    Dim FileDlg As New clsFileDialog
    Dim Pass() As Byte
    
    ' Use FileDialog Class to load a file
    With FileDlg
        .Filter = "All Files(*.*)|*.*"
        .FilterIndex = 1
        .WindowTitle = "Browse for executable"
        strFile = .FileOpen
    End With
    If Len(strFile) > 0 Then
        ' Split strFile using \ as a delimiter
        ' The last item in the array will be the filename e.g. test.exe
        StrFiles() = Split(strFile, "\")
TrySaveAgain:
        ' Is this filename already used in the vault?
        If Not FilenameUsed(StrFiles(UBound(StrFiles)), lstFiles) Then
            ' If not, then add it
            ' First, make sure the cipher key is set
            Pass = GetPassword(PassPhrase)
            m_AES.SetCipherKey Pass, KeySize, BlockSize
            ' Split the Filename using . as a delimiter
            ' FileBrk(0) will contain the name, and FileBrk(1)
            ' will contain the extension
            FileBrk() = Split(StrFiles(UBound(StrFiles)), ".")
            ' Encrypt the file
            m_AES.FileEncrypt strFile, App.Path & FileBrk(0) & "." & FileBrk(1), BlockSize
            ' Copy the file into the active data stream
            CopyFromFileIntoADS App.Path & FileBrk(0) & "." & FileBrk(1), VaultPath, StrFiles(UBound(StrFiles)), False, lstFiles
            ' Secure delete the tempfile
            KillIt App.Path & FileBrk(0) & "." & FileBrk(1)
        Else
            ' If filename is already in use, use get
            ' the extension and save it to a variable
            frmEnterNewNameExtension = GetExtension(Replace(Trim(StrFiles(UBound(StrFiles))), Chr(0), ""))
            ' Launch form to enter new filename
            frmEnterNewName.Show vbModal
            ' If the user input from the new name form is valid
            If frmEnterNewNameRet Then
                ' Replace the old filename with the new one and try the save again
                StrFiles(UBound(StrFiles)) = frmEnterNewNameFilename
                ' jump to the TrySaveAgain label to try it again
                GoTo TrySaveAgain
            End If
        End If
    End If
    ' Destroy the File Dialog
    Set FileDlg = Nothing
End Sub

Private Sub Form_Load()
    ReadFileStruct lstFiles, VaultPath
    gHW = Me.hwnd
    Hook
End Sub

Private Sub Form_Terminate()
    Unhook
End Sub

Private Sub lstFiles_DblClick()
    Dim ret As Long
    Dim tmpStr As String
    Dim Pass() As Byte
    Dim FileBrk() As String
    ' The trim may not be neccessary, I added it when I was having problems,
    ' I can't remember whether this was the fix, or if I had to try something
    ' else and just never took it out
    tmpStr = lstFiles.List(lstFiles.ListIndex)
    tmpStr = Trim(tmpStr)
    
    ' First, make sure the cipher key is set
    Pass = GetPassword(PassPhrase)
    m_AES.SetCipherKey Pass, KeySize, BlockSize
    ' Split tmpSTR using . as delimiter
    ' FileBrk(0) will hold the filname, and
    ' FileBrk(1) will hold the extension
    FileBrk() = Split(tmpStr, ".")
    ' Copy the file from the ADS to a temp file (tmp1.tmp)
    CopyFromADSToFile App.Path & "\tmp1.tmp", tmpStr
    ' Decrypt the tempfile and save it as filename and the extension
    ' by combining FileBrk(0) & "." & FileBrk(1) into a string
    m_AES.FileDecrypt App.Path & "\" & FileBrk(0) & "." & FileBrk(1), App.Path & "\tmp1.tmp", BlockSize
    ' Secure delete the temp file
    KillIt App.Path & "\tmp1.tmp"
    
    
    '-----------------------------------------------------------------
    ' Lazyman's Multithreading (I know it isn't realy  multithreading)
    '-----------------------------------------------------------------
    ' This shells to the SNW (Shell and Wait) exe, in the lpParamaters
    ' string of this ShellExecute Function, all the values needed to
    ' perform another ShellExecute are passed as a string. When the
    ' SNW.exe file loads, it will split the string using ^ as a delimiter.
    ' It will then perform another shell execute using the parameters that
    ' were passed to it. After it calls ShellExecute, the SNW program waits
    ' for the resulting program to end, then it deletes the file we just
    ' decrypted above. All this results the user simply double clicking on
    ' a file in the listbox, a bunch of stuff happens behind the scenes, and
    ' when they are done viewing the file, the unencrypted version is removed.
    ' NOTE: CURRENTLY,ANY CHANGES MADE TO THE FILE WILL NOT BE IMPORTED BACK
    ' INTO THE VAULT. I WILL WORK ON IMPLEMENTING THIS IN THE FUTURE.
    ret = ShellExecute(Me.hwnd, "open", App.Path & "\SNW.exe", "FVSNW^" & Str(Me.hwnd) & "^open^" & App.Path & "\" & FileBrk(0) & "." & FileBrk(1) & "^0&^^1", "", 1)
End Sub


