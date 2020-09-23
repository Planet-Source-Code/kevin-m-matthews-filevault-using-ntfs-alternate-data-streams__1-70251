Attribute VB_Name = "modGeneral"
' This file contains some public functions
' The ChrCode function was not written by me,
' I got it from an HTML parser by Kristian. S.Stangeland.
' It isn't actually used in this program, but I use it
' when I am debugging, so I left it in.

Option Explicit
Private Const PROCESS_QUERY_INFORMATION = &H400
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public m_AES As New clsAES
Public SDS As New clsSDS


Public Function ReadFileStruct(List As ListBox, Fil As String)
    'On Error GoTo FStructLoadError
    Dim tmpStr As String
    Dim Pass() As Byte
    Dim FirstInp As Boolean
    Pass() = GetPassword(PassPhrase)
    m_AES.SetCipherKey Pass, KeySize, BlockSize
    ' clear the listbox
    'list.Clear
    ' check to see if the file exists
TryAgain:
    If FileExists(Fil) Then
        ' Copy the FStruct.VLT file from the ADS to a temp file
        CopyFromADSToFile App.Path & "\tmp1.tmp", "FStruct.VLT"
        'Decrypt the tempfile to another tempfile
        m_AES.FileDecrypt App.Path & "\tmp2.tmp", App.Path & "\tmp1.tmp", 128
        ' open the vaults decrypted file structure
        Open App.Path & "\tmp2.tmp" For Input As #1
            ' read every item in the file
            ' The FirstInp boolean determines if the
            ' loop is running through for the first time
            ' If it is, then the first line pulled from the
            ' file should match the passphrase
            FirstInp = True
            Do While Not EOF(1)
                Input #1, tmpStr
                If FirstInp Then
                    ' If tmpSTR doesnmatch the pasphrase
                    ' The the wrong password was entered
                    If tmpStr <> ">" & PassPhrase Then
                        Close #1
                        'If Dir(App.Path & "\tmp1.tmp") <> "" Then
                        KillIt App.Path & "\tmp1.tmp"
                        'If Dir(App.Path & "\tmp2.tmp") <> "" Then
                        KillIt App.Path & "\tmp2.tmp"
                        MsgBox "Incorrect Passphrase", vbOKOnly
                        End
                    End If
                    FirstInp = False
                End If
                If Len(tmpStr) > 0 And Left(tmpStr, 1) <> ">" Then
                    ' add each item in this file to the list
                    List.AddItem tmpStr
                End If
            Loop
            'list.RemoveItem list.ListCount - 1
        Close #1
    Else
        Fil = OpenFile
        If Fil <> "<>" Then
            ' update the vaulpath string
            VaultPath = Fil
            GoTo TryAgain
        End If
    End If
    ' delete the tempfiles. Eventually, I will need to implement a secure delete,
    ' or do all of the tempfile stuff in RAM using virtual files.
    KillIt App.Path & "\tmp1.tmp": KillIt App.Path & "\tmp2.tmp"
    Exit Function
FStructLoadError:
    'If you get to this point in the function, then there was an error reading the file structure
    ' This mostly likely means that the file you specified at startup does not have a file structure
    ' file associated with it, so lets make one
    Debug.Print Error
    ' Create blank file structure
    Open App.Path & "\tmp1.tmp" For Output As #1
    'Fil & ":FStruct.VLT" For Output As #1
        Print #1, ">Banana"
    Close #1
    m_AES.FileEncrypt App.Path & "\tmp1.tmp", App.Path & "\tmp2.tmp", 128
    CopyFromFileIntoADS App.Path & "\tmp2.tmp", VaultPath, "FDtruct.VLT", False
End Function

Public Function WriteFileStruct(List As ListBox, Fil As String)
    Dim tmpStr As String
    Dim x As Integer
    Dim Pass() As Byte
    Pass() = GetPassword(PassPhrase)
    m_AES.SetCipherKey Pass, KeySize, BlockSize
    ' open the vaults file structure
    Open App.Path & "\tmp1.tmp" For Output As #1
        tmpStr = ">" & PassPhrase
        Print #1, tmpStr
        For x = 0 To List.ListCount
            tmpStr = List.List(x)
            ' print each item in listbox to file structure file
            Print #1, tmpStr
        Next x
    Close #1
    m_AES.FileEncrypt App.Path & "\tmp1.tmp", App.Path & "\tmp2.tmp", BlockSize
    CopyFromFileIntoADS App.Path & "\tmp2.tmp", VaultPath, "FStruct.VLT", True
    KillIt App.Path & "\tmp1.tmp": KillIt App.Path & "\tmp2.tmp"
End Function

' This function is called when we add a file
' It searches the list box and sees if the filename is in use
' The trim and other string manipulation trims any extra spaces,
' and chr(0)s to ensure we can compare the filename
Public Function FilenameUsed(FileName As String, List As ListBox) As Boolean
    Dim x As Integer
    FilenameUsed = False
    ' Trim off leading and trailing spaces
    FileName = Trim(FileName)
    ' This keeps repeating until we get rid of the chr(0)
    ' that remains after we trim of the extra spaces
    Do While Right(FileName, 1) = Chr(0)
        FileName = Left(FileName, Len(FileName) - 1)
    Loop
    ' Loop through the list and see if filename is in use
    For x = 0 To List.ListCount - 1
        If UCase(List.List(x)) = UCase(FileName) Then
            ' Function returns true if the name is in use
            FilenameUsed = True
            Exit Function
        End If
    Next x
    ' Function returns false if the name is not in use
    FilenameUsed = False
End Function


Public Function ContainsIllegalChars(ByVal sString As String) As Boolean
    Const IllegalChars As String = "\/:*?""<>|" ' The illegal chars
    Dim sResult As String
    Dim i As Single
    sResult = sString
    
    ' Cycle through each char in the IllegalChars string
    For i = 1 To Len(IllegalChars)
        ' Check to see if the charachter at position i of
        ' illegal charachter screen is in filename
        If InStr(1, sString, Mid$(IllegalChars, i, 1)) Then
            ' If it is, then the function returns true
            ContainsIllegalChars = True
            Exit Function
        End If
    Next i
    ' If no illegal chars then the function returns fals
    ContainsIllegalChars = False
End Function

Public Function GetExtension(FileName As String) As String
    Dim tmpStr() As String
    ' splits the filename into an array delimted by .
    tmpStr() = Split(FileName, ".")
    ' The last element in array is the extension,
    ' so we set the extension to equal . & the value
    ' of the last element in the array
    GetExtension = "." & tmpStr(UBound(tmpStr))
End Function

Public Function OpenFile() As String
    Dim openHandle As Long
    Dim lngReturn As Long
    Dim strFile As String
    Dim StrFiles() As String
    Dim FileDlg As New clsFileDialog
  
    ' Use FileDialog Class to load a file
    With FileDlg
        .Filter = "All Files(*.*)|*.*"
        .FilterIndex = 1
        .WindowTitle = "Browse for executable"
        strFile = .FileOpen
    End With
    If Len(strFile) > 0 Then
        strFile = Replace(Trim(strFile), Chr(0), "")
        OpenFile = strFile
    Else
        OpenFile = "<>"
    End If
End Function

Public Sub OldMain()
    Dim Fil As String
    ' If you launch a Vault File (one created with this program)
    ' it will send the Vault Files path to this program before terminating
    If Command$ <> "" Then
        ' Set the path to the VaultFile
        VaultPath = Trim(Command$)
TryAgain1:
        ' If this is not a vaild path
        If Not FileExists(VaultPath) Then
            ' Use Common Dialog Class to open a file
            Fil = OpenFile
            ' If the user selected a valid fil in common dialog control
            If Fil <> "<>" Then
                ' update the vaulpath string
                VaultPath = Fil
                ' Try validating path again
                GoTo TryAgain1
            End If
        End If
        ' Ask the user to enter passphrase
        ' and Key/Block Sizes
        frmLogin.Show vbModal
        ' If the entered information was correct
        If frmLoginRet Then
            ' Load the Vault File's File Structure
            ReadFileStruct frmMain.lstFiles, VaultPath
            Load frmMain
            frmMain.Show
        Else
            End
        End If
        ' Load the Vault File's File Structure
        'ReadFileStruct frmMain.lstFiles, VaultPath
    ' otherwise, we will have to open a test file for demo purposes
    Else
        ' !!!!! IMPORTANT!!!!!!!!
        VaultPath = App.Path & "\VaultFile.exe"
TryAgain2:
        ' If this is not a vaild path
        If Not FileExists(VaultPath) Then
            ' Use Common Dialog Class to open a file
            Fil = OpenFile
            ' If the user selected a valid fil in common dialog control
            If Fil <> "<>" Then
                ' update the vaulpath string
                VaultPath = Fil
                ' Try validating path again
                GoTo TryAgain2
            End If
        End If
        ' Ask the user to enter passphrase
        ' and Key/Block Sizes
        frmLogin.Show vbModal
        ' If the entered information was correct
        If frmLoginRet Then
            ' Load the Vault File's File Structure
            ReadFileStruct frmMain.lstFiles, VaultPath
            Load frmMain
            frmMain.Show
        Else
            End
        End If
    End If
End Sub

Sub Main()
    Dim Fil As String
    FirstRun = False
    Dim tmpStr As String
    ' If you launch a Vault File (one created with this program)
    ' it will send the Vault Files path to this program before terminating
    If Command$ <> "" Then
        ' Set the path to the VaultFile
        VaultPath = Trim(Command$)
TryAgain1:
        ' If the command line parameters link to a valid file
        If FileExists(VaultPath) Then
            If Not FileExists(VaultPath & ":FStruct.VLT") Then
                ' If it does not have a FileVault file structure
                ' then we will make one
                GoTo CreateFileStruct
            End If

            ' Allow user to enter passphrase,
            ' Key, and Block Size
            frmLogin.Show vbModal
            ' If the user entered valid info
            If frmLoginRet Then
                ' Load the file
                frmMain.Show
            Else
                End
            End If
        Else
            Fil = OpenFile
            ' If the user selected a valid fil in common dialog control
            Debug.Print Fil
            If Fil <> "<>" Then
                ' update the vaultpath string
                VaultPath = Fil
                Debug.Print Fil
                ' Try validating path again
                GoTo TryAgain1
            End If
        End If
    Else
        VaultPath = "<>><>"
        ' If the command line parameters link to a valid file
        If FileExists(VaultPath) Then
        'tmpSTR = Dir(VaultPath)
TryAgain2:
            ' Allow user to enter passphrase,
            ' Key, and Block Size
            frmLogin.Show vbModal
            ' If the user entered valid info
            If frmLoginRet Then
                ' Load the file
                frmMain.Show
            Else
                End
            End If
        Else
            Fil = OpenFile
            ' If the user selected a valid fil in common dialog control
            If Fil <> "<>" Then
                ' update the vaultpath string
                VaultPath = Fil
                ' Now that the user has selected a file, see if that
                ' File has a FileVault file structure
                If FileExists(VaultPath & ":FStruct.VLT") Then
                    ' If it does then jump back up to the
                    ' TryAgain2 label and continue on
                    GoTo TryAgain2
                Else
                    ' If it does not have a FileVault file structure
                    ' then we will make one
                    GoTo CreateFileStruct
                End If
            ' If the user did not select a file (e.g clicked cancel)
            Else
                ' The end the program
                End
            End If
        End If
    End If
    Exit Sub
CreateFileStruct:
        ' Set this to false before we continue
        frmLoginRet = False
        ' Tell the user what is going on
        MsgBox "This is the first time this VaultFile has been used." & vbNewLine & _
        "Please enter the Passphrase, as well as the Keysize and" & vbNewLine & _
        "Bytesize you would like to use for this VaultFile.", vbInformation, "New VaultFile"
        ' Set this to true so that after the user enters a password
        ' and sets the other parameters on the form we are about to
        ' we process this as a first run
        FirstRun = True
        ' Show the form
        frmLogin.Show vbModal
        ' If it returns true then we know the user entered a valid
        ' info
        If frmLoginRet Then
            FirstRun = False
            frmMain.Show
        Else
            ' If the user hit cancel, the quit the program
            End
        End If
End Sub

Public Function GetPassword(Phrase As String) As Byte()
    Dim Data() As Byte
    'convert the passphrase string to an array of bytes
    Data = StrConv(Phrase, vbFromUnicode)
    ReDim Preserve Data(31)

    GetPassword = Data
End Function

Private Function ChrCode(txt As String) As String

    Dim x
    If Len(txt) <= 0 Then Exit Function
    Dim outstring As String


    For x = 1 To Len(txt$)
        outstring$ = outstring$ + "Chr(" + CStr(Asc(Mid(txt$, x, 1))) + ") + "
    Next x

    outstring$ = Trim(outstring$)
    outstring$ = Mid(outstring$, 1, Len(outstring$) - 2)
    ChrCode = outstring$
End Function

Public Function KillIt(sPath As String)
    If SDS.File_Exists(sPath) Then
        With SDS
            '/* reset attributes to normal
            .p_Attributes = 1
            '/* number of passes
            .p_Passes = 4
            '/* file path
            .p_SourceFile = sPath
            '/* core
            .File_Shred
        End With
        SDS.File_Shred
    End If
End Function

Private Function FileExists(FileName As String) As Boolean
    ' If we get an error with the open Filename statement,
    ' which is the only statement in this function that
    ' could possibly cause an error, then we jump to the
    ' file does not exist label
    On Error GoTo FileNotExist
        Open FileName For Input As #1
        Close #1
        'If we get to this point, then the file exists
        FileExists = True
        ' Exit the function
        Exit Function
FileNotExist:
        'If we get to this point, then the file DOES NOT exist
        FileExists = False
        ' Not sure this is neccessary. Since the open statement
        ' returned an error, the file may not actually be open,
        ' but everything works fine with it in, so I will leave
        ' it. Any definitve answer on this?
        Close #1
End Function
