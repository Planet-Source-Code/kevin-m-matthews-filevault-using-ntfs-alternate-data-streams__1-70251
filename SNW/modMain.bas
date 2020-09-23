Attribute VB_Name = "modMain"
' The shell execute and wait code is based largely on someone elses code. I apologize, but
' the author gets to be the one person I can't provide proper credit to. If you wrote
' this, let me know and I will provide credit.
' Basically all this program does is launch, read the values passed to it by the main program,
' in the command line, parses the commands, and runs a shell execute. After executing, the program
' gets the handle to the program that just launched, then it waits for that program to terminate.
' When it does terminate, this program performs a secure delete on the unencrypted file that was
' being shelled to, then it terminates itself.


Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


' These are some constant values used for the
' ShellExecuteWait function
Private Const INFINITE As Long = &HFFFFFFFF
Private Const SEE_MASK_FLAG_NO_UI As Long = &H400
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40

' Defines a type that contains information about the
' process that we launch
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetLastError Lib "Kernel32.dll" () As Long
Private Declare Function ShellExecuteEx Lib "Shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
' Declare the Steppenwolfe's Secure Document Shredder class
Private SDS As New clsSDS
Private fso As New FileSystemObject
Private OldModDT As String
Private NewModDT As String

Sub Main()
    Dim Commands() As String
    Dim lReturn As Long
    ' The only way a valid execution of this program will occur,
    ' is if the main program passes it the correct parameters
    ' First we check to make sure command line parameters were
    ' passed.
    If Command$ <> "" Then
        ' Next we split the command line parameters using ^ as
        ' a delimiter
        Commands() = Split(Command$, "^")
        ' If the Commands(0) <>"FVSNW" then this was not vaild
        ' execution called by the main FileVaultViewer program
        If Commands(0) = "FVSNW" Then
            ' Get the DateLastModified attribute of the file
            OldModDT = GetLastModDate(Commands(3))
            'Call ShellExecutWait, passing the command line parameters as the
            ' values for the various variables in the function
            lReturn = ShellExecuteWait(Commands(1), Commands(2), Commands(3), Commands(4), Commands(5), Commands(6))
            ' Now that the user has closed the file, get the
            ' DateLastModified attribute again
            NewModDT = GetLastModDate(Commands(3))
            ' If the new DateLastModofied and the old one are
            ' the same then the user didn't make any changes
            If NewModDT = OldModDT Then
                ' Secure delete the temp file
                KillIt Commands(3)
            Else
                ' If they are different, then the user did make
                ' changes, so lets tell the mother program the
                ' situation. After the mother program is told
                ' what is going on, it will handle the deletion
                ' of the temp file, so this program can quit now.
                TellMommyWhatHappened Commands(1), Commands(3)
            End If
        End If
    Else
        ' Tell the user they can't just run this program by itself
        MsgBox "This program cannot be run independently.", vbExclamation, "SNW"
    End If
    End
End Sub
Private Function ShellExecuteWait(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Dim lReturn As Long, lResult As Long
    Dim tExecuteInfo As SHELLEXECUTEINFO
    'Fill the SHELLEXECUTEINFO structure
    tExecuteInfo.cbSize = Len(tExecuteInfo)
    tExecuteInfo.fMask = SEE_MASK_NOCLOSEPROCESS
    tExecuteInfo.hwnd = hwnd
    tExecuteInfo.lpVerb = lpOperation
    tExecuteInfo.lpFile = lpFile
    tExecuteInfo.lpParameters = lpParameters
    tExecuteInfo.lpDirectory = lpDirectory
    tExecuteInfo.nShow = nShowCmd
    
    'Call the API with the specified parameters
    lReturn = ShellExecuteEx(tExecuteInfo)
    If lReturn = 0 Then lReturn = GetLastError Else lReturn = tExecuteInfo.hInstApp
    
    'If there's a new process wait while it terminates
    If tExecuteInfo.hProcess <> 0 Then
        lResult = WaitForSingleObject(tExecuteInfo.hProcess, INFINITE)
    End If

    'Return the ShellExecuteEx return value
    ShellExecuteWait = lReturn
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

Function GetLastModDate(Filename As String) As String
    ' Use the FileSystemObject to get the date last modified
    GetLastModDate = fso.GetFile(Filename).DateLastModified
End Function


' This function sends a message to the mother program to let it know that the file that was
' opened was changed by the user. It is a slighlty modified version of Roger D. Taylor's
' Inter-Process Communication code. http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=8899&lngWId=1
Function TellMommyWhatHappened(ByVal Hndl As Long, Path As String)
    Dim cdCopyData As COPYDATASTRUCT
    Dim byteBuffer(1 To 255) As Byte
    Dim strTemp As String
    Dim TMPSTR() As String
    
    TMPSTR() = Split(Path, "\")
    
    strTemp = "FVSNW^" & Path & "^" & TMPSTR(UBound(TMPSTR))
    
    ' Copy the string into a byte array, converting it to ASCII
    Call CopyMemory(byteBuffer(1), ByVal strTemp, Len(strTemp))
    cdCopyData.dwData = 3
    cdCopyData.cbData = Len(strTemp) + 1
    cdCopyData.lpData = VarPtr(byteBuffer(1))
    i = SendMessage(Hndl, WM_COPYDATA, 0&, cdCopyData)
End Function
