Attribute VB_Name = "ModReadAndWriteADS"
'This code is largely based on an article at http://allapi.mentalis.org/apilist/CopyFileEx.shtml#
' This code simply implements the methods discussed in that article by using VB code instead
' of running command like tpye and echo at the dos prompt

Option Explicit
Public Const PROGRESS_CANCEL = 1        '
Public Const PROGRESS_CONTINUE = 0
Public Const PROGRESS_QUIET = 3
Public Const PROGRESS_STOP = 2
Public Const COPY_FILE_FAIL_IF_EXISTS = &H1
Public Const COPY_FILE_RESTARTABLE = &H2
Public bCancel As Long
Public Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long



Public Function CopyFromFileIntoADS(OriginalFile As String, AppendADSTo As String, StreamName As String, TheFileStruct As Boolean, Optional List As ListBox)
    Dim ret As Long
    ' Copy the file we want to add (OriginalFile) to the
    ' vault location (AppendADSTo & ":" & StreamName).
    ' The AddressOf CopyProgressRoutine causes the
    ' CopyFileEx function to call this routine, which in
    ' turn updates the file copy progress and then tells
    ' the CopyFileEx function to continue
    ret = CopyFileEx(OriginalFile, AppendADSTo & ":" & StreamName, AddressOf CopyProgressRoutine, ByVal 0&, bCancel, COPY_FILE_RESTARTABLE)
    'list.Clear
    ' Now we need to add this file to the FStruct.VLT file
    'Open AppendADSTo & ":" & "FStruct.VLT" For Append As #1
    '    Print #1, vbCrLf
    '    Print #1, StreamName
    'Close #1
    If Not TheFileStruct Then
        List.AddItem StreamName
        WriteFileStruct List, VaultPath
    End If
    'ReadFileStruct list, AppendADSTo
End Function

' This function is called by the CopyFileEx function
Public Function CopyProgressRoutine(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, ByVal lpData As Long) As Long
    ' I got this routine from a vb api site
    ' I had now idea the CopyFileEx routine
    ' had a variable for assigning functions
    ' to handle the progress updates.
    Dim tmpProgress As Integer
    'adjust the caption
    tmpProgress = Int((TotalBytesTransferred * 10000) / (TotalFileSize * 10000) * 100)
    'allow user input
    DoEvents
    'continue filecopy
    CopyProgressRoutine = PROGRESS_CONTINUE
    On Error Resume Next
    'frmMain.picProgress.Width = (frmMain.picProgContainer.Width * tmpProgress) / 100
    'frmMain.lblStat(0).Caption = CStr(tmpProgress & "%")
    'frmMain.lblStat(1).Caption = CStr(tmpProgress & "%")
    'frmMain.picProgress.ToolTipText = CStr(tmpProgress & "%")
End Function

Public Function DeleteStream(AppendedTo As String, StreamName As String, List As ListBox)
    ' Delete the file using delete file API
    DeleteFile AppendedTo & ":" & StreamName
    ' Remove the files entry from the list box
    List.RemoveItem List.ListIndex
    ' Update the FileStruct file based to match the listbox
    WriteFileStruct List, AppendedTo
    List.Clear
    ' Reload the filestructure
    ReadFileStruct List, AppendedTo
End Function

Public Function CopyFromADSToFile(NewFile As String, StreamName As String, Optional List As ListBox)
    Dim ret As Long
    Dim SourceFile As Integer
    Dim DestFile As Integer
    Dim SrcFileLen As Long
    Dim Chunk As String
    Dim BytesToGet As Integer
    Dim BytesCopied As Long
    
    'How many bytes to get each time
    BytesToGet = 4096 '4kb
    BytesCopied = 0
    SourceFile = 1
    DestFile = 2
    
    
    ' Copy the file from the ADS stream
    ' to a location specified by NewFile
    Open VaultPath & ":" & StreamName For Binary As SourceFile
        SrcFileLen = LOF(SourceFile)
        Open NewFile For Binary As DestFile
            Do While BytesCopied < SrcFileLen
                'Check how many bytes left
                If BytesToGet < (SrcFileLen - BytesCopied) Then
                    'Copy 4 KBytes
                    Chunk = Space(BytesToGet)
                    Get #SourceFile, , Chunk
                Else
                    'Copy the rest
                    Chunk = Space(SrcFileLen - BytesCopied)
                    Get #SourceFile, , Chunk
                End If
                BytesCopied = BytesCopied + Len(Chunk)
    
       
                'Put data in destination file
                Put #DestFile, , Chunk
            Loop
        Close DestFile
    Close SourceFile
End Function
