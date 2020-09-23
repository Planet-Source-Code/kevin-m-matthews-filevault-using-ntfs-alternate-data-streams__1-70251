Attribute VB_Name = "modVaultFile"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub Main()
    Dim VaultPath As String
    Dim FileVaultPath As String
    Dim ret As Integer
    
    '! CHANGE FILE PATH
    FileVaultPath = App.Path & "\FileVault.exe"
TryAgain:
    If Not (Dir(FileVaultPath) = "") Then
        VaultPath = App.Path
        If Right(VaultPath, 1) <> "\" Then VaultPath = VaultPath & "\"
        VaultPath = VaultPath & App.EXEName & ".exe"
        ret = ShellExecute(0&, "open", FileVaultPath & " ", VaultPath, "", 1)
        If ret <= 32 Then
            MsgBox "Error in Vault File"
        End If
    Else
        Dim openHandle As Long
        Dim lngReturn As Long
        Dim strFile As String
        Dim StrFiles() As String
        Dim FileDlg As New clsFileDialog
  
        ' Use FileDialog Class to load a file
        With FileDlg
            .Filter = "FileVault(*.exe)|*.exe"
            .FilterIndex = 1
            .WindowTitle = "Browse for executable"
            strFile = .FileOpen
        End With
        If Len(strFile) > 0 Then
            FileVaultPath = strFile
            GoTo TryAgain
        End If
    End If
        ' Destroy the File Dialog
    Set FileDlg = Nothing
    
    End
End Sub
