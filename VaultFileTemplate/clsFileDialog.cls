VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is used in place of common dialog control
' From this point on, all comments are by original author
'----------------------------------------------------------
' This class use API calls to use the Open and Save common dialog from the "Microsoft Common Dialog Control". This class also adds
' function to extract filenames when the multiselect is on. It still doesn't have all functions from the dialog control, but all useful
' functions are included. I will not be held responsible for any damage this could have done to your computer. But if it happens,
' please tell me right now, I'm always interested in paranormal phenomenons :)

' Made by Insomniaque :)

'modified by Dj's Computer Labs 10/22/99
    'I have not tested the GetMultiFilename function with my changes
    
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private FileDialog As OPENFILENAME

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Sub Class_Initialize()
  With FileDialog
    .hInstance = App.hInstance
    .hwndOwner = hwnd
    .lpstrFile = Chr(0) & Space(254)
    .nMaxFile = 255
    .lStructSize = Len(FileDialog)
  End With
End Sub

Public Property Get FilterIndex() As Long
  FilterIndex = FileDialog.nFilterIndex
End Property

Public Property Let FilterIndex(ByVal Index As Long)
  FileDialog.nFilterIndex = Index
End Property

Public Property Get Flags() As Long
  Flags = FileDialog.Flags
End Property

Public Property Let Flags(ByVal vFlags As Long)
  FileDialog.Flags = vFlags
End Property

Public Property Let DefaultExtension(ByVal Extention As String)
    ' Default extension for files
    FileDialog.lpstrDefExt = Extention
End Property

Public Property Get DefaultExtension() As String
    ' Default extension for files
    DefaultExtension = FileDialog.lpstrDefExt
End Property

Public Property Let ObjectOwner(Objet As Object)
    ' Object that 'owns' the common dialog.
    FileDialog.hwndOwner = Objet.hwnd
End Property

Public Property Let Filter(ByVal CustomFilter As String)
    ' File filter. Format : "Text for filter|filter|Next Text|Next filter". Ex : "All Files (*.*)|*.*|Text files (*.txt)|*.txt"
    Dim intCount As Integer
    FileDialog.lpstrFilter = ""
    For intCount = 1 To Len(CustomFilter)
        If Mid(CustomFilter, intCount, 1) = "|" Then FileDialog.lpstrFilter = FileDialog.lpstrFilter + Chr(0) Else FileDialog.lpstrFilter = FileDialog.lpstrFilter + Mid(CustomFilter, intCount, 1)
    Next intCount
    FileDialog.lpstrFilter = FileDialog.lpstrFilter + Chr(0)
End Property

Public Property Let WindowTitle(ByVal Title As String)
    ' Window title
    FileDialog.lpstrTitle = Title
End Property

Public Property Get WindowTitle() As String
    ' Window title
    WindowTitle = FileDialog.lpstrTitle
End Property

Public Property Let InitialDirectory(ByVal InitDir As String)
    ' Initial folder
    FileDialog.lpstrInitialDir = InitDir
End Property

Public Property Get InitialDirectory() As String
    ' Initial folder
    InitialDirectory = FileDialog.lpstrInitialDir
End Property

Public Function FileOpen() As String
  ' Show common dialog open file control and returns file name.
  'Use GetMultiFilename function to  extract filenames when the multiselect flag is true. The Count function returns the number of files selected.
  Dim lngReturn As Long
  
  lngReturn = GetOpenFileName(FileDialog)
  
  If lngReturn >= 1 Then
    FileOpen = FileDialog.lpstrFile
  End If

End Function

Public Function FileSave() As String
    ' Show common dialog save file control and returns file name.
  Dim lngReturn As Long
  
  lngReturn = GetSaveFileName(FileDialog)
  
  If lngReturn >= 1 Then
    FileSave = FileDialog.lpstrFile
  End If
 
End Function

'Public Function Count() As Integer
'    ' Returns the number of file selected. Use with GetMultiFilename to extract filename when multiselect is true.
'    Dim intCount As Integer
'    For intCount = 1 To Trim(Len(FileDialog.lpstrFile))
'        If Mid(Trim(FileDialog.lpstrFile), intCount, 1) = Chr(0) Then Count = Count + 1
'    Next intCount
'    Count = Count - 2
'    If Count < 1 Then Count = Count + 1
'End Function
'
'Public Function GetMultiFilename(ByVal Filenumber As Integer) As String
'    ' Returns the filename of the specified filenumber. Use only with open file dialog when multiselect is true.
'    Dim intCount As Integer
'    Dim intOne As Integer
'    Dim intFile As Integer
'    Dim intNext As Integer
'    intOne = InStr(1, Trim(FileDialog.lpstrFile), Chr(0))
'    intFile = 1
'    For intCount = 1 To Filenumber
'        intFile = InStr(intFile + 1, Trim(FileDialog.lpstrFile), Chr(0))
'    Next intCount
'    intNext = InStr(intFile + 1, Trim(FileDialog.lpstrFile), Chr(0))
'    GetMultiFilename = IIf(Right(Mid(Trim(FileDialog.lpstrFile), 1, intOne - 1), 1) = "\", Mid(Trim(FileDialog.lpstrFile), 1, intOne - 1), Mid(Trim(FileDialog.lpstrFile), 1, intOne - 1) + "\") + Mid(Trim(FileDialog.lpstrFile), intFile + 1, intNext - intFile - 1)
'    If Right(GetMultiFilename, 1) = "\" Then GetMultiFilename = Left(GetMultiFilename, Len(GetMultiFilename) - 1)
'End Function





