Attribute VB_Name = "modFileAttributes"
' From posting on PSC by Rde
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56686&lngWId=1

' The return value is the sum of the attribute values
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function GetAttributes Lib "kernel32" _
    Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long

' Sets the Attributes argument whose sum specifies file attributes
' An error occurs if you try to set the attributes of an open file
Public Declare Function SetAttributes Lib "kernel32" _
    Alias "SetFileAttributesA" (ByVal lpSpec As String, _
    ByVal dwAttributes As Long) As Long

Public Enum vbFileAttributes
  vbNormal = 0         ' Normal
  vbReadOnly = 1       ' Read-only
  vbHidden = 2         ' Hidden
  vbSystem = 4         ' System file
  vbVolume = 8         ' Volume label
  vbDirectory = 16     ' Directory or folder
  vbArchive = 32       ' File has changed since last backup
  vbTemporary = &H100  ' 256
  vbCompressed = &H800 ' 2048
End Enum

Public Function GetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes) As Boolean
  ' Returns True if the specified attribute(s) is currently set.
  If (LenB(sFileSpec) <> 0) Then
    GetAttrib = (GetAttributes(sFileSpec) And Attrib) = Attrib
  End If
End Function


Public Sub SetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes, Optional fTurnOff As Boolean)
  ' Sets/clears the specified attribute(s) without affecting other attributes. You
  ' do not need to know the current state of an attribute to set it to on or off.
  If (LenB(sFileSpec) <> 0) Then
    If (Attrib = vbNormal) Then
      SetAttributes sFileSpec, vbNormal
    ElseIf fTurnOff Then
      SetAttributes sFileSpec, GetAttributes(sFileSpec) And (Not Attrib)
    Else
      SetAttributes sFileSpec, GetAttributes(sFileSpec) Or Attrib
    End If
  End If
End Sub
