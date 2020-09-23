Attribute VB_Name = "modIPC"
' This module hooks the WndProc function, and replaces the code. It monitors for messages
' sent by the shell and wait apps. It is a slighlty modified version of Roger D. Taylor's
' Inter-Process Communication code. http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=8899&lngWId=1

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Global lpPrevWndProc As Long
Global gHW As Long

'Copies a block of memory from one location to another.
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub Hook()
    ' Setup the hook
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call InterProcessComms(lngParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lngParam)
End Function


Sub InterProcessComms(lngParam As Long)
          Dim cdCopyData As COPYDATASTRUCT
          Dim byteBuffer(1 To 255) As Byte
          Dim strTemp As String
          Dim tmpStr() As String
          Dim Pass() As Byte
          
          Call CopyMemory(cdCopyData, ByVal lngParam, Len(cdCopyData))

          Select Case cdCopyData.dwData
            Case 1
                Debug.Print "1"
            Case 2
                Debug.Print "2"
            Case 3
                    Call CopyMemory(byteBuffer(1), ByVal cdCopyData.lpData, cdCopyData.cbData)
                    strTemp = StrConv(byteBuffer, vbUnicode)
                    strTemp = Left$(strTemp, InStr(1, strTemp, Chr$(0)) - 1)
                    tmpStr() = Split(strTemp, "^")
                    If tmpStr(0) = "FVSNW" Then
                        If MsgBox("The file " & tmpStr(2) & " has been modified." & vbNewLine & "Would you like to import these changes into the vault?", vbYesNo, "File Change") = vbYes Then
                            Pass = GetPassword(PassPhrase)
                            m_AES.SetCipherKey Pass, KeySize, BlockSize
                            m_AES.FileEncrypt tmpStr(1), tmpStr(1) & ".tmp", BlockSize
                            CopyFromFileIntoADS tmpStr(1) & ".tmp", VaultPath, tmpStr(2), True
                            KillIt tmpStr(1) & ".tmp"
                            KillIt tmpStr(1)
                        Else
                            KillIt tmpStr(1)
                        End If
                    End If
          End Select
End Sub


