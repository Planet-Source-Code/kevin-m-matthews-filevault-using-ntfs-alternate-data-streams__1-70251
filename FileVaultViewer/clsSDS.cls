VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Not my Code. Credit goes entirely to John Underhill (Steppenwolfe)

Option Explicit

'***************************************************************************************
'*  SDS V1       Secure Document Shredder Class                                        *
'*                                                                                     *
'*  Created:     March 5, 2005                                                         *
'*  Updated:     April 13, 2006                                                        *
'*  Purpose:     Secure Document Destruction                                           *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.0                                                                   *
'*  Compile:     PCode                                                                 *
'*  Referenced:  Member Class SMT                                                      *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************

'/~ Properties ~/
'/~ p_SourceFile        - file to be destroyed
'/~ p_Passes            - number of overwrite cycles
'/~ p_Scattered         - scatter write blocks
'/~ p_Attributes        - reset file attributes

'/~ Exposed Routines ~/
'/~ File_Shred          - shred the file
'/~ File_Exists         - test for file existence


'/~ Notes ~/
'/~ Mysterious clicking sounds on the landline? Pizza van parked down the block
'/~ for over a week? Enron ex-pat?? Then this is exactly what you have been looking for!

'/~ I have tested this class against all the leading file recovery vendors, and none of
'/~ them could recover a viable file. [If anything at all], they can only recover a file with
'/~ the random data, nothing of the original document remained.
'/~ I considered using cellular data for the random blocks, but M$ crypto api is quite
'/~ fast, and in this type of application, very effective at producing random data.
'/~ Class uses write api to ensure commits. File cache is flushed at every pass, forcing
'/~ a write to the drive. I had to makes some concessions to speed on block size, but
'/~ given the number of passes, and techniques used, I doubt a file, [or even file fragment]
'/~ could be recovered from the drive. Maybe Rimnjants could chime in with some advice on this?
'/~ Anyhow, use it, don't abuse it, (and be careful! -improper use- of some of these api, can cause
'/~ serious harm to your file system!), and of course, there are no guarantees or warranties
'/~ for fitness of code or anything else.

'/~ You know the spiel.. for a comment or a job.. steppenwolfe_2000@yahoo.com
'/~ enjoy!


Private Const ALG_TYPE_ANY           As Long = 0
Private Const ALG_SID_MD5            As Long = 3
Private Const ALG_CLASS_HASH         As Long = 32768
Private Const HP_HASHVAL             As Long = 2
Private Const HP_HASHSIZE            As Long = 4
Private Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000
Private Const PROV_RSA_FULL          As Long = 1
Private Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal strFileName As String, _
                                                                        ByVal dwDesiredAccess As Long, _
                                                                        ByVal dwShareMode As Long, _
                                                                        ByVal lpSecurityAttributes As Long, _
                                                                        ByVal dwCreationDisposition As Long, _
                                                                        ByVal dwFlagsAndAttributes As Long, _
                                                                        ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
                                                        ByVal lDistanceToMove As Long, _
                                                        lpDistanceToMoveHigh As Long, _
                                                        ByVal dwMoveMethod As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
                                                   lpBuffer As Any, _
                                                   ByVal nNumberOfBytesToWrite As Long, _
                                                   lpNumberOfBytesWritten As Long, _
                                                   ByVal lpOverlapped As Any) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                                                                              ByVal pszContainer As String, _
                                                                                              ByVal pszProvider As String, _
                                                                                              ByVal dwProvType As Long, _
                                                                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, _
                                                            ByVal dwLen As Long, _
                                                            ByVal pbBuffer As String) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, _
                                                                 ByVal dwFlags As Long) As Long

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, _
                                                                                      ByVal dwFileAttributes As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, _
                                                                        ByVal lpNewFileName As String, _
                                                                        ByVal dwFlags As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

'/* shredder progress max
Public Event eSCompPMax(lMax As Long)
'/* shredder progress event
Public Event eSCompPTick(lCnt As Long)
'/* shredder complete
Public Event eSCompComplete()

Private m_sSourceFile       As String
Private m_lPasses           As Long
Private m_bScattered        As Boolean
Private m_bAttributes       As Boolean
Public KillComplete As Boolean


'/* source file path
Public Property Get p_SourceFile() As String
    p_SourceFile = m_sSourceFile
End Property

Public Property Let p_SourceFile(ByVal PropVal As String)
    m_sSourceFile = PropVal
End Property

'/* number of deletion passes
Public Property Get p_Passes() As Long
    p_Passes = m_lPasses
End Property

Public Property Let p_Passes(ByVal PropVal As Long)
    m_lPasses = PropVal
End Property

'/* use scattered block meshing
Public Property Get p_Scattered() As Boolean
    p_Scattered = m_bScattered
End Property

Public Property Let p_Scattered(ByVal PropVal As Boolean)
    m_bScattered = PropVal
End Property

'/* reset file attributes
Public Property Get p_Attributes() As Boolean
    p_Attributes = m_bAttributes
End Property

Public Property Let p_Attributes(ByVal PropVal As Boolean)
    m_bAttributes = PropVal
End Property


Public Sub File_Shred()
'/* core

Dim lLen        As Long
Dim lRemain     As Long
Dim lBCount     As Long
Dim lBlock      As Long
Dim sBlock      As String
Dim aBlock0()   As Byte
Dim aBlock1()   As Byte
Dim aBlock2()   As Byte
Dim aBlock3()   As Byte
Dim aBlock4()   As Byte
Dim aBlock5()   As Byte
Dim aBlock6()   As Byte
Dim aBlock7()   As Byte
Dim aBlock8()   As Byte
Dim aTemp()     As Byte
Dim lCount      As Long
Dim lPasses     As Long
Dim lByte       As Long
Dim lFile       As Long
Dim lOffset     As Long
Dim lNum        As Long
Dim lRet        As Long

On Error GoTo Handler

'/* block allocation table
'/* size random data chunks
'/* to file size perspective
'/* [8] divisible block units
'/* Note: done only for speed
'/* proper method [non vb] would use
'/* a fixed block size - 1024
'/* 128
'/* 1024
'/* 8192
'/* 65536
'/* 524288
'/* 4194304
    KillComplete = False
    lLen = FileLen(m_sSourceFile)
    If lLen < 8 Then GoTo Handler
    '/* reset file attributes to normal
    If p_Attributes Then Set_Attributes
    
    '/* choose the block size
    '/* based on file size
    '/* for speed
    Select Case True
    Case lLen > 4194304
        lBlock = 524288
    Case lLen > 524288
        lBlock = 65536
    Case lLen > 65536
        lBlock = 8192
    Case lLen > 8192
        lBlock = 1024
    Case lLen > 1024
        lBlock = 128
    Case lLen < 1024
        lBlock = lLen / 8
    End Select

    '/* block remainder allocation
    lBCount = Int(lLen / lBlock)
    lRemain = lLen - (lBCount * lBlock)

    '/* block building
    '/* build random data blocks
    '/* get the random sample string

    '~*** build random data blocks ***~

    '/* block 0
    sBlock = Set_Block(lBlock)
    'Debug.Print "length: " & lLen
    'Debug.Print "blocks: " & (lBlock * 12) + lRemain
    '/* dimension array to block size
    ReDim aBlock0(0 To lBlock - 1) As Byte
    '/* convert to byte array
    For lByte = 1 To lBlock
        aBlock0(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 1
    sBlock = Set_Block(lBlock)
    ReDim aBlock1(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock1(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 2
    sBlock = Set_Block(lBlock)
    ReDim aBlock2(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock2(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 3
    sBlock = Set_Block(lBlock)
    ReDim aBlock3(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock3(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 4
    sBlock = Set_Block(lBlock)
    ReDim aBlock4(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock4(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 5
    sBlock = Set_Block(lBlock)
    ReDim aBlock5(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock5(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 6
    sBlock = Set_Block(lBlock)
    ReDim aBlock6(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock6(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 7
    sBlock = Set_Block(lBlock)
    ReDim aBlock7(0 To lBlock - 1) As Byte
    For lByte = 1 To lBlock
        aBlock7(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
    Next lByte
    
    '/* block 8 - remainder
    If Not lRemain = 0 Then
        sBlock = Set_Block(lRemain)
        ReDim aBlock8(0 To lRemain - 1) As Byte
        For lByte = 1 To lRemain
            aBlock8(lByte - 1) = Asc(Mid$(sBlock, lByte, 1))
        Next lByte
    End If

    '~*** overwrite cycles ***~

    If m_lPasses = 0 Then m_lPasses = 1
    RaiseEvent eSCompPMax(m_lPasses)
    For lPasses = 1 To m_lPasses
        lOffset = 0
        '/* open for file handle
        lFile = CreateFile(m_sSourceFile, &H40000000, &H1, ByVal 0&, &H3, &H80, 0&)
        '/* a standard overwrite cycle
        For lNum = 1 To lBCount
            '/* move file pointer to next offset
            SetFilePointer lFile, lOffset, ByVal 0&, 0&
            '/* dimension temporary byte array
            ReDim aTemp(0 To lBlock - 1) As Byte
            '/* scatter overwrite block assignment
            If m_bScattered Then
                lCount = Int(Rnd * 8) + 1
            Else
                lCount = lCount + 1
            End If
            If lCount > 8 Then lCount = 1
            '/* sequential block writes
            Select Case lCount
            Case 1
                aTemp = aBlock0
            Case 2
                aTemp = aBlock1
            Case 3
                aTemp = aBlock2
            Case 4
                aTemp = aBlock3
            Case 5
                aTemp = aBlock4
            Case 6
                aTemp = aBlock5
            Case 7
                aTemp = aBlock6
            Case 8
                aTemp = aBlock7
            End Select
            '/* write to file
            WriteFile lFile, aTemp(0), lBlock, lRet, ByVal 0&
            lOffset = lOffset + lBlock
        Next lNum
        
        '/* test for valid remainder
        If Not lRemain = 0 Then
            SetFilePointer lFile, (lOffset), ByVal 0&, 0&
            WriteFile lFile, aBlock8(0), lRemain, lRet, ByVal 0&
        End If
        
        '/* write results and close
        FlushFileBuffers lFile
        CloseHandle lFile
        RaiseEvent eSCompPTick(lPasses)
        DoEvents
    Next lPasses
    '/* delete the file
    lRet = DeleteFile(m_sSourceFile)
    '/* delay on reboot
    If lRet = 0 Then Delayed_Deletion
    
    RaiseEvent eSCompComplete
    KillComplete = True
    
Handler:

End Sub

Private Function Set_Block(ByVal lBlock As Long) As String
'/* create chunks of random data using
'/* cryptogenrandom api. If trust is an issue
'/* consider using cellular data, ex. [ISAAC]
'/* but this is very fast, and [I believe],
'/* sufficiently random in this context

Dim lReturn     As Long
Dim sBlock      As String
Dim lProv       As Long

    '/* size the string
    sBlock = Space$(lBlock)
    '/* acquire crypto context
    lReturn = CryptAcquireContext(lProv, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
    If lReturn = 0 Then GoTo Handler
    '/* generate random block
    lReturn = CryptGenRandom(lProv, lBlock, sBlock)
    If lReturn = 0 Then GoTo Handler
    '/* release crypto engine
    lReturn = CryptReleaseContext(lProv, 0)
    If lReturn = 0 Then GoTo Handler
    '/* set block and exit
    Set_Block = sBlock
    'Debug.Print Len(sBlock)
Exit Function

Handler:

    '/* on error alternative
    Set_Block = Random_Key(lBlock)

End Function

Private Sub Delayed_Deletion()
'/* destroy on reboot

    p_SourceFile = Get_ShortPath
    MoveFileEx m_sSourceFile, &H0, &H4

End Sub

Private Function Get_ShortPath() As String
'/* get short file path

Dim lLen    As Long
Dim sBuffer As String

    sBuffer = String$(254, 0)
    lLen = GetShortPathName(m_sSourceFile, sBuffer, 255)
    Get_ShortPath = Left$(sBuffer, lLen)

End Function

Public Function File_Exists(ByVal sDir As String) As Boolean
'/* test file presence

Dim lR      As Long
Dim sP      As String

    sP = String$(260, 0)
    lR = GetShortPathName(sDir, sP, 259)
    File_Exists = lR > 0

End Function

Private Sub Set_Attributes()
'/* reset file attributes to normal

    SetFileAttributes m_sSourceFile, &H80

End Sub

Private Function Random_Key(ByVal lSize As Long) As String
'/* generate a random string (alternative)

Dim sTemp   As String
Dim sChar   As String
Dim iTemp   As Integer

    Do Until Len(sTemp) = lSize
        Randomize
        iTemp = Int(Rnd * 255)
        sChar = Chr$(iTemp)
        sTemp = sTemp + sChar
    Loop

    Random_Key = sTemp

End Function


