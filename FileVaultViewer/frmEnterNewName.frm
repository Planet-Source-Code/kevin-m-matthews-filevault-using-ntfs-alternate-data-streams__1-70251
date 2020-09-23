VERSION 5.00
Begin VB.Form frmEnterNewName 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter new filename"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   0
      Width           =   1575
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   240
         Picture         =   "frmEnterNewName.frx":0000
         ScaleHeight     =   1230
         ScaleWidth      =   1260
         TabIndex        =   5
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtNewStreamName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00BA7457&
      Caption         =   $"frmEnterNewName.frx":50FA
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmEnterNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ' Set return variable to false
    ' and exit the form
    frmEnterNewNameRet = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim tmpSTR() As String
    ' Does the text entered in the textbox contain illegal charachters?
    If ContainsIllegalChars(txtNewStreamName.Text) Then
        ' If so, alert user
        MsgBox "Filenames cannot contain any of" & vbNewLine & "the following illegal chars:" & vbNewLine & "\/:*?""<>|", vbExclamation, "Filename Error"
        frmEnterNewNameRet = False
    ' Otherwise, perform some more validation on the filename
    Else
        frmEnterNewNameFilename = txtNewStreamName.Text
        ' we need to find out if the proper extension was entered by user
        tmpSTR() = Split(frmEnterNewNameFilename, ".")
        ' if it wasn't, then tack a . and the extension on the end of the filename
        If "." & tmpSTR(UBound(tmpSTR)) <> frmEnterNewNameExtension Then
            frmEnterNewNameFilename = frmEnterNewNameFilename & frmEnterNewNameExtension
        End If
        ' set the return variable to true (because the string passes validation)
        frmEnterNewNameRet = True
        Unload Me
    End If
End Sub
