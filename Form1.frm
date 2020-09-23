VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim clsStartTerm1 As New clsStartTerminateProgram
    Dim clsStartTerm2 As New clsStartTerminateProgram
    Dim clsStartTerm3 As New clsStartTerminateProgram
    'Dim starttime As Single
    
    clsStartTerm1.StartProgram "c:\windows\notepad.exe"
    clsStartTerm1.WaitForProgramToEnd
    MsgBox "Notepad closed!", vbInformation, "Done waiting for Notepad 1 to close"
    
    clsStartTerm2.StartProgram "c:\windows\notepad.exe"
    clsStartTerm2.WaitForProgramToEnd
    MsgBox "Notepad closed!", vbInformation, "Done waiting for Notepad 2 to close"
    
    clsStartTerm3.StartProgram "c:\windows\notepad.exe"
    clsStartTerm3.WaitForProgramToEnd
    MsgBox "Notepad closed!", vbInformation, "Done waiting for Notepad 3 to close"

''''''    starttime = Timer + 10
''''''    While Timer < starttime
''''''        DoEvents: DoEvents: DoEvents
''''''    Wend
    
''''''    clsStartTerm1.KillProgram
''''''    clsStartTerm2.KillProgram
''''''    clsStartTerm3.KillProgram
    
    Set clsStartTerm1 = Nothing
    Set clsStartTerm2 = Nothing
    Set clsStartTerm3 = Nothing
    
    
End Sub
