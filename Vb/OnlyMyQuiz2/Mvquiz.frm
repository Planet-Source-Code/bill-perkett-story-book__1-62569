VERSION 5.00
Begin VB.Form frmMVQuiz 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quiz with Animation"
   ClientHeight    =   6510
   ClientLeft      =   1080
   ClientTop       =   1785
   ClientWidth     =   9480
   Icon            =   "Mvquiz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optAnswer4 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer3 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   555
      Left            =   7845
      TabIndex        =   0
      Top             =   5475
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image imgSpanner 
      Height          =   480
      Left            =   2400
      Picture         =   "Mvquiz.frx":030A
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblQuestions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF0000&
      Caption         =   "Question:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF0000&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF0000&
      Caption         =   "4:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "3:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF0000&
      Caption         =   "2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblAnswer4 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReStart 
         Caption         =   "Restart Quiz"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Character &Options"
      Begin VB.Menu mnuChange 
         Caption         =   "&Change Character"
         Visible         =   0   'False
         Begin VB.Menu mnuGenie 
            Caption         =   "&Genie the Genie"
         End
         Begin VB.Menu mnuMerlin 
            Caption         =   "&Merlin the Magician"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPeedy 
            Caption         =   "&Peedy the Parrot"
         End
         Begin VB.Menu mnuRobot 
            Caption         =   "&Robby the Robot"
         End
      End
      Begin VB.Menu mnuBalloons 
         Caption         =   "&Display Word Balloons"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuRepeat 
      Caption         =   "Repeat &Question"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMVQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iMyClick As Integer
Dim FileNum As Integer
'Dim Question(40) As String
'Dim Answer(40, 5) As String
'Dim MyAnswer As String
'Dim YesNo(5) As String
'Dim Ans(5) As Integer
'Dim Q As Integer
Dim Guess As Integer
Dim Score As Integer
Dim Questions As Integer
Dim Finalscore As String
Dim Level As String
Dim Response As Integer
'Dim Search As Integer
Dim FileLoc As String
Dim iMax As Integer


Private Sub cmdContinue_Click()
Dim n As Integer
 n = Int((100 * Rnd) + 1)
 n = n Mod 10 + 1
MyAgent.StopAll
If UCase(Mid(YesNo(Guess), 1, 1)) = "Y" Then
     Score = Score + 1
     Sleep 100
     Select Case n
       Case 1
          MyAgent.Play "Congratulate"
        Case 2
          MyAgent.Play "Announce"
          MyAgent.Speak "Well Done"
        Case 3
          MyAgent.Play "DoMAgic1"
          MyAgent.Speak "Right"
        Case 4
          MyAgent.Speak "Good job"
        Case 5
           MyAgent.Play "Pleased"
           MyAgent.Speak "That is right"
        Case 6
          MyAgent.Speak "Super"
        Case 7
          MyAgent.Play "Write"
          MyAgent.Speak "Great"
        Case 8
          MyAgent.Play "Show"
          MyAgent.Speak "You did it"
        Case 9
          'MyAgent.Play "Show"
          MyAgent.Think "You are right"
         Case 10
          'MyAgent.Play "Show"
          MyAgent.Play "Wave"
     End Select
     Sleep 5000
Else
      Sleep 100
      Select Case n
       Case 1
          MyAgent.Play "Confused"
        Case 2
          MyAgent.Play "Uncertain"
        Case 3
          MyAgent.Speak "That is not it"
        Case 4
          MyAgent.Speak "Sorry"
        Case 5
          MyAgent.Speak "Wrong one"
        Case 6
          MyAgent.Play "Blink"
          MyAgent.Play "Blink"
        Case 7
          MyAgent.Play "Read"
        Case 8
          MyAgent.Play "Explain"
       Case 9
          'MyAgent.Play "Show"
          MyAgent.Think "Oh no"
       Case 10
          'MyAgent.Play "Show"
          MyAgent.Play "Decline"
       End Select
 MyAgent.Speak "The correct answer is": MyAgent.Speak MyAnswer: Sleep 5000
End If
lblScore.Caption = Score

If Q >= iMax Then
  result
 Exit Sub
End If
Runtime

End Sub





Private Sub Command1_Click()
  Q = 13
  Questions = 13
  Score = 12
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKey1 Then optAnswer1.Value = True
If KeyCode = vbKeyNumpad1 Then optAnswer1.Value = True
If KeyCode = vbKey2 Then optAnswer2.Value = True
If KeyCode = vbKeyNumpad2 Then optAnswer2.Value = True
If KeyCode = vbKey3 Then optAnswer3.Value = True
If KeyCode = vbKeyNumpad3 Then optAnswer3.Value = True
If KeyCode = vbKey4 Then optAnswer4.Value = True
If KeyCode = vbKeyNumpad4 Then optAnswer4.Value = True

End Sub

Private Sub Form_Load()
bQuiz = True
On Error GoTo ErrHandler
Search = 1
Dim DirName As String
iMyClick = 2
If bHide Then
            MyAgent.Show
            bHide = False
 End If
Randomize
CentreMe Me
Me.Show
'retval = waveOutGetNumDevs()
iMax = 15


'mnuInstructions_Click
optAnswer1.Visible = True
optAnswer2.Visible = True
optAnswer3.Visible = True
optAnswer4.Visible = True

optAnswer1.Value = False
cmdContinue.Visible = False
Init
Quiz
Runtime

Exit Sub
ErrHandler:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

MyAgent.StopAll
If Q < 2 Then
    MyAgent.Speak "You haven't started the quiz yet. Why are you leaving?"
ElseIf Q >= 2 And Q < iMax Then
    MyAgent.Speak "You haven't finished the quiz yet. You have only attempted " & Q & " questions. Why are you leaving?"
End If
Sleep 6000
MyAgent.StopAll
FinishSave
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
End

End Sub

Private Sub mnuAbout_Click()

MyAgent.StopAll
frmAbout.Show 1

End Sub

Private Sub mnuExit_Click()

MyAgent.StopAll
If Q < 2 Then
    MyAgent.Speak "You haven't started the quiz yet. Why are you leaving?"
ElseIf Q >= 2 And Q < 40 Then
    MyAgent.Speak "You haven't finished the quiz yet. You have only attempted " & Q & " questions. Why are you leaving?"
End If
Sleep 6000
MyAgent.StopAll
FinishSave
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
End

End Sub

Public Sub Runtime()
Dim cMsg As String
Dim L As Integer
Randomize Timer

Q = Q + 1

Questions = Questions + 1
lblScore.Caption = Score
lblQuestions.Caption = Questions

Ans(1) = Int(Rnd * 4 + 1)
Do
    Ans(2) = Int(Rnd * 4 + 1)
Loop Until Ans(2) <> Ans(1)

Do
    Ans(3) = Int(Rnd * 4 + 1)
Loop Until Ans(3) <> Ans(1) And Ans(3) <> Ans(2)
Ans(4) = 10 - Ans(1) - Ans(2) - Ans(3)

LblQuestion.Caption = Question(Q)

L = Len(Answer(Q, Ans(1)))
lblAnswer1.Caption = Right(Answer(Q, Ans(1)), L - 1)
YesNo(1) = UCase(Answer(Q, Ans(1)))
If Mid(YesNo(1), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(1)), L - 1)
L = Len(Answer(Q, Ans(2)))
lblAnswer2.Caption = Right(Answer(Q, Ans(2)), L - 1)
YesNo(2) = UCase(Answer(Q, Ans(2)))
If Mid(YesNo(2), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(2)), L - 1)
L = Len(Answer(Q, Ans(3)))
lblAnswer3.Caption = Right(Answer(Q, Ans(3)), L - 1)
YesNo(3) = UCase(Answer(Q, Ans(3)))
If Mid(YesNo(3), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(3)), L - 1)
L = Len(Answer(Q, Ans(4)))
lblAnswer4.Caption = Right(Answer(Q, Ans(4)), L - 1)
'YesNo(1) = UCase(Answer(Q, Ans(1)))
'YesNo(2) = UCase(Answer(Q, Ans(2)))
'YesNo(3) = UCase(Answer(Q, Ans(3)))
YesNo(4) = UCase(Answer(Q, Ans(4)))
'If Mid(YesNo(1), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(1)), L - 1)
'If Mid(YesNo(2), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(2)), L - 1)
'If Mid(YesNo(3), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(3)), L - 1)
If Mid(YesNo(4), 1, 1) = "Y" Then MyAnswer = Right(Answer(Q, Ans(4)), L - 1)
Text1.Text = MyAnswer
Guess = 0
optAnswer1.Value = False
optAnswer2.Value = False
optAnswer3.Value = False
optAnswer4.Value = False
cmdContinue.Visible = False
MyAgent.Speak "Question " & Q
MyAgent.Speak LblQuestion.Caption
MyAgent.Speak lblAnswer1.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer2.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer3.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer4.Caption

End Sub

Private Sub mnuGenie_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Genie.acs")
If Success = False Then
    MyAgent.Play "Search"
    MyAgent.Speak "Sorry, Genie is not installed on this computer!"
    MyAgent.Speak "You can download him from:-"
    MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
Anim = "Genie"
FrmStart.Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = FrmStart.Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = True
mnuPeedy.Checked = False
mnuRobot.Checked = False
mnuMerlin.Checked = False

End Sub

Private Sub mnuInstructions_Click()

MyAgent.Speak "You will be shown 15 questions."
MyAgent.Speak "Read each question carefully."
MyAgent.Speak "Keys:-"
MyAgent.Speak "You can use the Mouse to enter your choice."
MyAgent.Speak "You can also use the Numeric Keys to enter your choice, then press the Enter key."
MyAgent.Speak "Good Luck!!!"

optAnswer1.Value = False
optAnswer2.Value = False
optAnswer3.Value = False
optAnswer4.Value = False

End Sub

Private Sub mnuBalloons_Click()

mnuBalloons.Checked = Not mnuBalloons.Checked

Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select

End Sub

Private Sub mnuMerlin_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Merlin.acs")
If Success = False Then
    MyAgent.Play "Search"
    MyAgent.Speak "Sorry, Merlin is not installed on this computer!"
    MyAgent.Speak "You can download him from"
    MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
Anim = "Merlin"
FrmStart.Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = FrmStart.Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = False
mnuRobot.Checked = False
mnuMerlin.Checked = True

End Sub

Private Sub mnuPeedy_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Peedy.acs")
If Success = False Then
    MyAgent.Play "Search"
    MyAgent.Speak "Sorry, Peedy is not installed on this computer!"
    MyAgent.Speak "You can download him from"
    MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
Anim = "Peedy"
FrmStart.Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = FrmStart.Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = True
mnuRobot.Checked = False
mnuMerlin.Checked = False

End Sub

Private Sub mnuRepeat_Click()

MyAgent.StopAll
MyAgent.Speak "Question " & Questions
MyAgent.Speak LblQuestion.Caption
MyAgent.Speak lblAnswer1.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer2.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer3.Caption
MyAgent.Speak "or"
MyAgent.Speak lblAnswer4.Caption

End Sub

Private Sub mnuReStart_Click()

MyAgent.StopAll
If Q >= 2 And Q < 40 Then
    MyAgent.Speak "You haven't finished the quiz yet. You have only attempted " & Q & " questions. Are you sure you want to restart the quiz ?"
End If
Sleep 6000
Response = MsgBox("Do you wish to Restart the Quiz.", 36, "Motor Vehicle Quiz")
If Response = vbNo Then Exit Sub

optAnswer1.Visible = True
optAnswer2.Visible = True
optAnswer3.Visible = True
optAnswer4.Visible = True
Init
Runtime

End Sub

Private Sub mnuRobot_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Robby.acs")
If Success = False Then
    MyAgent.Play "Search"
    MyAgent.Speak "Sorry, Robby is not installed on this computer!"
    MyAgent.Speak "You can download him from"
    MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
Anim = "Robby"
FrmStart.Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = FrmStart.Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = False
mnuRobot.Checked = True
mnuMerlin.Checked = False

End Sub

Private Sub optAnswer1_Click()

Guess = 1
cmdContinue.Visible = True

End Sub

Private Sub optAnswer1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer1.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer1.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer2_Click()

Guess = 2
cmdContinue.Visible = True

End Sub

Private Sub optAnswer2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer2.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer2.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer3_Click()

Guess = 3
cmdContinue.Visible = True

End Sub

Private Sub optAnswer3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer3.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer3.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer4_Click()

Guess = 4
cmdContinue.Visible = True

End Sub

Public Sub Init()

Q = 0
Score = 0
Questions = 0
Search = 0
Label7.Caption = "1:"
Label8.Caption = "2:"
Label9.Caption = "3:"
Label10.Caption = "4:"
End Sub

Public Sub result()

optAnswer1.Visible = False
optAnswer2.Visible = False
optAnswer3.Visible = False
optAnswer4.Visible = False
lblAnswer1.Caption = ""
lblAnswer2.Caption = ""
lblAnswer3.Caption = ""
lblAnswer4.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
cmdContinue.Visible = False

If Score < 9 Then Level = "Try the quiz again."
If Score >= 9 And Score <= 11 Then Level = "You are average."
If Score >= 12 And Score <= 14 Then Level = "You are good."
If Score >= 15 And Score <= 15 Then Level = "You are great."

Finalscore = "You scored " & Score & " out of " & Questions & "."
lblAnswer1.Caption = "Finalscore: "
lblAnswer3.Caption = Finalscore
If Q >= iMax Then
    Select Case Score
        Case Is >= 15
            MyAgent.Play "Congratulate"
        Case Is >= 9 And Score <= 14
            MyAgent.Play "acknowledge"
        Case Is < 9
            MyAgent.Play "Confused"
    End Select
    MyAgent.Speak Finalscore
    MyAgent.Speak Level
    MyAgent.Speak "Do you wish to Re-Run the Quiz?"
    MyAgent.Play "Writing"
    Response = MsgBox("Do you wish to Re-Run the Quiz.", 36, "Quiz")
    If Response = vbYes Then
        FinishSave
        MyAgent.StopAll
        mnuReStart_Click
    Else
        FinishSave
        Set MyAgent = Nothing
        FrmStart.Agent1.Characters.Unload Anim
        End
    End If
End If

End Sub

Public Sub StartSave()

'Dim File As String
'On Error GoTo ErrHandler
'
'File = Ap & "Results.dat"
'FileNum = FreeFile()
'Open File For Append As #FileNum
'Write #FileNum, "Start - " & Me.Caption, Format$(Date$, "dd/mm/yyyy"), Time$
'Close #FileNum
'
'Exit Sub
'ErrHandler:
'If Err = 61 Then
'    MsgBox "This Disk is Full!" & vbCrLf & vbCrLf & "Put this program on a blank disk!", vbCritical + vbOKOnly, "Program Error"
'    End
'Else
'    MsgBox Err.Description
'End If

End Sub

Public Sub FinishSave()

'Dim File As String
'On Error GoTo ErrHandler
'
'File = Ap & "Results.dat"
'If Questions < 40 Then Questions = Questions - 1
'FileNum = FreeFile()
'Open File For Append As #FileNum
'Write #FileNum, "Finish - " & Me.Caption, Format$(Date$, "dd/mm/yyyy"), Time$, Score, Questions
'Close #FileNum
'
'Exit Sub
'ErrHandler:
'If Err = 61 Then
'    MsgBox "This Disk is Full!" & vbCrLf & vbCrLf & "Put this program on a blank disk!", vbCritical + vbOKOnly, "Program Error"
'    End
'Else
'    MsgBox Err.Description
'End If

End Sub

Public Sub Quiz()
 Dim ival As Integer
 Dim j, L As Integer
 Dim iCnt As Integer
 Dim cMsg As String
 Call ReadQuiz(cQuizfile)
 If iMax > iQuizcnt Then iMax = iQuizcnt
 iCnt = 1
PickQuestion:
 'iCnt = iCnt + 1
 ival = Int((Rnd * iQuizcnt) + 1)
 If ival > iQuizcnt Then ival = 1
 j = Len(cQuiz(ival, 0))
 'lblQuestion = Mid(cQuiz(iVal, 0), 2, j - 1)
 If cQuiz(ival, 5) = "Y" Then GoTo PickQuestion
 cQuiz(ival, 5) = "Y"
 'For j = 1 To iQuizcnt
  L = Len(cQuiz(ival, 0))
  If L < 1 Then
     Answer(iCnt, 1) = ""
  End If
  Question(iCnt) = Right(cQuiz(ival, 0), L - 1)
  Answer(iCnt, 1) = cQuiz(ival, 1)
  Answer(iCnt, 2) = cQuiz(ival, 2)
  Answer(iCnt, 3) = cQuiz(ival, 3)
  Answer(iCnt, 4) = cQuiz(ival, 4)
' Next
 If iCnt < iMax Then
  iCnt = iCnt + 1
  GoTo PickQuestion
End If

End Sub

Function GetWindowsDir() As String

Dim Temp As String
Dim Ret As Long
Const MAX_LENGTH = 145

Temp = String$(MAX_LENGTH, 0)
Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
Temp = Left$(Temp, Ret)
If Temp <> "" And Right$(Temp, 1) <> "\" Then
    GetWindowsDir = Temp & "\"
Else
    GetWindowsDir = Temp
End If

End Function

Sub CentreMe(P1 As Form)

If TypeOf P1 Is Form Then
    P1.Left = (Screen.Width - P1.Width) / 2
    P1.Top = (Screen.Height - P1.Height) / 2
End If

End Sub

Function IfFileExists(Fname As String) As Boolean

On Local Error Resume Next
Dim F As Integer

F = FreeFile()
Open Fname For Input As #F
If Err Then
    IfFileExists = False
Else
    IfFileExists = True
End If
Close #F

End Function

Private Sub optAnswer4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer4.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

optAnswer4.MouseIcon = imgSpanner

End Sub
