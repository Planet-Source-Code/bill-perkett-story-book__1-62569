VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FrmStart 
   Caption         =   "Main"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "FrmStart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBook 
      Caption         =   "Create A book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton ComGenie 
      Caption         =   "Genie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ComMerlin 
      Caption         =   "Merlin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ComPeedy 
      Caption         =   "Peedy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ComJames 
      Caption         =   "James"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Talking Spell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdProgram 
      Caption         =   "Create a Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   5040
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2880
      Picture         =   "FrmStart.frx":030A
      ScaleHeight     =   705
      ScaleWidth      =   825
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Do Magic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Install Msagent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Tell Jokes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "James"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Genie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Merlin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Peedy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdStory 
      Caption         =   "Story"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton CmdQuiz 
      Caption         =   "Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   360
      Top             =   2520
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim reqDisappear As Object
Dim reqSpeak As Object
Dim reqStart As Object
Dim reqConfused As Object
Dim FileLoc As String
Dim iMagic As Integer

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
'Function GetWindowsDir() As String
'
'Dim Temp As String
'Dim Ret As Long
'Const MAX_LENGTH = 145
'
'Temp = String$(MAX_LENGTH, 0)
'Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
'Temp = Left$(Temp, Ret)
'If Temp <> "" And Right$(Temp, 1) <> "\" Then
'    GetWindowsDir = Temp & "\"
'Else
'    GetWindowsDir = Temp
'End If
'
'End Function

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
     Dim n, i As Integer
     n = Int((700 * Rnd) + 1)
      n = n Mod 7 + 1
     MyAgent.StopAll
    Select Case n
     Case 1
      MyAgent.Play "Alert"
      MyAgent.Speak "Please don't touch me"
     Case 2
      MyAgent.Play "Blink"
      MyAgent.Speak "Hello"
     Case 3
      MyAgent.Play "Idle1_2"
      MyAgent.Speak "That tickles"
     Case 4
      MyAgent.Play "Idle1_1"
      MyAgent.Speak "How are you doing?"
     Case 5
        For i = 1 To 2
        MyAgent.Play "gestureup"
        MyAgent.Play "blink"
        MyAgent.Play "gesturedown"
         MyAgent.Play "blink"
         Next
         MyAgent.Speak "Dance!"
         
     Case 6
      'MyAgent.Play "Search"
      MyAgent.Think "I think that you should get back to work"
    Case 7
       MyAgent.Height = "240" 'Sets the Characters Height
       MyAgent.Width = "360" 'Sets the Characters Width
       'MyAgent.Show
       MyAgent.Speak "Look at me now!"
       Sleep 5000
       MyAgent.Height = "128"
       MyAgent.Width = "160"
    End Select
    MyAgent.Play "RestPose"
End Sub

Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
  'MyAgent.StopAll
  'MyAgent.Play "Blink"
End Sub

Private Sub Agent1_DragStart(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
  MyAgent.StopAll
  Dim n, L As Integer

'Randomize Timer
n = Int((4 * Rnd) + 1)
Select Case n
    Case 1
    MyAgent.Play "Think"
    Case 2
     MyAgent.Play "blink"
     Case 3
     MyAgent.Play "lookdown"
      Case 4
     MyAgent.Play "surprised"
 End Select
End Sub

Private Sub Agent1_IdleStart(ByVal CharacterID As String)
If bNoIdle = False Then Exit Sub
Dim n, L As Integer

'Randomize Timer
n = Int((1400 * Rnd) + 1)
n = n Mod 14 + 1
MyAgent.StopAll
'N = 3
Select Case n
    Case 1
        If Anim = "Robby" Then
            MyAgent.Play "Idle3_1"
        Else
            MyAgent.Play "LookDownBlink"
            MyAgent.Play "LookDownBlink"
            MyAgent.Play "LookDownBlink"
            MyAgent.Play "LookDownReturn"
            MyAgent.Stop
            MyAgent.MoveTo 300, 700
            MyAgent.Speak "Man It's really dark ..inside your monitor!"
            MyAgent.MoveTo 300, 50
            MyAgent.MoveTo 400, 350
            MyAgent.Speak "Nice to be back!"
            MyAgent.Speak "Lets try again"
        End If

    Case 2
            MyAgent.Play "LookDown"
            MyAgent.Play "LookDownBlink"
            MyAgent.Play "LookLeft"
            MyAgent.Play "LookLeftBlink"
            MyAgent.Play "LookUp"
            MyAgent.Play "LookUpBlink"
            MyAgent.Play "LookRight"
            MyAgent.Play "LookRightBlink"
            If bQuiz = False Then Exit Sub
            ' If s = False Then Exit Sub
            If Search >= 1 Then
                MyAgent.Speak "Are you still having a problem with this question?"
            Else
                MyAgent.Speak "Are you having a problem with this question?"
            End If
            Search = Search + 1
        'End If
        
    Case 3
       If bQuiz = False Then
         If bStart = False Then MyAgent.Speak "I think I know a story"
            MyAgent.Play "Think"
       Else
            MyAgent.Speak "I'll think I know the answer"
            MyAgent.Play "Think"
            MyAgent.Speak "You're in luck, I think the answer is "
            n = Int((4 * Rnd) + 1)
            L = Len(Answer(Q, n))
            MyAgent.Speak Right(Answer(Q, n), L - 1)
        End If
        
    Case 4
        MyAgent.Play "Idle1_1"
        MyAgent.Play "Idle1_2"
        MyAgent.Play "Idle1_3"
        MyAgent.Play "Idle1_4"
        If bQuiz = False Then Exit Sub
        If Search >= 1 Then
            MyAgent.Speak "It appears you are still struggling with this question!"
        Else
            MyAgent.Speak "It appears you are struggling with this question!"
        End If
        Search = Search + 1
        
    Case 5
        If bStart Then
          MyAgent.Play "Idle2_1"
          Exit Sub
        End If
        If bQuiz = False Then
         MyAgent.Speak "I'll search for a story"
        Else
            If Search >= 1 Then
                MyAgent.Speak "I'll look for the correct answer"
            Else
                MyAgent.Speak "I'll try to search for the correct answer"
            End If
            Search = Search + 1
        End If
        MyAgent.Play "Process"
        MyAgent.Speak "Sorry, no luck with the search"

    Case 6
        If bQuiz = False Then
          MyAgent.Play "Idle2_2"
        Else
        MyAgent.Speak "I'll try a different search pattern for the correct answer"
        MyAgent.Play "Search"
        MyAgent.Speak "I think it might be. But I wouldn't place any bets on it"
        MyAgent.Speak "because I'm not sure"
        n = Int((4 * Rnd) + 1)
        L = Len(Answer(Q, n))
        MyAgent.Speak Right(Answer(Q, n), L - 1)
        End If
    Case 7
       If bQuiz = False Then
          MyAgent.Play "Idle2_1"
        Else
            MyAgent.Speak "I can't hear you"
            MyAgent.Play "Alert"
        End If
    Case 8
        If bQuiz = False Then
          If bStart = False Then MyAgent.Speak "Are you ready for a story?"
        Else
          MyAgent.Speak "You want to make a guess?"
        End If
        MyAgent.Play "Wave"
        
    Case 9
        MyAgent.Play "Suggest"
         If bQuiz = False Then
           If bStart = False Then
                MyAgent.Speak "Let me read you a story."
           Else
            L = Int((2 * Rnd) + 1)
             If L = 1 Then MyAgent.Speak "Try the quiz."
             If L = 2 Then MyAgent.Speak "Let me read you a story."
           End If
        Else
            MyAgent.Speak "The answer might be"
            n = Int((4 * Rnd) + 1)
            L = Len(Answer(Q, n))
            MyAgent.Speak Right(Answer(Q, n), L - 1)
        End If
     Case 10
        MyAgent.Play "Idle3_1"
        MyAgent.Speak "I almost went to sleep"
        
     Case 11
        MyAgent.Play "GetAttention"
        MyAgent.Speak "Is any body there?"
        
     Case 12
         MyAgent.Play "Confused"
        If bQuiz = False Then Exit Sub
        MyAgent.Speak "I don't know the answer?"
     Case 13
       MyAgent.MoveTo 2000, 300 'Moves him to co ordinates 2000,300 (off the screen!)
       MyAgent.MoveTo 300, 300 'Moves to co ordinates 300,300 (lower middle of screen)
      ' MyAgent.Play "confused" 'Looks Confused
       MyAgent.Speak "Nothing like a little flying to clear the head!" 'Speaks
       MyAgent.Play "pleased" 'Looks pleased
    Case 14
       MyAgent.Height = "240" 'Sets the Characters Height
       MyAgent.Width = "360" 'Sets the Characters Width
       'MyAgent.Show
       MyAgent.Speak "Look at me now!"
       Sleep 5000
       MyAgent.Height = "128"
       MyAgent.Width = "160"
    Case Else
        MyAgent.Play "Idle3_1"
        
 End Select
 
End Sub

Private Sub Agent1_RequestComplete(ByVal Request As Object)
  If Request Is reqDisappear Then
    ' picPicture.Visible = False
     If iMagic = 2 Or iMagic = 3 Then
          Set reqSpeak = MyAgent.Speak("Yes, it worked!")
          picPicture.Visible = False
     End If
     If iMagic = 4 Then
          MyAgent.Play ("Confused")
          Set reqSpeak = MyAgent.Speak("What happened?")
         ' MyAgent.Play ("Confused")
          'picPicture.Visible = False
     End If
     If iMagic = 1 Then
       MyAgent.Height = "240" 'Sets the Characters Height
       MyAgent.Width = "360" 'Sets the Characters Width
       'MyAgent.Show
       MyAgent.Speak "Look at me now!"
       Sleep 5000
       MyAgent.Height = "128"
       MyAgent.Width = "160"
       bNoIdle = False
     End If
ElseIf Request Is reqSpeak Then
    Select Case iMagic
     Case 2
        picPicture.Visible = True
        MyAgent.Play "surprised"
        'MyAgent.Play "RestPose"
        MyAgent.Speak "Oh, no. I will have to work" _
        & " on my magic."
        MyAgent.Play "process"
        MyAgent.Speak "This calls for more work, I will" _
        & " have to consult the Grand Wizard!"
        MyAgent.Hide
        bHide = True
     Case 3
      MyAgent.Play "suggest"
      MyAgent.Speak ("I think i got it!")
     Case 4
       picPicture.Visible = False
   End Select
   bNoIdle = False
End If
If Request Is reqStart Then
     picPicture.Visible = True
     'Command4.Visible = True
 End If
End Sub

Private Sub CmdBook_Click()
 bStart = False
  MyAgent.StopAll
  MyAgent.Play "Hide"
  Agent1.Characters.Unload Anim
  FrmNewBook.Show
End Sub

Private Sub CmdEdit_Click()
    MyAgent.StopAll
    Form2.Show
End Sub

Private Sub CmdProgram_Click()
  bStart = False
  MyAgent.StopAll
  MyAgent.Play "Hide"
  'MyAgent.c
  'Set MyAgent = Nothing
  Agent1.Characters.Unload Anim
  'MyAgent.StopAll
  frmClient.Show
End Sub

Private Sub CmdQuiz_Click()
  bStart = False
  MyAgent.StopAll
  frmMVQuiz.Show
End Sub

Private Sub CmdStory_Click()
  bStart = False
  MyAgent.StopAll
  FrmStory.Show
End Sub

Private Sub ComGenie_Click()
 Dim cTrain As String
 Dim cShellcmd As String
 cTrain = Mid(App.Path, 1, 2)
 cTrain = cTrain & "\kids\Genie.exe"
 cShellcmd = """" & cTrain & """"
 Shell cShellcmd, vbNormalFocus
End Sub

Private Sub ComJames_Click()
 Dim cTrain As String
 Dim cShellcmd As String
 cTrain = Mid(App.Path, 1, 2)
 cTrain = cTrain & "\kids\James.exe"
 cShellcmd = """" & cTrain & """"
 Shell cShellcmd, vbNormalFocus
End Sub

Private Sub Command1_Click()
Form1.Show
MyAgent.StopAll
MyAgent.Hide
Set MyAgent = Nothing
Agent1.Characters.Unload Anim
bAgentActive = False
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Command3_Click()
'
'
'
  MsgBox "Go to http://www.microsoft.com/msagent/downloads/user.asp to download msagent", vbOKOnly, "No Agent"
  Exit Sub
' Dim cTrain As String
' Dim cMyOut As String
' Dim j, k As Integer
' j = InStr(1, cQuizfile, "Quiz") '- 1
' k = Len(cQuizfile) + 1
' 'cTrain = App.Path & "\MvQuiz.exe"
' Dim cShellcmd As String
'' cShellcmd = """" & cTrain & """"
'' Shell cShellcmd, vbNormalFocus
'' Run an external ProgramMerlin.exe
' cTrain = App.Path & "\MSagent.exe"
' cShellcmd = """" & cTrain & """"
'  Shell cShellcmd, vbNormalFocus
'  Sleep 8000
''Run an external Program
'cTrain = App.Path & "\Peedy.exe"
' cShellcmd = """" & cTrain & """"
' Shell cShellcmd, vbNormalFocus
' 'Run an external Program
' cTrain = App.Path & "\tv_enua.exe"
'  cShellcmd = """" & cTrain & """"
' Shell cShellcmd, vbNormalFocus
'  Sleep 8000
' 'Run an external ProgramMerlin.exe
' cTrain = App.Path & "\Genie.exe"
'  cShellcmd = """" & cTrain & """"
' Shell cShellcmd, vbNormalFocus
'  Sleep 8000
' 'Run an external ProgramMerlin.exe
' cTrain = App.Path & "\Merlin.exe"
' cShellcmd = """" & cTrain & """"
' Shell cShellcmd, vbNormalFocus
'  Sleep 8000
'  'Run an external ProgramMerlin.exe
' cTrain = App.Path & "\James.exe"
' cShellcmd = """" & cTrain & """"
' Shell cShellcmd, vbNormalFocus
'' 'Run an external ProgramMerlin.exe
'' cTrain = App.Path & "\MSagent.exe"
'' cShellcmd = """" & cTrain & """" & " /" & cMyOut
'' Shell cShellcmd, vbNormalFocus
End Sub





Private Sub Command4_Click()
Dim n As Integer
Dim srtSpeak As String
   MyAgent.Show 'This shows Merlin
    MyAgent.MoveTo Me.Left / 15 + picPicture.Left / _
    15 + 120, Me.Top / 15 + picPicture.Top _
    / 15 'Makes Merlin move next to the picture.
    n = Int((500 * Rnd) + 1)
    n = n Mod 5 + 1
    iMagic = Int((500 * Rnd) + 1)
    iMagic = iMagic Mod 3 + 1
    'iMagic = 1
    bNoIdle = True
    Select Case n
    Case 1
      srtSpeak = "Hocus Pocus"
    Case 2
      srtSpeak = "Abracadabra"
    Case 3
      srtSpeak = "Bipity Bopity Boo"
    Case 4
      srtSpeak = "Ala peanut butter sandwich"
    Case 5
      srtSpeak = "Open sesame"
    End Select
    MyAgent.Speak srtSpeak  'This makes Merlin _
   MyAgent.Play "DoMagic1" 'This makes Merlin lift his wand
   Set reqDisappear = MyAgent.Play("DoMagic2") 'Thi
   Command4.Visible = False
End Sub



Private Sub Command5_Click()
MyAgent.Speak "Open sesame"
Dim n, i As Integer
Text1.Text = ""
For i = 1 To 10
 n = Int((1400 * Rnd) + 1)
 n = n Mod 14 + 1
 Text1.Text = Text1.Text & n & " "
Next

End Sub



Private Sub Command7_Click()
   bStart = False
  'Set MyAgent = Nothing
  'Agent1.Characters.Unload Anim
  'MyAgent.StopAll
  FrmWord.Show
End Sub

Private Sub ComMerlin_Click()
 Dim cTrain As String
 Dim cShellcmd As String
 cTrain = Mid(App.Path, 1, 2)
 cTrain = cTrain & "\kids\Merlin.exe"
 cShellcmd = """" & cTrain & """"
 Shell cShellcmd, vbNormalFocus
End Sub

Private Sub ComPeedy_Click()
 Dim cTrain As String
 Dim cShellcmd As String
 cTrain = Mid(App.Path, 1, 2)
 cTrain = cTrain & "\kids\Peedy.exe"
 cShellcmd = """" & cTrain & """"
 Shell cShellcmd, vbNormalFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim DirName As String
bAgentActive = False
bFirstTime = False
bStory = False
bQuiz = False
bNoIdle = False
bStart = True
bHide = False
If App.PrevInstance = True Then
    MsgBox "This application is already running!", vbInformation + vbOKOnly, "Motor Vehicle Quiz is Running"
    End
End If
If Right(App.Path, 1) = "\" Then
    Ap = App.Path
Else
    Ap = App.Path & "\"
End If

CentreMe Me
Me.Show
retval = waveOutGetNumDevs()
Randomize
If retval = 0 Then
    MsgBox "Your system cannot play Sound Files." & vbCrLf & vbCrLf & "So you won't hear any speech!", 48, "SoundCard Check"
End If
DirName = GetWindowsDir()
FileLoc = DirName & "Msagent\"
Success = IfFileExists(FileLoc & "Agentctl.dll")
If Success = False Then
    MsgBox "Msagent is not installed on this computer!" & vbCrLf & vbCrLf & "You can download them from:-" & vbCrLf & vbCrLf & "http://www.microsoft.com/msagent"
    Command3.Visible = True
    CmdQuiz.Visible = False
    CmdStory.Visible = False
    'Exit Sub
End If
FileLoc = DirName & "Msagent\Chars\"
Success = IfFileExists(FileLoc & "Peedy.acs")
If Success = False Then
    MsgBox "Merlin is not installed on this computer!" & vbCrLf & vbCrLf & "You my need MSagent as well." & vbCrLf & vbCrLf & "You can download them from:-" & vbCrLf & vbCrLf & "http://www.microsoft.com/msagent"
    Command3.Visible = True
    CmdQuiz.Visible = False
    CmdStory.Visible = False
    Exit Sub
End If

Q = 0
Anim = "Peedy"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English

MyAgent.MoveTo 400, 350
MyAgent.Show
bAgentActive = True
'myagent.
cQuizfile = App.Path & "\Quiz1.txt"
Option1(2).Value = True

'Label3.Caption = cQuizfile
Exit Sub
ErrHandler:
MsgBox Err.Description
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    Success = IfFileExists(FileLoc & "Genie.acs")
      cStoryAgent = "Genie"
    If Success = False Then
'        MyAgent.Play "Search"
'        MyAgent.Speak "Sorry, Genie is not installed on this computer!"
'        MyAgent.Speak "You can download him from"
'        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
        Command3.Visible = True
        ComGenie.Visible = True
        Exit Sub
       End If
End If
If Index = 1 Then
    Success = IfFileExists(FileLoc & "Merlin.acs")
     cStoryAgent = "Merlin"
    If Success = False Then
'        MyAgent.Play "Search"
'        MyAgent.Speak "Sorry, Merlin is not installed on this computer!"
'        MyAgent.Speak "You can download him from"
'        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
         Command3.Visible = True
         ComMerlin.Visible = True
        Exit Sub
       End If
End If
If Index = 2 Then
    Success = IfFileExists(FileLoc & "Peedy.acs")
     cStoryAgent = "Peedy"
    If Success = False Then
'        MyAgent.Play "Search"
'        MyAgent.Speak "Sorry, Peedy is not installed on this computer!"
'        MyAgent.Speak "You can download him from"
'        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
         Command3.Visible = True
         ComPeedy.Visible = True
        Exit Sub
       End If
End If
If Index = 3 Then
    Success = IfFileExists(FileLoc & "James.acs")
    cStoryAgent = "James"
    If Success = False Then
'        MyAgent.Play "Search"
'        MyAgent.Speak "Sorry, Peedy is not installed on this computer!"
'        MyAgent.Speak "You can download him from"
'        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
         Command3.Visible = True
         ComJames.Visible = True
        Exit Sub
       End If
End If
If bAgentActive Then
    MyAgent.StopAll
    MyAgent.Hide
    Set MyAgent = Nothing
    Agent1.Characters.Unload Anim
End If
'Agent1.Characters.Unload Anim
If Index = 0 Then Anim = "Genie"
If Index = 1 Then Anim = "Merlin"
If Index = 2 Then Anim = "Peedy"
If Index = 3 Then Anim = "James"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H409 'Language ID = English
MyAgent.MoveTo 400, 350

MyAgent.Show
'If bFirstTime Then Exit Sub
'bFirstTime = True
MyAgent.Play "Wave"
If bFirstTime Then Exit Sub
bFirstTime = True
MyAgent.Speak "Do you want to try the quiz"
MyAgent.Speak "or"
MyAgent.Speak "do you want to hear a story?"
Set reqStart = MyAgent.Play("Pleased") 'Thi
End Sub

