VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FrmStory 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSpiral 
      Caption         =   "Create Story Spiral"
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
      Left            =   2880
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuiz 
      Caption         =   "Save Story"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox TxtName 
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
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit Story"
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
      Left            =   360
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2160
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
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
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox TxtFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmmine.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox CheckSpeed 
         Caption         =   "Read one word at a time"
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
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox CheckWord 
         Caption         =   "Display Word Balloons"
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
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Read Story"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblSave 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
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
      Left            =   5280
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblQuiz 
      AutoSize        =   -1  'True
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
      Left            =   6480
      TabIndex        =   16
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "StoryName"
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
      Left            =   5160
      TabIndex        =   15
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "<-Enter a new story name to create a new story"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   9
      Top             =   360
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Story Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   1080
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   2400
      Top             =   0
   End
End
Attribute VB_Name = "FrmStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iMyClick As Integer
Dim FileNum As Integer
Dim Question(40) As String
Dim Answer(40, 5) As String
Dim MyAnswer As String
Dim YesNo(5) As String
Dim Ans(5) As Integer
Dim Q As Integer
Dim Guess As Integer
Dim Score As Integer
Dim Questions As Integer
Dim Finalscore As String
Dim Level As String
Dim Response As Integer
Dim Search As Integer
Dim FileLoc As String
Dim bOneword As Boolean
Sub Rest()
 Dim n As Integer
  n = Int((6 * Rnd) + 1)
  'Text1.Text = Text1.Text & " " & N
  Select Case n
        Case 1
           MyAgent.Play "Idle1_1"
        Case 2
          MyAgent.Play "Idle1_2"
        Case 3
          If Anim = "Peedy" Then
             MyAgent.Play "Blink"
             MyAgent.Play "Idle1_1"
          Else
            MyAgent.Play "Idle1_3"
          End If
        Case 4
          MyAgent.Play "Idle1_4"
        Case 5
          MyAgent.Play "Blink"
          MyAgent.Play "Blink"
        Case 6
        If Anim = "Merlin" Then
          MyAgent.Play "Idle2_1"
        Else
          MyAgent.Play "Idle1_4"
        End If
      End Select
      
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

Private Sub CheckSpeed_Click()
 If CheckWord.Value Then
       bOneword = True
   Else
        bOneword = False
   End If
End Sub

Private Sub CheckWord_Click()
    If CheckWord.Value Then
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
   
    Else
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
   
   End If
End Sub

Private Sub CmdEdit_Click()
    MyAgent.StopAll
    bHide = True
    MyAgent.Hide
    FrmTeacher.Show
End Sub

Private Sub CmdExit_Click()
 End
End Sub

Private Sub CmdQuiz_Click()
'
' Update quiz text file
   If TxtFile.Text = "" Then Exit Sub
   If TxtName.Text = "" Then Exit Sub
   File1.Path = Dir1.Path
   File1.Pattern = "*.txt"
   If Len(cQuizfile) > 0 Then
   'Dim MyString As String
   cQuizfile = "c:/kids/" & TxtName.Text & "Story.txt"
   Open cQuizfile For Output As #1 ' Open file for output.
   Print #1, TxtFile.Text
   Close #1    ' Close file.
  End If
   File1.Path = Dir1.Path
   File1.Pattern = "*story.txt"
   LblSave.Caption = TxtName.Text & " Saved."
   LblSave.Visible = True
End Sub

Private Sub CmdSpiral_Click()
   FrmSpiral.Show
End Sub

Private Sub CmdStart_Click()
    If bHide Then
            MyAgent.Show
            bHide = False
     End If
     MyAgent.StopAll
     Dim n As Integer
      n = Int((40 * Rnd) + 1)
      n = n Mod 4 + 1
      Select Case n
        Case 1
       MyAgent.Play "Greet"
       MyAgent.Speak "Hello "
       Case 2
       MyAgent.Play "Acknowledge"
       MyAgent.Speak "How are you "
       Case 3
       MyAgent.Play "Wave"
       MyAgent.Speak "Hi  "
       Case 4
       MyAgent.Play "Pleased"
       MyAgent.Speak "I am glad to see you "
     End Select
     File1.Visible = True
     Sleep 80
      MyAgent.Speak "I am " & Anim
     MyAgent.Play "RestPose"
     
     n = Int((90 * Rnd) + 1)
     n = n Mod 9 + 1
     Select Case n
        Case 1
          MyAgent.Play "Process"
          MyAgent.Speak "Are you ready for a story?"
        Case 2
          MyAgent.Play "DoMAgic1"
          MyAgent.Speak "Here is a story"
        Case 3
          MyAgent.Play "Think"
          MyAgent.Speak "I think I found a story"
        Case 4
           MyAgent.Play "Search"
           MyAgent.Speak "I am searching for a story"
        Case 5
          MyAgent.Play "Read"
          MyAgent.Speak "I found a good book"
        Case 6
          MyAgent.Play "Write"
          MyAgent.Speak "Once upon a time"
        Case 7
          MyAgent.Play "Announce"
          MyAgent.Speak "I will read you a story"
        Case 8
          MyAgent.Play "Suggest"
          MyAgent.Speak "I have an idea. I will read to you"
        Case 9
         ' MyAgent.Play "Suggest"
          MyAgent.Think "What story would you like to hear?"
     End Select
     MyAgent.Play "Blink"
     MyAgent.Speak "Click on the story you would like to hear."
End Sub






Private Sub Command1_Click()
  File1.Path = Dir1.Path
  File1.Pattern = "*story.txt"
End Sub

Private Sub Command2_Click()
'  'MyAgent.Play "MoveRight"
'  'MyAgent.MoveTo 200, 350
''MyAgent.moverirgt
'  Text1.Text = iMyClick
'  Select Case iMyClick
'        Case 1
'           MyAgent.Play "Idle1_1"
'        Case 2
'          MyAgent.Play "Idle1_2"
'        Case 3
'          If Anim = "Peedy" Then
'             MyAgent.Play "Blink"
'             MyAgent.Play "Idle1_1"
'          Else
'            MyAgent.Play "Idle1_3"
'          End If
'         'MyAgent.Play "Idle1_3"
'
'        Case 4
'          MyAgent.Play "Idle1_4"
'        Case 5
'          MyAgent.Play "Blink"
'          MyAgent.Play "Blink"
'        Case 5
'        If Anim = "Merlin" Then
'          MyAgent.Play "Idle2_1"
'        Else
'          MyAgent.Play "Idle1_4"
'        End If
''        Case 1
''           MyAgent.MoveTo 400, 350
''        Case 2
''           MyAgent.MoveTo 200, 350
''        Case 3
''          MyAgent.MoveTo 500, 350
''        Case 4
''          MyAgent.MoveTo 400, 250
''        Case 5
''           MyAgent.MoveTo 400, 450
'      End Select
'       iMyClick = iMyClick + 1
'      If iMyClick > 5 Then iMyClick = 1
End Sub



Private Sub Command3_Click()

End Sub

Private Sub Command6_Click()
   Dim cTxt, cMY As String
   Dim bRest As Boolean
   Dim bFirst As Boolean
   Dim i, L As Integer
  ' MyAgent.Speak Label2.Caption
   bRest = True
    bFirst = True
    Dim n As Integer
   n = Int((3 * Rnd) + 1)
   Select Case n
    Case 1
      MyAgent.Play "Read"
    Case 2
     MyAgent.Play "Announce"
    Case 3
     MyAgent.Play "Wave"
   End Select
   cTxt = ""
   For i = 1 To Len(TxtFile.Text)
    cMY = Mid(TxtFile.Text, i, 1)
    If bOneword Then
        If cMY > " " Then
               cTxt = cTxt & cMY
        '       If bOneword Then
        '       Else
            Else
               If cTxt <> "" Then MyAgent.Speak cTxt
                cTxt = ""
                Sleep 400
          End If
    
    Else
        If cMY >= " " Then
           cTxt = cTxt & cMY
    '       If bOneword Then
    '       Else
           If cMY = "." Or cMY = "?" Or cMY = "!" Then
             If cTxt <> "" Then MyAgent.Speak cTxt
             'If bFirst Then
             
             'End If
              cTxt = ""
              If bRest Then
                 'Call Rest
                 bRest = False
              Else
                 bRest = True
              End If
           End If
        Else
        End If
   End If
     
   Next
   
    If cTxt <> "" Then MyAgent.Speak cTxt
    Sleep 1800
     n = Int((3 * Rnd) + 1)
   Select Case n
    Case 1
     MyAgent.Speak "The end"
    Case 2
     MyAgent.Speak "That is all folks"
    Case 3
     MyAgent.Speak "Another Story?"
   End Select
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
  File1.Pattern = "*story.txt"
End Sub

Private Sub File1_Click()
 MyAgent.StopAll
  LblSave.Visible = False
 Dim L As Integer
 Command6.Visible = True
TxtFile.Visible = True
TxtFile.Text = ""
Dim cQuizfile As String
Dim MyString As String
If bHide Then
    MyAgent.Show
     bHide = False
 End If
cQuizfile = File1.Path & "\" & File1.FileName
 cQuizfile = File1.Path & "\" & File1.FileName
   LblQuiz.Caption = cQuizfile
   L = Len(cQuizfile)
   TxtName.Text = Mid(cQuizfile, 9, L - 17)
Dim bFirst As Boolean
 bFirst = True
Open cQuizfile For Input As #1 ' Open file for input.
  Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
    'Debug.Print MyString,  ' Print data to Debug window.
    If bFirst Then
     ' Label2.Caption =
     TxtFile.Text = MyString & Chr(13) & Chr(10) '& " " & Chr(13) & Chr(10)
     'TxtFile.Text = " " & Chr(13) & Chr(10)
      bFirst = False
    Else
      TxtFile.Text = TxtFile.Text & MyString & Chr(13) & Chr(10)
    End If
    strStory = strStory & MyString
    CmdSpiral.Visible = True
    Loop
   Close #1    ' Close file.
   Sleep 600
  Command6_Click
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Search = 1
CheckWord.Value = 1
 bOneword = False
Dim DirName As String
Dir1.Path = App.Path '"C:\kids"
strStory = ""
'chec
'If App.PrevInstance = True Then
'    MsgBox "This application is already running!", vbInformation + vbOKOnly, "Motor Vehicle Quiz is Running"
'    End
'End If
'If Right(App.Path, 1) = "\" Then
'    Ap = App.Path
'Else
'    Ap = App.Path & "\"
'End If
'
CentreMe Me

Command1_Click
Exit Sub
ErrHandler:
MsgBox Err.Description

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

MyAgent.StopAll
If Q = 0 Then
    MyAgent.Speak "You haven't heard the story yet. Why are you leaving?"
End If
Sleep 6000
MyAgent.StopAll
Set MyAgent = Nothing
FrmStart.Agent1.Characters.Unload Anim
End
End Sub



