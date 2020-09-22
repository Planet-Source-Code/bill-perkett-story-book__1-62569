VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Jokes"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Answer"
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
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Question"
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
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LblAnswer 
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
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label LblQuestion 
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jokes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   825
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCnt As Integer
Dim Question As String
Dim Answer As String
Private Sub Command1_Click()
  Dim ival As Integer
 Dim j, L As Integer
  Dim cMsg As String
  
PickQuestion:
 'iCnt = iCnt + 1
 ival = Int((Rnd * iQuizcnt) + 1)
 If ival > iQuizcnt Then ival = 1
 j = Len(cQuiz(ival, 0))
 
 If cQuiz(ival, 5) = "Y" And iCnt < iQuizcnt Then GoTo PickQuestion
 cQuiz(ival, 5) = "Y"
 'For j = 1 To iQuizcnt
  L = Len(cQuiz(ival, 0))
  iCnt = iCnt + 1
  LblQuestion.Caption = cQuiz(ival, 0)
  LblAnswer.Caption = cQuiz(ival, 1)
  LblAnswer.Visible = False
  Question = Right(cQuiz(ival, 0), L - 10)
   L = Len(cQuiz(ival, 1))
  Answer = Right(cQuiz(ival, 1), L - 8)
  MyAgent.Speak "Question:"
  MyAgent.Speak Question
  Command2.SetFocus
End Sub

Private Sub Command2_Click()
  MyAgent.Speak "Answer:"
  Sleep 50
  MyAgent.Speak Answer
  LblAnswer.Visible = True
  Command1.SetFocus
End Sub

Private Sub Form_Load()

 
 Dim cMsg As String
  cMsg = App.Path & "\joke.txt"
  Call ReadQuiz(cMsg)
 
iCnt = 0

End Sub
