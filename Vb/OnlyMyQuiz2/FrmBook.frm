VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmBook 
   Caption         =   "FrmBook"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "FrmBook"
   ScaleHeight     =   8640
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   34
      Text            =   "000"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   315
      Left            =   6960
      TabIndex        =   33
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   8400
   End
   Begin VB.CommandButton ComGo 
      Caption         =   "Go to Picture Start"
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
      Left            =   0
      TabIndex        =   32
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Picture Start"
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
      Left            =   0
      TabIndex        =   31
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Picture Action"
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
      Left            =   0
      TabIndex        =   30
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox ComPictureAction 
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Text            =   "Combo1"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   870
      Left            =   7920
      TabIndex        =   26
      Top             =   7680
      Width           =   2295
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
      Left            =   3720
      TabIndex        =   24
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Do Action Now"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Play sound Now"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6120
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox ComPicture 
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox ComBack 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "Add Action"
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
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   1695
   End
   Begin VB.ComboBox ComAction 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton CmdRead 
      Caption         =   "Read"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdSound 
      Caption         =   "Add Sound"
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
      Left            =   0
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
   End
   Begin VB.ComboBox ComSound 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   6480
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox TxtPicture 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Text            =   "0"
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox RichTextBox1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   4080
         Width           =   8415
      End
      Begin VB.PictureBox Picture2 
         Height          =   3735
         Left            =   0
         Picture         =   "FrmBook.frx":0000
         ScaleHeight     =   3675
         ScaleWidth      =   8355
         TabIndex        =   13
         Top             =   240
         Width           =   8415
         Begin VB.Image Image1 
            Height          =   375
            Index           =   9
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   8
            Left            =   1440
            Stretch         =   -1  'True
            Top             =   120
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   7
            Left            =   840
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   6
            Left            =   1080
            Stretch         =   -1  'True
            Top             =   120
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   5
            Left            =   720
            Stretch         =   -1  'True
            Top             =   120
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   4
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   3
            Left            =   480
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   1
            Left            =   360
            Stretch         =   -1  'True
            Top             =   120
            Width           =   135
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   135
         End
      End
      Begin VB.TextBox TxtStory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3960
         Width           =   8415
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Picture Name"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Click on a name to load a book "
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
      Left            =   7680
      TabIndex        =   28
      Top             =   7440
      Width           =   2730
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Book Name"
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
      Left            =   3840
      TabIndex        =   25
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   8280
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Picture Height"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Picture Width"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "BackGround"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   7320
      Top             =   7320
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Image1m(20) As Boolean
Dim bApicture As Boolean
Dim bStart As Boolean
Dim iPic As Integer
Private agAgent As IAgentCtlCharacterEx
Dim Request As Object
' Image Picture Name
Dim strName(20) As String
' Image Picture Actions
'Dim strMovie(2000, 5) As String
' Image Picture Start
Dim strStart(20, 6) As String
'Dim iMovie As Integer
Dim bFirst As Boolean
Private Sub FindSounds()
  '
  ' Find Sounds
  '
    File1.Path = Dir1.Path
    File1.Pattern = "*.txt"
    File1.Path = Dir1.Path
    File1.Pattern = "*.wav"
End Sub
Private Sub FindPicture()
  '
  ' Find Picture
  '
    File1.Path = Dir1.Path
    File1.Pattern = "*.txt"
    File1.Path = Dir1.Path
    File1.Pattern = "*.wmf"
End Sub
Private Sub FindBooks()
  '
  ' Find Books
  '
    File1.Path = Dir1.Path
    File1.Pattern = "*.txt"
    File1.Path = Dir1.Path
    File1.Pattern = "*book.txt"
End Sub

Private Sub ExecutePicture(strCommand As String)
  Dim i, j As Integer
If UCase((Mid(strCommand, 1, 4))) = "MOVE" Then
    j = Val(Mid(strCommand, 5, 3))
   For i = 0 To 9
         j = j + 1
'      iMovie = iMovie + 1
      If strName(i) <> "" And Len(strMovie(j, 0)) > 20 Then
        If Image1(i).Width <> strMovie(j, 1) Then Image1(i).Width = strMovie(j, 1)
        If Image1(i).Height <> strMovie(j, 1) Then Image1(i).Height = strMovie(j, 2)
        If Image1(i).Top <> strMovie(j, 1) Then Image1(i).Top = strMovie(j, 3)
        If Image1(i).Left <> strMovie(j, 1) Then Image1(i).Left = strMovie(j, 4)
     End If
   Next
  End If
  If UCase((Mid(strCommand, 1, 4))) = "HIDE" Then
    i = Val(Mid(strCommand, 5, 2)) - 1
     Image1(i).Visible = False
  End If
   If UCase((Mid(strCommand, 1, 4))) = "SHOW" Then
    i = Val(Mid(strCommand, 5, 2)) - 1
     Image1(i).Visible = True
  End If
End Sub
Private Sub ExecuteCommand(strCommand As String)
    Dim strData As String
    Dim booltemp As Boolean
    Dim bNoExist As Boolean
    Dim bRequest As Boolean
    Dim strInstructions As String
    Dim intStartPosition As Integer 'Calculated Start Position for parsing string
    Dim strAgent As String
    Dim intCount As Integer
    Dim i As Integer
    Dim strTemp As String
    Dim strCommandIn As String
    Dim X, Y As Integer
    'Dim Request As Object
    Dim PauseTime, Start, Finish, TotalTime As Long
    Dim cMsg As String
    cMsg = strCommand
    If Mid(UCase(cMsg), 1, 4) = "SIZE" Then
     If UCase(cMsg) = "SIZE 3" Then strCommand = "/SIZE " & cStoryAgent & " 3"
     If UCase(cMsg) = "SIZE 2" Then strCommand = "/SIZE " & cStoryAgent & " 2"
     If UCase(cMsg) = "SIZE 1" Then strCommand = "/SIZE " & cStoryAgent & " 1"
    ElseIf Mid(UCase(cMsg), 1, 4) = "MOVE" Then
       strCommand = "/" & cMsg & " " & cStoryAgent & " " '& strCommand
    Else
       strCommand = "/PLAY " & cStoryAgent & " " & strCommand
    End If
    strData = strCommand
    bRequest = False
    bRequestDone = False
    strCommandIn = strCommand
    On Error GoTo ExitError
    'Debug.Print strData
    Static result As IAgentCtlRequest
    If Left(strData, 1) = "/" And Len(strData) > 1 Then
        If InStr(1, strData, " ") = 0 Then
            strCommand = Mid(strData, 2)
            strInstructions = ""
        Else
            strCommand = UCase(Mid(strData, 2, InStr(1, strData, " ") - 2))
            strInstructions = Mid(strData, InStr(1, strData, " ") + 1)
        End If
        Select Case strCommand
            Case "THINK"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.Think(strTemp)
                 bRequest = True
            Case "SOUND"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
               ' agAgent.StopAll
                If strTemp = "OFF" Then
                  agAgent.SoundEffectsOn = False
                Else
                agAgent.SoundEffectsOn = True
                End If
                PauseTime = 1   ' Set duration.
                   For i = 1 To 2
                    Start = Timer   ' Set start time.
                    Do While Timer < Start + PauseTime
                        DoEvents    ' Yield to other processes.
                    Loop
                    Finish = Timer  ' Set end time.
                    Next
                   List1.AddItem "D " & strCommandIn
                 
            Case "SAY"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.Speak(strTemp)
                bRequest = True
            Case "GESTUREAT"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Trim(Mid(strInstructions, InStr(1, strInstructions, " ") + 1))
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left ' / 15
                Y = agAgent.Top '- 150 '/ 15
                Text1.Text = strTemp & " " & X & " " & Y
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                 Set Request = agAgent.GestureAt(X, Y)
                  bRequest = True
           Case "POINT"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Trim(Mid(strInstructions, InStr(1, strInstructions, " ") + 1))
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left ' / 15
                Y = agAgent.Top '- 150 '/ 15
                Text1.Text = strTemp & " " & X & " " & Y
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.GestureAt(X, Y)
                bRequest = True
             Case "MOVELEFT"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                 Set agAgent = Me.MyAgent.Characters(strAgent)
                X = agAgent.Left + 200 ' / 15
                Y = agAgent.Top '- 150 '/ 15
                agAgent.StopAll
                Set Request = agAgent.MoveTo(X, Y)
               bRequest = True
            Case "MOVERIGHT"
               strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                X = agAgent.Left - 200 ' / 15
                Y = agAgent.Top '- 150 '/ 15
                 Set Request = agAgent.MoveTo(X, Y)
               bRequest = True
            Case "MOVEDOWN"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                X = agAgent.Left '+ 100 ' / 15
                Y = agAgent.Top + 150 '/ 15
                Set Request = agAgent.MoveTo(X, Y)
                bRequest = True
            Case "MOVEUP"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                X = agAgent.Left '+ 100 ' / 15
                Y = agAgent.Top - 150 '/ 15
                Set Request = agAgent.MoveTo(X, Y)
               bRequest = True
            Case "GESTURE"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                On Error Resume Next
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.Play(strTemp)
                bRequest = True
           Case "PLAY"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                On Error Resume Next
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.Play(strTemp)
                bRequest = True
            Case "HIDE"
                If strInstructions <> "" Then
                    Set agAgent = Me.MyAgent.Characters(strInstructions)
                    agAgent.StopAll
                    agAgent.Hide
                End If
               List1.AddItem "D " & strCommandIn
          Case "SHOW"
             If iCharCnt = 0 Then
                 MyAgent.Characters.Load strInstructions, strInstructions & ".acs"
                 iCharCnt = iCharCnt + 1
                 cCharacters(iCharCnt) = strInstructions
                 Set agAgent = MyAgent.Characters(strInstructions)
                 Set Request = agAgent.Show
                 bRequest = True
             Else
                bNoExist = True
                For i = 1 To iCharCnt
                  If UCase(cCharacters(i)) = UCase(strInstructions) Then bNoExist = False
                Next
                If bNoExist Then
                 MyAgent.Characters.Load strInstructions, strInstructions & ".acs"
                 iCharCnt = iCharCnt + 1
                 cCharacters(iCharCnt) = UCase(strInstructions)
                End If
                 Set agAgent = MyAgent.Characters(strInstructions)
                 agAgent.StopAll
                 Set Request = agAgent.Show
                 bRequest = True
             End If
             
'            Case "LISTAGENTS"
'                ListAgents intIndex
'            Case "LISTACTIVEAGENTS"
'                ListActiveAgents intIndex
'            Case "LISTGESTURES"
'                ListGestures intIndex, strInstructions
'            Case "LOAD"
'                LoadAgent Left(strInstructions, InStr(1, strInstructions, " ") - 1), Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
'                ListActiveAgents intIndex
            Case "UNLOAD"
                If iCharCnt = 0 Then Exit Sub
                 bNoExist = True
                For i = 1 To iCharCnt
                  If UCase(cCharacters(i)) = UCase(strInstructions) Then
                   Set agAgent = Me.MyAgent.Characters(strInstructions)
                   Me.MyAgent.Characters.Unload (strInstructions)
                   cCharacters(i) = ""
                  End If
                Next
               List1.AddItem "D " & strCommandIn
'                Set agAgent = Me.MyAgent.Characters(strInstructions)
'                Me.MyAgent.Characters.Unload (strInstructions)
'                ListActiveAgents intIndex
            Case "PAUSE"
'                PauseTime = 4   ' Set duration.
'                Start = Timer   ' Set start time.
'                Do While Timer < Start + PauseTime
'                   DoEvents    ' Yield to other processes.
'                 Loop
'                 Exit Sub
                   PauseTime = 1   ' Set duration.
                   For i = 1 To 8
                   If bRequestDone = False Then 'Request is in Queue
                    Start = Timer   ' Set start time.
                    Do While Timer < Start + PauseTime
                        DoEvents    ' Yield to other processes.
                    Loop
                    Finish = Timer  ' Set end time.
                    
            '     'Add your code here (you can send text to status bar or something)
            '        Debug.Print i, Request.Status
                   Else 'Request successfully completed
                    Start = Timer   ' Set start time.
                    Do While Timer < Start + PauseTime
                        DoEvents    ' Yield to other processes.
                    Loop
                    Finish = Timer
                     'agAgent.StopAll
                     List1.AddItem "D " & strCommandIn
                 'Add your code here (you can do something like display the annimation)
                      Exit Sub
                    End If
                    Next
                    List1.AddItem "N " & strCommandIn
                    Exit Sub
           Case "MOVE"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                On Error Resume Next
                Set agAgent = Me.MyAgent.Characters(strAgent)
                i = InStr(1, strTemp, ",")
                If i = 0 Then i = InStr(1, strTemp, " ")
                agAgent.StopAll
                Set Request = agAgent.MoveTo(CDbl(Mid(strTemp, 1, i - 1)), CDbl(Mid(strTemp, i + 1, Len(strTemp) - i)))
                bRequest = True
          Case "SIZE"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                If Trim(strTemp) = "3" Then
                  agAgent.Height = "240" 'Sets the Characters Height
                  agAgent.Width = "360" 'Sets the Characters Width
                  'Set Request = agAgent.Width("360")
                  'bRequest = True
                ElseIf Trim(strTemp) = "1" Then '
                   agAgent.Height = "64" 'Sets the Characters Height
                   agAgent.Width = "80"
                   ';Set Request = agAgent.Width("80")
                   'bRequest = True  'Sets the Characters Width
                Else
                   agAgent.Height = "128"
                   agAgent.Width = "160"
                  'Set Request = agAgent.Width("160")
              End If
              'List1.AddItem "D " & strCommandIn
         Case "STOP"
                strAgent = strInstructions
                Set agAgent = Me.MyAgent.Characters(strAgent)
                 agAgent.StopAll
                 PauseTime = 1   ' Set duration.
           For i = 1 To 3
            Start = Timer   ' Set start time.
            Do While Timer < Start + PauseTime
                DoEvents    ' Yield to other processes.
            Loop
            Finish = Timer  ' Set end time.
          Next
            'List1.AddItem "D " & strCommandIn
        End Select
    End If
   If bRequest Then
       PauseTime = 1   ' Set duration.
       For i = 1 To 8
       
       If bRequestDone = False Then 'Request is in Queue
          Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer  ' Set end time.
        
'     'Add your code here (you can send text to status bar or something)
'        Debug.Print i, Request.Status
       Else 'Request successfully completed
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
           Finish = Timer
'          CharPosn(0) = agAgent.Left
'          CharPosn(1) = agAgent.Top
          ' List1.AddItem "D " & strCommandIn
          Exit Sub
        End If
       Next
        'gAgent.StopAll
        'List1.AddItem "N " & strCommandIn
'        CharPosn(0) = agAgent.Left
'        CharPosn(1) = agAgent.Top
    End If
    Exit Sub
ExitError:
      'List1.AddItem "E " & strCommandIn
End Sub
Private Sub Myspeak(cTxtin As String)
 Dim PauseTime, Start, Finish, TotalTime As Long
  Static result As IAgentCtlRequest
    Set agAgent = MyAgent.Characters(cStoryAgent)
   agAgent.StopAll
    Set Request = agAgent.Speak(cTxtin)
    bRequestDone = False
        PauseTime = 1   ' Set duration.
       For i = 1 To 8
       
       If bRequestDone = False Then 'Request is in Queue
          Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer  ' Set end time.
        
'     'Add your code here (you can send text to status bar or something)
'        Debug.Print i, Request.Status
       Else 'Request successfully completed
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
           Finish = Timer
          Label2.Caption = "D " & cTxt
          Exit Sub
        End If
       Next
        'gAgent.StopAll
        Label2.Caption = "N " & cTxt
    
End Sub
Private Sub MySound(cFile As String)
Dim PauseTime, Start, Finish, TotalTime As Long
  Label2.Caption = ""
  PauseTime = 1   ' Set duration.
Dim i, j, JOLD As Integer
  With MMControl1
     .Wait = False
     .Shareable = False
     .DeviceType = "Sequencer"
      .Command = "close"
   End With
   MMControl1.FileName = soundpath & cFile & ".wav" 'App.Path & "\applause.wav" '" & iMyicon & ".MID"
          'If iSong = 3 Then
            'frmOpen.MMControl1.FileName = App.Path & "\applause.wav" '" & iMyicon & ".MID"
          ' Else
            ' MMControl1.FileName = App.Path & "\0.MID"
          'End If
   MMControl1.DeviceType = "WaveAudio"
   MMControl1.Command = "Open"
   MMControl1.Command = "play"
   JOLD = 0
   For i = 1 To 15
      Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer
     j = MMControl1.Position
     If j = JOLD Then
        Label2.Caption = "DS"
        Exit Sub
     End If
     JOLD = j
   Next
End Sub


Private Sub Agent1_ActivateInput(ByVal CharacterID As String)
 If Request Then
      bStory = True

 End If
End Sub

Private Sub CmdAction_Click()
Dim cMsg As String
cMsg = "/PLAY " & cStoryAgent & " " & ComAction.Text
    TxtStory = TxtStory & " {" & ComAction.Text & "}"
    ExecuteCommand (ComAction.Text)
End Sub




Private Sub CmdRead_Click()
  Label2.Caption = ""
  
  Dim cTxt, cMY As String
   Dim bRest As Boolean
   Dim bFirst As Boolean
   Dim i, j, L, jPic As Integer
   'Call Command3_Click
   cTxt = "SIZE 2"
   ExecuteCommand (cTxt)
    '
    ' Set Starting Spot
    '
    For i = 0 To 9
    If strName(i) <> "" Then
      If strStart(i, 1) = "HIDE" Then
          strStart(i, 1) = "SHOW"
          strStart(i, 2) = Image1(i).Width
          strStart(i, 3) = Image1(i).Height
          strStart(i, 4) = Image1(i).Top
          strStart(i, 5) = Image1(i).Left
          jPic = 0
          If iMovie > 0 Then
             For j = 1 To iMovie
               If jPic = i Then
                 cTxt = strMovie(j, 0) '= "00000000000000000000" & strName(i)
                 strMovie(j, 1) = strStart(i, 2)
                 strMovie(j, 2) = strStart(i, 3)
                 strMovie(j, 3) = strStart(i, 4)
                 strMovie(j, 4) = strStart(i, 5)
               End If
                 jPic = jPic + 1
                If jPic > 9 Then jPic = 0
    
               Next
            End If
      End If
       Image1(i).Visible = True
       'ExecutePicture (strStart(i, 0))
       Image1(i).Width = strStart(i, 2)
       Image1(i).Height = strStart(i, 3)
       Image1(i).Top = strStart(i, 4)
       Image1(i).Left = strStart(i, 5)
     End If
    Next
    '
    '
    '
  ' MyAgent.Speak Label2.Caption
   bRest = True
    bFirst = True
    Dim n As Integer
'   n = Int((3 * Rnd) + 1)
'   Select Case n
'    Case 1
'      MyAgent.Play "Read"
'    Case 2
'     MyAgent.Play "Announce"
'    Case 3
'     MyAgent.Play "Wave"
'   End Select
   cTxt = ""
   For i = 1 To Len(TxtStory.Text)
    cMY = Mid(TxtStory.Text, i, 1)
      If cMY = "[" Then
        If cTxt <> "" Then Myspeak (cTxt)
        cTxt = ""
      ElseIf cMY = "]" Then
          MySound (cTxt)
          cTxt = ""
      ElseIf cMY = "{" Then
        If cTxt <> "" Then Myspeak (cTxt)
        cTxt = ""
        ElseIf cMY = "}" Then
          ExecuteCommand (cTxt)
          cTxt = ""
        ElseIf cMY = "(" Then
        If cTxt <> "" Then Myspeak (cTxt)
        cTxt = ""
        ElseIf cMY = ")" Then
          ExecutePicture (cTxt)
          cTxt = ""
      Else
         cTxt = cTxt & cMY
      End If
  
   Next
   If cTxt <> "" Then
      Myspeak (cTxt)
    End If
End Sub

Private Sub CmdSave_Click()
 Dim cQuizfile As String
 Dim cMsg As String
 Dim i, jHeight, jTop, jLeft, jWidth As Integer
 Dim j1, j2, j3, j4, jPic As Integer
  For i = 0 To 9
    If strName(i) <> "" Then
      If strStart(i, 1) = "HIDE" Then
          strStart(i, 1) = "SHOW"
          strStart(i, 2) = Image1(i).Width
          strStart(i, 3) = Image1(i).Height
          strStart(i, 4) = Image1(i).Top
          strStart(i, 5) = Image1(i).Left
          jPic = 0
          If iMovie > 0 Then
             For j = 1 To iMovie
               If jPic = i Then
                 strMovie(j, 1) = strStart(i, 2)
                 strMovie(j, 2) = strStart(i, 3)
                 strMovie(j, 3) = strStart(i, 4)
                 strMovie(j, 4) = strStart(i, 5)
               End If
                 jPic = jPic + 1
                If jPic > 9 Then jPic = 0
    
               Next
            End If
      End If
     End If
 Next
'
' Update quiz text file
   If TxtStory.Text = "" Then Exit Sub
'   If TxtName.Text = "" Then Exit Sub
'   File1.Path = Dir1.Path
'   File1.Pattern = "*.txt"
'   If Len(cQuizfile) > 0 Then
   'Dim MyString As String
   cQuizfile = "c:/kids/" & TxtName.Text & "book.txt"
   Open cQuizfile For Output As #1 ' Open file for output.
   Print #1, ComBack.Text
   For i = 0 To 9
       j1 = Val(strStart(i, 2)) 'Image1(i).Width
       j2 = Val(strStart(i, 3)) ' Image1(i).Height
       j3 = Val(strStart(i, 4)) 'Image1(i).Top
       j4 = Val(strStart(i, 5)) 'Image1(i).Left
       cMsg = Format(j1, "00000") & Format(j2, "00000")
       cMsg = cMsg & Format(j3, "00000") & Format(j4, "00000") & strName(i)
      Print #1, cMsg
   Next
   j1 = Val(iMovie)
   cMsg = Format(j1, "00000") & Format(j1, "00000")
   cMsg = cMsg & Format(j1, "00000") & Format(j1, "00000") & "iMovie"
   Print #1, cMsg
       If iMovie > 0 Then
         jPic = 0
         For i = 1 To iMovie
            j1 = Val(strMovie(i, 1)) 'Image1(i).Width
            j2 = Val(strMovie(i, 2)) ' Image1(i).Height
            j3 = Val(strMovie(i, 3)) 'Image1(i).Top
            j4 = Val(strMovie(i, 4)) 'Image1(i).Left
            cMsg = Format(j1, "00000") & Format(j2, "00000")
            cMsg = cMsg & Format(j3, "00000") & Format(j4, "00000") & strName(jPic)
            jPic = jPic + 1
           
            If jPic > 9 Then
               jPic = 0
            End If
            Print #1, cMsg
        Next
   End If
   Print #1, TxtStory.Text
   Close #1    ' Close file.
   Command1_Click
  'End If
'   File1.Path = Dir1.Path
'   File1.Pattern = "*story.txt"
'   LblSave.Caption = TxtName.Text & " Saved."
'   LblSave.Visible = True
End Sub

Private Sub CmdSound_Click()
    If bStart = False Then Command3_Click
    TxtStory = TxtStory & " [" & ComSound.Text & "]"
    MySound (ComSound.Text)
End Sub

Private Sub ComAction_Click()
 'If bStart = False Then Command3_Click
 If Check2 Then ExecuteCommand (ComAction.Text)
End Sub

Private Sub ComBack_Change()
'  If ComBack.Text = "Stream" Then Picture2.Picture = LoadPicture(soundpath & bg1.jpg)
'   If ComBack.Text = "Field" Then Picture2.Picture = LoadPicture(soundpath & bg2.jpg)
'    If ComBack.Text = "Desert" Then Picture2.Picture = LoadPicture(soundpath & bg3.jpg)
'     If ComBack.Text = "Ocean" Then Picture2.Picture = LoadPicture(soundpath & bg4.jpg)
End Sub

Private Sub ComBack_Click()
 If ComBack.Text = "Stream" Then Picture2.Picture = LoadPicture(soundpath & "bg1.jpg")
' If ComBack.Text = "Field" Then Picture2.Picture = LoadPicture(soundpath & "bg2.jpg")
' If ComBack.Text = "Desert" Then Picture2.Picture = LoadPicture(soundpath & "bg3.jpg")
' If ComBack.Text = "Ocean" Then Picture2.Picture = LoadPicture(soundpath & "bg4.jpg")
' If ComBack.Text = "Rocks" Then Picture2.Picture = LoadPicture(soundpath & "bg5.jpg")
' If ComBack.Text = "Jungle" Then Picture2.Picture = LoadPicture(soundpath & "bg6.jpg")

End Sub



Private Sub ComGo_Click()
 Dim i As Integer
   bApicture = False
 For i = 0 To 9
     Image1m(i) = False
     Image1(i).BorderStyle = 0
     bApicture = False
     If strName(i) <> "" Then
       'ExecutePicture (strStart(i, 0))
       Image1(i).Visible = True
       Image1(i).Width = strStart(i, 2)
       Image1(i).Height = strStart(i, 3)
       Image1(i).Top = strStart(i, 4)
       Image1(i).Left = strStart(i, 5)
       If i = 2 Then
         Image1(i).Visible = True
       End If
     End If
    Next
End Sub


Private Sub Command2_Click()
   Dim i As Integer
   If bStart = False Then Command3_Click
    If UCase(ComPictureAction.Text) = "SHOW" Then TxtStory = TxtStory & " (" & ComPictureAction.Text & Format(HScroll1.Value + 1, "00") & ")"
    If UCase(ComPictureAction.Text) = "HIDE" Then TxtStory = TxtStory & " (" & ComPictureAction.Text & Format(HScroll1.Value + 1, "00") & ")"
    If UCase(ComPictureAction.Text) = "MOVE" Then
     TxtStory = TxtStory & " (" & ComPictureAction.Text & Format(iMovie, "000") & ")"
     For i = 0 To 9
       iMovie = iMovie + 1
       strMovie(iMovie, 0) = "00000000000000000000" & strName(i)
       strMovie(iMovie, 1) = Image1(i).Width
       strMovie(iMovie, 2) = Image1(i).Height
       strMovie(iMovie, 3) = Image1(i).Top
       strMovie(iMovie, 4) = Image1(i).Left
    Next
     ' TxtStory = TxtStory & " (" & ComPictureAction.Text & Format(HScroll1.Value + 1, "00") & Format(HScroll2.Value, "00") & Format(HScroll3.Value, "00") & ")"
    End If
    ' Else
'     If UCase(ComPictureAction.Text) = "SHOW" Then strStart(TxtPicture.Text, 0) = "SHOW" & Format(HScroll1.Value + 1, "00")
'     If UCase(ComPictureAction.Text) = "HIDE" Then strStart(TxtPicture.Text, 0) = "HIDE" & Format(HScroll1.Value + 1, "00")
 'End If
End Sub

Private Sub Command3_Click()
  Dim i, j2, j3, j2v, j3v, jPic As Integer
  bStart = True
  RichTextBox1.Visible = False
  For i = 0 To 9
     If strName(i) <> "" Then
       If strStart(i, 1) = "HIDE" Then
          jPic = 0
          If iMovie > 0 Then
             For j = 1 To iMovie
               If jPic = i Then
                 strMovie(j, 1) = strStart(i, 2)
                 strMovie(j, 2) = strStart(i, 3)
                 strMovie(j, 3) = strStart(i, 4)
                 strMovie(j, 4) = strStart(i, 5)
               End If
                 jPic = jPic + 1
                If jPic > 9 Then jPic = 0
    
               Next
            End If
        End If
         strStart(i, 1) = "SHOW"
        strStart(i, 2) = Image1(i).Width
        strStart(i, 3) = Image1(i).Height
        strStart(i, 4) = Image1(i).Top
        strStart(i, 5) = Image1(i).Left
     End If
   Next
End Sub

Private Sub Command4_Click()
 RichTextBox1.Text = "To Create a book" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  1) Click on the background to change the Background picture" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  2) Click on the Picture Name to change the item shown " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  3) Click on the actual picture to move it " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  4) Change the Picture width and height" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  5) Click on the PictureBar to add another picture" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  6) Once you have added all your pictures, then click Set Picture Start" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  7) You can now type in your story" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  8) Do not forget to add sounds and actions" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  9) Move a picture then pick MOVE then click ADD PICTURE ACTION " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & " 10) Click READ to hear your story" & Chr(13) & Chr(10)
End Sub

Private Sub ComPicture_Change()
'    Image1(TxtPicture.Text).Picture = LoadPicture(ComPicture.Text & ".WMF")
'    Image1m(TxtPicture.Text) = False
End Sub

Private Sub ComPicture_Click()
  Dim cPicBmp As String
   cPicBmp = soundpath & ComPicture.Text & ".WMF"
   Image1(TxtPicture.Text).Picture = LoadPicture(cPicBmp)
   Image1m(TxtPicture.Text) = False
   
   Image1(TxtPicture.Text).Width = 500 + (75 * HScroll2.Value)
   Image1(TxtPicture.Text).Height = 500 + (75 * HScroll3.Value)
    Image1(TxtPicture.Text).Visible = True
   strName(TxtPicture.Text) = ComPicture.Text
   Label1.Caption = "PictureBar " & HScroll1.Value + 1 & " " & strName(HScroll1.Value)
   If bFirst Then
     bFirst = False
     cPicBmp = "To move a picture just click on it"
     Myspeak (cPicBmp)
     bFirst = False
   End If
   If bStart = False Or strStart(TxtPicture.Text, 0) = "HIDE" Then
      'strStart(TxtPicture.Text, 0) = "SHOW"
      strStart(TxtPicture.Text, 2) = HScroll2.Value
      strStart(TxtPicture.Text, 3) = HScroll3.Value
      strStart(TxtPicture.Text, 4) = Image1(i).Top
      strStart(TxtPicture.Text, 5) = Image1(i).Left
   End If
End Sub

Private Sub ComPicture_KeyUp(KeyCode As Integer, Shift As Integer)
' Image1(TxtPicture.Text).Picture = LoadPicture(ComPicture.Text & ".WMF")
'    Image1m(TxtPicture.Text) = False
End Sub





Private Sub ComSound_Click()
 If Check1 Then MySound (ComSound.Text)
End Sub

Private Sub File1_Click()
 Dim cQuizfile, cPic, cBack As String
 Dim jMOVIE As Integer
  cQuizfile = File1.Path & "\" & File1.FileName
  jMOVIE = 0
  Dim i, iLen As Integer
  For i = 0 To 19
   strName(i) = ""
   strStart(i, 1) = "HIDE"
   Image1m(1) = False
  Next
  For i = 0 To 9
    Image1(i).Visible = False
  Next
  i = 0
  TxtStory.Text = ""
  TxtName.Text = Mid(File1.FileName, 1, Len(File1.FileName) - 8)
 Open cQuizfile For Input As #1 ' Open file for input.
  Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
    'Debug.Print MyString,  ' Print data to Debug window.
   ' if i=0 then ComBack.Text
'   For i = 0 To 9
  If i = 0 Then
    cBack = MyString
    GoTo NextRecord
  End If
    If i > 0 And i < 11 Then
       strStart(i - 1, 2) = Mid(MyString, 1, 5)
       strStart(i - 1, 3) = Mid(MyString, 6, 5)
       strStart(i - 1, 4) = Mid(MyString, 11, 5)
       strStart(i - 1, 5) = Mid(MyString, 16, 5)
      iLen = Len(MyString)
      cPic = ""
      If iLen > 20 Then
        cPic = soundpath & Mid(MyString, 21, iLen - 20) & ".WMF"
        'cPic = soundpath & ComPicture.Text & ".WMF"
       Image1(i - 1).Picture = LoadPicture(cPic)
       Image1(i - 1).Visible = True
       strName(i - 1) = Mid(MyString, 21, iLen - 20)
       strStart(i - 1, 1) = "SHOW"
       Else
       Image1(i - 1).Visible = False
      End If
      GoTo NextRecord
   End If
    If i = 11 Then
      iMovie = Mid(MyString, 1, 5) 'TxtStory.Text = MyString
      GoTo NextRecord
    End If
    If i > 11 And jMOVIE < iMovie Then
       jMOVIE = jMOVIE + 1
       strMovie(jMOVIE, 0) = MyString
       strMovie(jMOVIE, 1) = Mid(MyString, 1, 5)
       strMovie(jMOVIE, 2) = Mid(MyString, 6, 5)
       strMovie(jMOVIE, 3) = Mid(MyString, 11, 5)
       strMovie(jMOVIE, 4) = Mid(MyString, 16, 5)
       GoTo NextRecord
    Else
       TxtStory.Text = TxtStory.Text & MyString
    End If
    
NextRecord:
   i = i + 1
  Loop
  Close #1
  For i = 0 To 5
    If cBack = ComBack.List(i) Then ComBack.ListIndex = i
  Next
  Call ComGo_Click
  If bStart = False Then Command3_Click
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim j As Integer
  bFirst = False
  bStart = False
  Dir1.Path = App.Path
  iPic = 0
  iMovie = 0
  bApicture = False
  HScroll1.Min = 0     ' Set values of the scroll bar.
  HScroll1.Max = 9
  HScroll2.Min = 0    ' Set values of the scroll bar.
  HScroll2.Max = 25
  HScroll3.Min = 0    ' Set values of the scroll bar.
  HScroll3.Max = 25
  ComPictureAction.Clear
  ComPictureAction.AddItem "Move"
  ComPictureAction.AddItem "Hide"
  ComPictureAction.AddItem "Show"
  ComPictureAction.ListIndex = 0
  Label1.Caption = "Picture " & HScroll1.Value + 1
  TxtPicture.Text = HScroll1.Value
  Dim cMyLetter As String
  Dim MyString As String
  '
  ' Add pictures
  '
  ComPicture.Clear
  Call FindPicture
  For i = 0 To File1.ListCount - 1
  'File1.List (i)
  MyString = File1.List(i)
  j = Len(MyString)
    ComPicture.AddItem Left(MyString, j - 4)
  Next
  '
  ' Add Background
  '
  ComPicture.ListIndex = 0
  ComBack.Clear
  ComBack.AddItem "Stream"
  ComBack.AddItem "Field"
  ComBack.ListIndex = 0
  '
  ' Get all the sounds
  '
  Call FindSounds
  
  ComSound.Clear
  For i = 0 To File1.ListCount - 1
  'File1.List (i)
  MyString = File1.List(i)
  j = Len(MyString)
    ComSound.AddItem Left(MyString, j - 4)
  Next
  
  ComSound.ListIndex = 0
  '
  ' Add Character Action
  '
  ComAction.Clear
  ComAction.AddItem "Acknowledge"
  ComAction.AddItem "Announce"
  ComAction.AddItem "Confused"
  ComAction.AddItem "Congratulate"
  ComAction.AddItem "Decline"
  ComAction.AddItem "DoMagic2"
  ComAction.AddItem "DontRecognize"
  ComAction.AddItem "Explain"
  ComAction.AddItem "GestureDown"
  ComAction.AddItem "GestureLeft"
  ComAction.AddItem "GestureRight"
  ComAction.AddItem "GestureUp"
  ComAction.AddItem "GetAttention"
  ComAction.AddItem "Greet"
  ComAction.AddItem "MoveDown"
  ComAction.AddItem "MoveLeft"
  ComAction.AddItem "MoveRight"
  ComAction.AddItem "MoveUp"
  ComAction.AddItem "Pleased"
  ComAction.AddItem "Process"
  ComAction.AddItem "Read"
  ComAction.AddItem "Sad"
  ComAction.AddItem "Search"
  ComAction.AddItem "Show"
  ComAction.AddItem "Size 1"
  ComAction.AddItem "Size 2"
  ComAction.AddItem "Size 3"
  ComAction.AddItem "Suggest"
  ComAction.AddItem "Surprised"
  ComAction.AddItem "Uncertain"
  ComAction.AddItem "Wave"
  ComAction.AddItem "Write"
  ComAction.AddItem "Search"
  ComAction.ListIndex = 0
  'Dim i As Integer
  For i = 0 To 19
   strName(i) = ""
   Image1m(1) = False
  Next
  strName(0) = "Ball"
  Image1(0).Visible = True
  Dim cPicBmp As String
   cPicBmp = soundpath & ComPicture.Text & ".WMF"
   Image1(0).Picture = LoadPicture(cPicBmp)
    'Image1m(0) = False
  For i = 0 To 9
    Image1(i).Visible = False
    strStart(i, 0) = "HIDE"
    strStart(i, 2) = "0"
    strStart(i, 3) = "0"
    strStart(i, 4) = "0"
    strStart(i, 5) = "0"
  Next
  Image1(0).Visible = True
  strName(0) = ComPicture.Text
  MyAgent.Characters.Load cStoryAgent, cStoryAgent & ".acs"
                 'iCharCnt = iCharCnt + 1
                 'cCharacters(iCharCnt) = strInstructions
      Set agAgent = MyAgent.Characters(cStoryAgent)
                 Set Request = agAgent.Show
  Call FindBooks
  Command4_Click
  bStory = True
  bFirst = True
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  For i = 0 To 3
    If Image1m(i) Then
       Image1(i).Left = X
       Image1(i).Top = Y
    End If
  Next
End Sub

Private Sub BILL()
 If Image1m(0) Then
   Image1m(0) = False
   
   Label1.Caption = "F"
 Else
   Image1m(0) = True
   Label1.Caption = "T"
 End If
 'Label1.Caption = "TRUE"
End Sub



Private Sub HScroll1_Change()
  Label1.Caption = "PictureBar " & HScroll1.Value + 1 & " " & strName(HScroll1.Value)
  TxtPicture.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
 Image1(TxtPicture.Text).Width = 500 + (75 * HScroll2.Value)
 If bStart = False Then
      strStart(TxtPicture.Text, 2) = Image1(TxtPicture.Text).Width
   End If
End Sub

Private Sub HScroll3_Change()
 Image1(TxtPicture.Text).Height = 500 + (75 * HScroll3.Value)
 If bStart = False Then
       strStart(TxtPicture.Text, 3) = Image1(TxtPicture.Text).Height
   End If
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Label4.Caption = Index & " / " & strName(Index) & " / " & bApicture
If strName(Index) = "" Then Exit Sub
If Image1m(Index) Then
   Image1m(Index) = False
   Image1(Index).BorderStyle = 0
   bApicture = False
 Else
    If bApicture = False Then
        Image1m(Index) = True
        Image1(Index).BorderStyle = 1
        HScroll1.Value = Index
        i = Image1(TxtPicture.Text).Width - 500
        HScroll2.Value = Int(i / 75)
        i = Image1(TxtPicture.Text).Height - 500
        HScroll3.Value = Int(i / 75)
        bApicture = True
   End If
 End If
End Sub

Private Sub MyAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
 Set agAgent = Me.MyAgent.Characters(CharacterID)
 'CharPosn(0) = agAgent.Left
 'CharPosn(1) = agAgent.Top
End Sub

Private Sub MyAgent_RequestComplete(ByVal Request As Object)
 If Request Then
       bRequestDone = True
'      Merlin.Play "confused"
'      Merlin.Speak "Hey, Genie. What are you doing?"
'      Merlin.Interrupt GenieRequest
'
'      Genie.Speak "I was just checking on something."
 End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim i As Integer
  For i = 0 To 9
    If Image1m(i) Then
       Image1(i).Left = X
       Image1(i).Top = Y
    End If
  Next
End Sub

