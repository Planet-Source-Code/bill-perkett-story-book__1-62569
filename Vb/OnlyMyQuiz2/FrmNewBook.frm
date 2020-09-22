VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmNewBook 
   Caption         =   "Book"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10245
   LinkTopic       =   "Form4"
   ScaleHeight     =   9120
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtLabel 
      Height          =   405
      Left            =   240
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdLabel 
      Caption         =   "Add Label"
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
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   1680
      TabIndex        =   30
      Top             =   4320
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5106
      _Version        =   393217
      TextRTF         =   $"FrmNewBook.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   2775
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   4320
      Width           =   8415
   End
   Begin VB.CommandButton CmdErase 
      Caption         =   "Erase"
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
      TabIndex        =   28
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Picture"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Picture"
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton CmdHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   680
         Width           =   1095
      End
      Begin VB.CommandButton CmdMove 
         Caption         =   "Move"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBook 
      Caption         =   "Old Books"
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
      Left            =   120
      TabIndex        =   22
      Top             =   7200
      Width           =   1215
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
      Left            =   3960
      TabIndex        =   20
      Top             =   7680
      Width           =   1455
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
      Left            =   7440
      TabIndex        =   19
      Top             =   7680
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
      Left            =   5520
      TabIndex        =   18
      Top             =   7800
      Width           =   1815
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
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
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
      Left            =   120
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   4320
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   8880
      TabIndex        =   6
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
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
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox ComList 
      Height          =   6420
      Left            =   3720
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdPicture 
      Caption         =   "Add  Picture"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdBack 
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
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   4815
      Left            =   1680
      Picture         =   "FrmNewBook.frx":0082
      ScaleHeight     =   4755
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   5280
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   5280
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   5280
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3840
         TabIndex        =   37
         Top             =   2640
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   15
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   14
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   13
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   12
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   11
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   10
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   135
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
         Index           =   2
         Left            =   0
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
         Index           =   4
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   600
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
         Height          =   375
         Index           =   6
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
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
         Index           =   8
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   9
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Label Text"
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
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
      Left            =   5520
      TabIndex        =   21
      Top             =   7560
      Width           =   1575
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   9120
      Top             =   7200
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   8400
      Width           =   3975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Picture Width"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Picture Height"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label LblOption 
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "FrmNewBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iPicture As Integer
Dim iLabel As Integer
Dim iLabelChg As Integer
Dim strName(20) As String
Dim bApicture As Boolean
Dim strSaveBack As String
Dim Request As Object
Dim bFirst As Boolean
Private Sub ExecutePicture(strCommand As String)
  Dim i, j As Integer
If UCase((Mid(strCommand, 1, 4))) = "MOVE" Then
    j = Val(Mid(strCommand, 5, 3))
   For i = 0 To 14
         j = j + 1
'      iMovie = iMovie + 1
      If strName(i) <> "" And Len(strMovie(j, 0)) > 20 Then
        If Image1(i).Width <> strMovie(j, 1) Then Image1(i).Width = strMovie(j, 1)
        If Image1(i).Height <> strMovie(j, 1) Then Image1(i).Height = strMovie(j, 2)
        If Image1(i).Top <> strMovie(j, 1) Then Image1(i).Top = strMovie(j, 3)
        If Image1(i).Left <> strMovie(j, 1) Then Image1(i).Left = strMovie(j, 4)
     End If
   Next
   '
    For i = 0 To 7
      j = j + 1
      '
      '  iMovie = iMovie + 1
      '
      If Label1(i).Caption <> "" And Len(strMovie(j, 0)) > 20 Then
        Label1(i).ForeColor = Mid(strMovie(j, 0), 21, 20)
        Label1(i).Font.Size = Mid(strMovie(j, 0), 41, 5)
        Label1(i).Caption = Mid(strMovie(j, 0), 46, Len(strMovie(j, 0)) - 42)
        If Label1(i).Width <> strMovie(j, 1) Then Label1(i).Width = strMovie(j, 1)
        If Label1(i).Height <> strMovie(j, 1) Then Label1(i).Height = strMovie(j, 2)
        If Label1(i).Top <> strMovie(j, 1) Then Label1(i).Top = strMovie(j, 3)
        If Label1(i).Left <> strMovie(j, 1) Then Label1(i).Left = strMovie(j, 4)
     End If
   Next
  End If
  If UCase((Mid(strCommand, 1, 4))) = "HIDE" Then
    i = Val(Mid(strCommand, 5, 3))
     Image1(i).Visible = False
  End If
   If UCase((Mid(strCommand, 1, 4))) = "SHOW" Then
    i = Val(Mid(strCommand, 5, 3))
     Image1(i).Visible = True
  End If
End Sub
Private Sub Myspeak(cTxtin As String)
 Dim PauseTime, Start, Finish, TotalTime As Long
  Static result As IAgentCtlRequest
    Set agAgent = MyAgent.Characters(cStoryAgent)
   agAgent.StopAll
    Set Request = agAgent.Speak(cTxtin)
    bRequestDone = False
        PauseTime = 2   ' Set duration.
       For i = 1 To 3
       
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
          'Label2.Caption = "D " & cTxt
          Exit Sub
        End If
       Next
        'gAgent.StopAll
        'Label2.Caption = "N " & cTxt
    
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
Private Sub FindBooks()
  '
  ' Find Books
  '
    File1.Path = App.Path & "\"
    File1.Pattern = "*.txt"
    File1.Path = App.Path & "\"
    File1.Pattern = "*book.txt"
End Sub
Private Sub CmdAction_Click()
  Dim i As Integer
   Dim j As Integer
   Dim MyString  As String
   strButton = "ACTION"
   strChoice = ""
  ' FrmChoose.Show 1
   'If strChoice <> "" Then Picture2.Picture = LoadPicture(strChoice)
   ComList.Clear
   CmdLabel.Visible = False
   ComList.Visible = True
  ComList.AddItem "Acknowledge"
  ComList.AddItem "Announce"
  ComList.AddItem "Confused"
  ComList.AddItem "Congratulate"
  ComList.AddItem "Decline"
  ComList.AddItem "DoMagic2"
  ComList.AddItem "DontRecognize"
  ComList.AddItem "Explain"
  ComList.AddItem "GestureDown"
  ComList.AddItem "GestureLeft"
  ComList.AddItem "GestureRight"
  ComList.AddItem "GestureUp"
  ComList.AddItem "GetAttention"
  ComList.AddItem "Greet"
  ComList.AddItem "MoveDown"
  ComList.AddItem "MoveLeft"
  ComList.AddItem "MoveRight"
  ComList.AddItem "MoveUp"
  ComList.AddItem "Pleased"
  ComList.AddItem "Process"
  ComList.AddItem "Read"
  ComList.AddItem "Sad"
  ComList.AddItem "Search"
  ComList.AddItem "Show"
  ComList.AddItem "Size 1"
  ComList.AddItem "Size 2"
  ComList.AddItem "Size 3"
  ComList.AddItem "Suggest"
  ComList.AddItem "Surprised"
  ComList.AddItem "Uncertain"
  ComList.AddItem "Wave"
  ComList.AddItem "Write"
  ComList.AddItem "Search"
  ComList.Text = "Acknowledge"
   CmdPicture.Visible = False
   CmdBack.Visible = False
   LblOption.Caption = "Action"
   LblOption.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
   Label6.Visible = False
   CmdOk.Caption = "Add"
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdCancel.Visible = True
   CmdAction.Visible = False
   CmdBook.Visible = False
  cmdChange.Visible = False
   Frame1.Visible = False
End Sub

Private Sub cmdChange_Click()
   Dim i As Integer
   Dim j As Integer
   Dim MyString  As String
   strButton = "CHANGEPICTURE"
   strChoice = ""
  ' FrmChoose.Show 1
   'If strChoice <> "" Then Picture2.Picture = LoadPicture(strChoice)
   ComList.Clear
   ComList.Visible = True
   Call FindPicture
     For i = 0 To File1.ListCount - 1
           MyString = File1.List(i)
           j = Len(MyString)
            ComList.AddItem Left(MyString, j - 4)
      Next
      MyString = File1.List(0)
      j = Len(MyString)
      ComList.Text = Left(MyString, j - 4)
   CmdPicture.Visible = False
   CmdBack.Visible = False
   LblOption.Caption = "Picture"
   LblOption.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
   Label6.Visible = False
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdAction.Visible = False
   CmdBook.Visible = False
   cmdChange.Visible = False
   Frame1.Visible = False
End Sub

Private Sub CmdBack_Click()
   strButton = "BACKGROUND"
   strChoice = ""
  ' FrmChoose.Show 1
   'If strChoice <> "" Then Picture2.Picture = LoadPicture(strChoice)
   ComList.Clear
   ComList.Visible = True
   'ComList.AddItem "Desert"
   ComList.AddItem "Field"
   ComList.AddItem "Jungle"
  ' ComList.AddItem "Ocean"
   ComList.AddItem "Stream"
  ' ComList.AddItem "Rocks"
   ComList.Text = strSaveBack
   CmdPicture.Visible = False
   CmdBack.Visible = False
   LblOption.Caption = "Background"
   LblOption.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
   Label6.Visible = False
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdLabel.Visible = False
   CmdAction.Visible = False
   CmdBook.Visible = False
   RichTextBox1.Visible = False
   cmdChange.Visible = False
   Frame1.Visible = False
End Sub


Private Sub CmdBook_Click()
 Dim i As Integer
 Dim j As Integer
 Dim MyString As String
 Call FindBooks
 ComList.Clear
  For i = 0 To File1.ListCount - 1
      MyString = File1.List(i)
      j = Len(MyString)
       ComList.AddItem Left(MyString, j - 4)
   Next
      MyString = File1.List(0)
      j = Len(MyString)
      ComList.Text = Left(MyString, j - 4)
      CmdPicture.Visible = False
   CmdBack.Visible = False
   strButton = "BOOK"
   LblOption.Caption = "Book"
   LblOption.Visible = True
   ComList.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
    CmdLabel.Visible = False
   Label6.Visible = False
   CmdOk.Caption = "Load"
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdCancel.Visible = True
   CmdAction.Visible = False
   CmdBook.Visible = False
   cmdChange.Visible = False
   Frame1.Visible = False
End Sub

Private Sub CmdCancel_Click()
  Command1_Click
End Sub

Private Sub CmdErase_Click()
'  Dim i As Integer
'  iPicture = 0
'  iMovie = 15
'  Label4.Caption = ""
'  HScroll2.Min = 0    ' Set values of the scroll bar.
'  HScroll2.Max = 25
'  HScroll3.Min = 0    ' Set values of the scroll bar.
'  HScroll3.Max = 25
'  For i = 0 To 14
'   Image1(i).Visible = False
'   strMovie(i, 0) = "00000000000000000000"
'   strMovie(i, 1) = "00000000000000000000"
'   strMovie(i, 2) = "00000000000000000000"
'   strMovie(i, 3) = "00000000000000000000"
'   strMovie(i, 4) = "00000000000000000000"
'  Next
'  TxtStory.Text = ""
'  CmdPicture.Visible = True
   strButton = "ERASE"
'  Command1_Click
   CmdOk.Caption = "Ok to Erase"
   CmdOk.Visible = True
   CmdCancel.Visible = True
End Sub

Private Sub CmdHide_Click()
 TxtStory.Text = TxtStory.Text & " (HIDE" & Format(Label4.Caption, "000") & ")"
End Sub

Private Sub CmdLabel_Click()
   If iLabel < 8 Then
     iLabelChg = iLabel
    'Label1(iLabel).Visible = True
    strButton = "LABEL"
    Label2.Visible = True
    TxtLabel.Visible = True
    CmdSound.Visible = False
    CmdAction.Visible = False
    CmdBook.Visible = False
    CmdOk.Visible = True
    'CmdPicture.Visible = False
    CmdCancel.Visible = True
    cmdChange.Visible = False
   End If
End Sub

Private Sub CmdMove_Click()
    Dim i As Integer
    Dim cMsg As String
    TxtStory.Text = TxtStory.Text & " (MOVE" & Format(iMovie, "000") & ")"
    '
    ' Image
    '
     For i = 0 To 14
       iMovie = iMovie + 1
       strMovie(iMovie, 0) = "00000000000000000000" & strName(i)
       strMovie(iMovie, 1) = Image1(i).Width
       strMovie(iMovie, 2) = Image1(i).Height
       strMovie(iMovie, 3) = Image1(i).Top
       strMovie(iMovie, 4) = Image1(i).Left
    Next
    '
    '  Label
    '
    For i = 0 To 7
       iMovie = iMovie + 1
       cMsg = Label1(i).ForeColor
       cMsg = cMsg & Space(20 - Len(cMsg))
       strMovie(iMovie, 0) = "LABEL000000000000000" & cMsg & Format(Label1(i).Font.Size, "00.00") & Label1(i).Caption
       strMovie(iMovie, 1) = Label1(i).Width
       strMovie(iMovie, 2) = Label1(i).Height
       strMovie(iMovie, 3) = Label1(i).Top
       strMovie(iMovie, 4) = Label1(i).Left
    Next
End Sub


Private Sub CmdOk_Click()
 Dim cQuizfile, cPic, cBack As String
 Dim cMsg As String
 Dim jMOVIE As Integer
 Dim i, iLen As Integer
 Dim k As Integer
   If strButton = "LABEL" Then
    Label1(iLabel).Caption = TxtLabel.Text
    Label1(iLabel).Visible = True
    cMsg = Label1(iLabel).ForeColor
    cMsg = cMsg & Space(20 - Len(cMsg))
    'Format(Label1(iLabel).Font.Size, "00.00")
    strMovie(iLabel + 15, 0) = "LABEL000000000000000" & cMsg & Format(Label1(iLabel).Font.Size, "00.00") & Label1(iLabel).Caption
    strMovie(iLabel + 15, 1) = Label1(iLabel).Width
    strMovie(iLabel + 15, 2) = Label1(iLabel).Height
    strMovie(iLabel + 15, 3) = Label1(iLabel).Top
    strMovie(iLabel + 15, 4) = Label1(iLabel).Left
    iLabel = iLabel + 1
    strButton = ""
'    TxtLabel.Visible = False
'    Label2.Visible = False
  End If
  If strButton = "LABELCHANGE" Then
    Label1(iLabelChg).Caption = TxtLabel.Text
    strButton = ""
    TxtLabel.Visible = False
    Label2.Visible = False
  End If
   If strButton = "SOUND" Then
    TxtStory.Text = TxtStory.Text & " [" & ComList.Text & "]"
    strButton = ""
  End If
  If strButton = "ACTION" Then
    TxtStory.Text = TxtStory.Text & " {" & ComList.Text & "}"
    strButton = ""
  End If
  If strButton = "PICTURE" Then
    If bFirst Then
     bFirst = False
     cPicBmp = "To move a picture just click on it"
     Myspeak (cPicBmp)
     bFirst = False
   End If
  End If
  If strButton = "BOOK" Then
     cQuizfile = "c:/kids/" & ComList.Text & ".txt"
     jMOVIE = 0
     For i = 0 To 14
       strName(i) = ""
       Image1(i).Visible = False
      Next
      i = 0
      k = 0
      TxtStory.Text = ""
      TxtName.Text = ComList.Text
      strButton = ""
      Open cQuizfile For Input As #1 ' Open file for input.
      Line Input #1, MyString
      'Dir1.Path
      
      Picture2.Picture = LoadPicture(Dir1.Path & "\" & MyString & ".jpg")
     Do While Not EOF(1) ' Loop until end of file.
       Line Input #1, MyString
       Debug.Print MyString,  ' Print data to Debug window.
       If Mid(MyString, 1, 1) = "~" Or Mid(MyString, 1, 1) = "|" Then
          iLen = Len(MyString)
          strMovie(i, 0) = MyString
          If Mid(MyString, 1, 1) = "|" Then
            strMovie(i, 0) = "LABEL000000000000000"
            If iLen > 46 Then
              cMsg = Mid(MyString, 22, 20)
              cMsg = cMsg & Space(20 - Len(cMsg))
              strMovie(i, 0) = "LABEL000000000000000" & cMsg & Format(Mid(MyString, 42, 5), "00.00") & Mid(MyString, 47, Len(MyString) - 42)
           End If
          End If
          strMovie(i, 1) = Mid(MyString, 2, 5)
          strMovie(i, 2) = Mid(MyString, 7, 5)
          strMovie(i, 3) = Mid(MyString, 12, 5)
          strMovie(i, 4) = Mid(MyString, 17, 5)
          If i < 15 And iLen > 21 Then
             strName(i) = Mid(MyString, 22, iLen - 21)
             Image1(i) = LoadPicture(Dir1.Path & "\" & strName(i) & ".wmf")
             Image1(i).Width = strMovie(i, 1)
             Image1(i).Height = strMovie(i, 2)
             Image1(i).Top = strMovie(i, 3)
             Image1(i).Left = strMovie(i, 4)
             Image1(i).Visible = True
         ElseIf i < 23 And iLen > 21 Then
             'strName(i) = Mid(MyString, 22, iLen - 21)
             Label1(i - 15).ForeColor = Mid(MyString, 22, 20)
             Label1(i - 15).Font.Size = Mid(MyString, 42, 5)
             Label1(i - 15).Caption = Mid(MyString, 47, Len(MyString) - 42)
            ' Label1(i) = LoadPicture(Dir1.Path & "\" & strName(i) & ".wmf")
             Label1(i - 15).Width = strMovie(i, 1)
             Label1(i - 15).Height = strMovie(i, 2)
             Label1(i - 15).Top = strMovie(i, 3)
             Label1(i - 15).Left = strMovie(i, 4)
             Label1(i - 15).Visible = True
             cMsg = Label1(i - 15).ForeColor
             cMsg = cMsg & Space(20 - Len(cMsg))
             strMovie(i, 0) = "LABEL000000000000000" & cMsg & Format(Label1(i - 15).Font.Size, "00.00") & Label1(i - 15).Caption
         End If
          i = i + 1
      Else
         TxtStory.Text = TxtStory.Text & MyString
      End If
  Loop
  Close #1
  End If
  If strButton = "ERASE" Then
         iPicture = 0
        iMovie = 15
        Label4.Caption = ""
        HScroll2.Min = 0    ' Set values of the scroll bar.
        HScroll2.Max = 25
        HScroll3.Min = 0    ' Set values of the scroll bar.
        HScroll3.Max = 25
        For i = 0 To 14
         Image1(i).Visible = False
         strMovie(i, 0) = "00000000000000000000"
         strMovie(i, 1) = "00000000000000000000"
         strMovie(i, 2) = "00000000000000000000"
         strMovie(i, 3) = "00000000000000000000"
         strMovie(i, 4) = "00000000000000000000"
        Next
        TxtStory.Text = ""
        CmdPicture.Visible = True
        strButton = ""
      Command1_Click
  End If
  CmdOk.Caption = "OK"
  Command1_Click
End Sub

Private Sub CmdPicture_Click()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim L As Integer
   Dim MyString  As String
   strButton = "PICTURE"
   strChoice = ""
   HScroll2.Min = 0    ' Set values of the scroll bar.
   HScroll2.Max = 25
   HScroll3.Min = 0    ' Set values of the scroll bar.
   HScroll3.Max = 25
  ' FrmChoose.Show 1
   'If strChoice <> "" Then Picture2.Picture = LoadPicture(strChoice)
   ComList.Clear
   ComList.Visible = True
   Call FindPicture
     For i = 0 To File1.ListCount - 1
           MyString = File1.List(i)
           j = Len(MyString)
           k = InStr(1, UCase(MyString), "FLA")
           L = InStr(1, UCase(MyString), "MAP")
            If k + L = 0 Then ComList.AddItem Left(MyString, j - 4)
      Next
      MyString = File1.List(0)
      j = Len(MyString)
      ComList.Text = Left(MyString, j - 4)
   TxtLabel.Visible = False
   Label2.Visible = False
   CmdPicture.Visible = False
   CmdLabel.Visible = False
   CmdBack.Visible = False
   LblOption.Caption = "Picture"
   LblOption.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
   Label6.Visible = False
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdAction.Visible = False
   CmdBook.Visible = False
   RichTextBox1.Visible = False
   cmdChange.Visible = False
   Frame1.Visible = False
End Sub

Private Sub CmdRead_Click()
  'Label2.Caption = ""
  
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
    For i = 0 To 14
     If strName(i) <> "" Then
       Image1(i).Width = strMovie(i, 1)
       Image1(i).Height = strMovie(i, 2)
       Image1(i).Top = strMovie(i, 3)
       Image1(i).Left = strMovie(i, 4)
       Image1(i).Visible = True
     End If
    Next
    For i = 0 To 7
     If Label1(i).Caption <> "" Then
       Label1(i).ForeColor = Mid(strMovie(i + 15, 0), 21, 20)
       Label1(i).Font.Size = Mid(strMovie(i + 15, 0), 41, 5)
       Label1(i).Caption = Mid(strMovie(i + 15, 0), 46, Len(strMovie(i + 15, 0)) - 42)
       Label1(i).Width = strMovie(i + 15, 1)
       Label1(i).Height = strMovie(i + 15, 2)
       Label1(i).Top = strMovie(i + 15, 3)
       Label1(i).Left = strMovie(i + 15, 4)
       Label1(i).Visible = True
     End If
    Next
    '
    '
      cTxt = ""
     
   For i = 1 To Len(TxtStory.Text)
    cMY = Mid(TxtStory.Text, i, 1)
      If cMY = "[" Then
        If Trim(cTxt) <> "" Then Myspeak (cTxt)
        cTxt = ""
      ElseIf cMY = "]" Then
          MySound (cTxt)
          cTxt = ""
      ElseIf cMY = "{" Then
        If Trim(cTxt) <> "" Then Myspeak (cTxt)
        cTxt = ""
        ElseIf cMY = "}" Then
          ExecuteCommand (cTxt)
          cTxt = ""
        ElseIf cMY = "(" Then
        If Trim(cTxt) <> "" Then Myspeak (cTxt)
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
 If TxtName.Text = "" Then
   MsgBox "Enter a Book Name", vbOKOnly, "Book Name"
   TxtName.SetFocus
   Exit Sub
 End If
 Dim cQuizfile As String
 Dim cMsg As String
 Dim i, jHeight, jTop, jLeft, jWidth As Integer
 Dim k As Integer
 Dim j1, j2, j3, j4, jPic As Integer
 If UCase(Right(TxtName.Text, 4)) = "BOOK" Then
   cQuizfile = "c:/kids/" & TxtName.Text & ".txt"
 Else
   cQuizfile = "c:/kids/" & TxtName.Text & "book.txt"
 End If
   Open cQuizfile For Output As #1 ' Open file for output.
   Print #1, strSaveBack
     k = 0
   For i = 0 To iMovie
     If Mid(strMovie(i, 0), 1, 5) = "LABEL" Then
        j1 = Val(strMovie(i, 1)) 'Image1(i).Width
        j2 = Val(strMovie(i, 2)) ' Image1(i).Height
        j3 = Val(strMovie(i, 3)) 'Image1(i).Top
        j4 = Val(strMovie(i, 4)) 'Image1(i).Left
        cMsg = "|" & Format(j1, "00000") & Format(j2, "00000")
        cMsg = cMsg & Format(j3, "00000") & Format(j4, "00000") & Mid(strMovie(i, 0), 21, Len(strMovie(i, 0)) - 20)
        Print #1, cMsg
     Else
        j1 = Val(strMovie(i, 1)) 'Image1(i).Width
        j2 = Val(strMovie(i, 2)) ' Image1(i).Height
        j3 = Val(strMovie(i, 3)) 'Image1(i).Top
        j4 = Val(strMovie(i, 4)) 'Image1(i).Left
        cMsg = "~" & Format(j1, "00000") & Format(j2, "00000")
        cMsg = cMsg & Format(j3, "00000") & Format(j4, "00000") & strName(k)
        Print #1, cMsg
        k = k + 1
     If k > 14 Then k = 0
    End If
   Next
   Print #1, TxtStory.Text
   Close #1    ' Close file.
End Sub

Private Sub CmdShow_Click()
  TxtStory.Text = TxtStory.Text & " (SHOW" & Format(Label4.Caption, "000") & ")"
End Sub

Private Sub CmdSound_Click()
  Dim i As Integer
   Dim j As Integer
   Dim MyString  As String
   strButton = "SOUND"
   strChoice = ""
  ' FrmChoose.Show 1
   'If strChoice <> "" Then Picture2.Picture = LoadPicture(strChoice)
   ComList.Clear
   ComList.Visible = True
   Call FindSound
     For i = 0 To File1.ListCount - 1
           MyString = File1.List(i)
           j = Len(MyString)
           If j > 6 Then ComList.AddItem Left(MyString, j - 4)
      Next
      MyString = File1.List(0)
      j = Len(MyString)
      ComList.Text = Left(MyString, j - 4)
   CmdPicture.Visible = False
   CmdBack.Visible = False
   LblOption.Caption = "Sound"
   LblOption.Visible = True
   CmdOk.Visible = True
   Label5.Visible = False
   Label6.Visible = False
   CmdOk.Caption = "Add" '
    CmdLabel.Visible = False
   HScroll2.Visible = False
   HScroll3.Visible = False
   CmdSound.Visible = False
   CmdCancel.Visible = True
   CmdAction.Visible = False
   CmdBook.Visible = False
  cmdChange.Visible = False
   Frame1.Visible = False
End Sub

Private Sub ComList_Click()
  Dim i As Integer
  Select Case strButton
     Case "BACKGROUND"
       'Picture1.Picture LoadPicture(App.Path & "/" & "bg1.jpg")
      ' Picture1.Picture = LoadPicture(App.Path & "/" & ComList.Text & ".jpg")
        Picture2.Picture = LoadPicture(Dir1.Path & "\" & ComList.Text & ".jpg")
      ' strChoice = App.Path & "/" & ComList.Text & ".jpg"
        strSaveBack = ComList.Text
     Case "PICTURE"
       Image1(iPicture) = LoadPicture(Dir1.Path & "\" & ComList.Text & ".wmf")
       Image1(iPicture).Width = 700
       Image1(iPicture).Height = 700
       Image1(iPicture).Visible = True
       strChoice = App.Path & "\" & ComList.Text & ".wmf"
       strName(iPicture) = ComList.Text
       strMovie(iPicture, 0) = "00000000000000000000" & strName(iPicture)
       strMovie(iPicture, 1) = Image1(iPicture).Width
       strMovie(iPicture, 2) = Image1(iPicture).Height
       strMovie(iPicture, 3) = Image1(iPicture).Top
       strMovie(iPicture, 4) = Image1(iPicture).Left
       If TxtStory = "" Then
           For i = 0 To iPicture
             strMovie(i, 0) = "00000000000000000000" & strName(iPicture)
             strMovie(i, 1) = Image1(i).Width
             strMovie(i, 2) = Image1(i).Height
             strMovie(i, 3) = Image1(i).Top
             strMovie(i, 4) = Image1(i).Left
           Next
       End If
     Case "CHANGEPICTURE"
       Image1(Label4.Caption) = LoadPicture(Dir1.Path & "\" & ComList.Text & ".wmf")
       Image1(Label4.Caption).Width = 700
       Image1(Label4.Caption).Height = 700
       strChoice = Dir1.Path & "\" & ComList.Text & ".wmf"
       strName(Label4.Caption) = ComList.Text
       strMovie(iPicture, 0) = "00000000000000000000" & strName(iPicture)
     Case "SOUND"
      MySound (ComList.Text)
     Case "ACTION"
      'Dim cMsg As String
      '/cMsg = "/PLAY " & cStoryAgent & " " & ComAction.Text
      'TxtStory = TxtStory & " {" & ComAction.Text & "}"
        ExecuteCommand (ComList.Text)
   End Select

End Sub

Private Sub Command1_Click()
 If strButton = "PICTURE" Then
    If iPicture < 14 Then
      iPicture = iPicture + 1
    Else
      iPicture = iPicture + 1
      CmdPicture.Visible = False
    End If
    strButton = ""
 End If
  CmdOk.Visible = False
  ComList.Visible = False
  LblOption.Visible = False
  CmdPicture.Visible = True
  CmdBack.Visible = True
  CmdLabel.Visible = True
  CmdSound.Visible = True
  CmdCancel.Visible = False
  CmdAction.Visible = True
  CmdBook.Visible = True
  TxtLabel.Visible = False
  Label2.Visible = False
  RichTextBox1.Visible = False
  If iPicture = 15 Then CmdPicture.Visible = False
End Sub

Private Sub Form_Load()
  ComList.Left = 20
  bFirst = True
  iPicture = 0
  iLabel = 0
  iMovie = 15 + 8
  strSaveBack = "Stream"
  Label4.Caption = ""
  HScroll2.Min = 0    ' Set values of the scroll bar.
  HScroll2.Max = 25
  HScroll3.Min = 0    ' Set values of the scroll bar.
  HScroll3.Max = 25
  Dir1.Path = App.Path
  MyAgent.Characters.Load cStoryAgent, cStoryAgent & ".acs"
  Set agAgent = MyAgent.Characters(cStoryAgent)
         Set Request = agAgent.Show
         '
 RichTextBox1.Text = "To Create a book" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  1) Click on the background to change the Background picture" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  2) Click on Add Picture to add a picture" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  3) Click on the actual picture to move it " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  4) Change the Picture width and height" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  5) Click on Add Picture to add another picture" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  6) Once you have added all your pictures, then " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  7) You can type in your story" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  8) Do not forget to add sounds and actions" & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & "  9) Move a picture then click MOVE to ADD PICTURE ACTION " & Chr(13) & Chr(10)
 RichTextBox1.Text = RichTextBox1.Text & " 10) Click READ to hear your story" & Chr(13) & Chr(10)
 For i = 0 To 15
   Image1(i).Visible = False
   strMovie(i, 0) = "00000000000000000000"
   strMovie(i, 1) = "00000000000000000000"
   strMovie(i, 2) = "00000000000000000000"
   strMovie(i, 3) = "00000000000000000000"
   strMovie(i, 4) = "00000000000000000000"
  Next
  For i = 0 To 15
   Image1(i).Visible = False
   strMovie(i + 15, 0) = "LABEL000000000000000"
   strMovie(i + 15, 1) = "00000000000000000000"
   strMovie(i + 15, 2) = "00000000000000000000"
   strMovie(i + 15, 3) = "00000000000000000000"
   strMovie(i + 15, 4) = "00000000000000000000"
  Next
End Sub
Private Sub HScroll2_Change()
 Dim i As Integer
 Dim cMsg As String
 If HScroll2.Max = 15 Then
  Select Case HScroll2.Value
   Case 0
    Label1(iLabelChg).ForeColor = vbWhite
   Case 1
    Label1(iLabelChg).ForeColor = &HFF '&H8080FF
   Case 2
    Label1(iLabelChg).ForeColor = &HC0&
   Case 3
    Label1(iLabelChg).ForeColor = &H80FF&
  Case 4
    Label1(iLabelChg).ForeColor = &H40C0&
   Case 5
    Label1(iLabelChg).ForeColor = &HFFFF&       '&H80C0FF
   Case 6
    Label1(iLabelChg).ForeColor = &HC0C0&
   Case 7
    Label1(iLabelChg).ForeColor = &HFF00&
   Case 8
    Label1(iLabelChg).ForeColor = &H8000&
   Case 9
    Label1(iLabelChg).ForeColor = &HFFFF00
   Case 10
    Label1(iLabelChg).ForeColor = &H808000
   Case 11
    Label1(iLabelChg).ForeColor = &HFF0000
   Case 12
     Label1(iLabelChg).ForeColor = &H800000
   Case 13
    Label1(iLabelChg).ForeColor = &HFF00FF
   Case 14
    Label1(iLabelChg).ForeColor = &HC000C0
   Case 15
    Label1(iLabelChg).ForeColor = &H0&
   End Select
   If TxtStory = "" Then
      cMsg = Label1(iLabelChg).ForeColor
      cMsg = cMsg & Space(20 - Len(cMsg))
      'Format(Label1(Index).Font.Size, "00.00")
      '                         "00000000000000000000"
      strMovie(iLabelChg + 15, 0) = "LABEL000000000000000" & cMsg & Format(Label1(iLabelChg).Font.Size, "00.00") & Label1(iLabelChg).Caption
      strMovie(iLabelChg + 15, 1) = Label1(iLabelChg).Width
     strMovie(iLabelChg + 15, 2) = Label1(iLabelChg).Height
     strMovie(iLabelChg + 15, 3) = Label1(iLabelChg).Top
     strMovie(iLabelChg + 15, 4) = Label1(iLabelChg).Left
   
    End If
 Else
 Image1(Label4.Caption).Width = 500 + (75 * HScroll2.Value)
 If TxtStory = "" Then
           For i = 0 To iPicture
             strMovie(i, 0) = "00000000000000000000" & strName(iPicture)
             strMovie(i, 1) = Image1(i).Width
             strMovie(i, 2) = Image1(i).Height
             strMovie(i, 3) = Image1(i).Top
             strMovie(i, 4) = Image1(i).Left
           Next
       End If
 End If
End Sub

Private Sub HScroll3_Change()
 Dim i As Integer
 Dim cMsg As String
  If HScroll3.Max = 6 Then
  Select Case HScroll3.Value
   Case 0
    Label1(iLabelChg).Font.Size = 12
   Case 1
    Label1(iLabelChg).Font.Size = 14
   Case 2
    Label1(iLabelChg).Font.Size = 18
   Case 3
    Label1(iLabelChg).Font.Size = 24
   Case 4
    Label1(iLabelChg).Font.Size = 28
   Case 5
    Label1(iLabelChg).Font.Size = 36
   Case 6
    Label1(iLabelChg).Font.Size = 48
  End Select
  If TxtStory = "" Then
      cMsg = Label1(iLabelChg).ForeColor
      cMsg = cMsg & Space(20 - Len(cMsg))
      'Format(Label1(Index).Font.Size, "00.00")
      '                         "00000000000000000000"
      strMovie(iLabelChg + 15, 0) = "LABEL000000000000000" & cMsg & Format(Label1(iLabelChg).Font.Size, "00.00") & Label1(iLabelChg).Caption
      strMovie(iLabelChg + 15, 1) = Label1(iLabelChg).Width
     strMovie(iLabelChg + 15, 2) = Label1(iLabelChg).Height
     strMovie(iLabelChg + 15, 3) = Label1(iLabelChg).Top
     strMovie(iLabelChg + 15, 4) = Label1(iLabelChg).Left
   
    End If
 Else
 Image1(Label4.Caption).Height = 500 + (75 * HScroll3.Value)
    If TxtStory = "" Then
              For i = 0 To iPicture
                strMovie(i, 0) = "00000000000000000000" & strName(iPicture)
                strMovie(i, 1) = Image1(i).Width
                strMovie(i, 2) = Image1(i).Height
                strMovie(i, 3) = Image1(i).Top
                strMovie(i, 4) = Image1(i).Left
              Next
    End If
 End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer
'If TxtStory = "" Then
'     For i = 0 To iPicture
'       strMovie(i, 0) = "00000000000000000000" & strName(iPicture)
'       strMovie(i, 1) = Image1(i).Width
'       strMovie(i, 2) = Image1(i).Height
'       strMovie(i, 3) = Image1(i).Top
'       strMovie(i, 4) = Image1(i).Left
'    Next
' End If
End Sub

Private Sub Label1_Click(Index As Integer)
  Dim cMsg As String
  If Label1(Index).BorderStyle = 0 Then
    Label1(Index).BorderStyle = 1
'    CmdSound.Visible = False
'    CmdAction.Visible = False
'    CmdBook.Visible = False
'    Label2.Visible = True
'    TxtLabel.Visible = True
    TxtLabel.Text = Label1(Index).Caption
    cmdChange.Visible = False
  Else
   Label5.Visible = True
   Label6.Visible = True
   HScroll2.Visible = True
   HScroll3.Visible = True
   Frame1.Visible = True
   Label1(Index).BorderStyle = 0
   iLabelChg = Index
'   CmdSound.Visible = False
'    CmdAction.Visible = False
'    CmdBook.Visible = False
    'CmdPicture.Visible = False
    Frame1.Caption = "Add Label"
    Label2.Visible = True
    TxtLabel.Visible = True
    TxtLabel.Text = Label1(Index).Caption
    CmdOk.Visible = True
    CmdCancel.Visible = True
     strButton = "LABELCHANGE"
    Label5.Caption = TxtLabel.Text & " Color"
    Label6.Caption = TxtLabel.Text & " Size"
    HScroll2.Min = 0    ' Set values of the scroll bar.
    HScroll2.Max = 15
     HScroll3.Min = 0    ' Set values of the scroll bar.
     HScroll3.Max = 6
    If TxtStory = "" Then
      cMsg = Label1(Index).ForeColor
      cMsg = cMsg & Space(20 - Len(cMsg))
      'Format(Label1(Index).Font.Size, "00.00")
      '                         "00000000000000000000"
      strMovie(Index + 15, 0) = "LABEL000000000000000" & cMsg & Format(Label1(Index).Font.Size, "00.00") & Label1(Index).Caption
      strMovie(Index + 15, 1) = Label1(Index).Width
     strMovie(Index + 15, 2) = Label1(Index).Height
     strMovie(Index + 15, 3) = Label1(Index).Top
     strMovie(Index + 15, 4) = Label1(Index).Left
   
   End If
  End If
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
  For i = 0 To 15
    If Image1(i).BorderStyle = 1 Then
       Image1(i).Left = X
       Image1(i).Top = Y
    End If
  Next
  
  For i = 0 To 5
    If Label1(i).BorderStyle = 1 Then
       Label1(i).Left = X
       Label1(i).Top = Y
    End If
  Next
End Sub
Private Sub MySound(cFile As String)
Dim PauseTime, Start, Finish, TotalTime As Long
  'Label2.Caption = ""
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
   For i = 1 To 2
      Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer
     j = MMControl1.Position
     If j = JOLD Then
        'Label2.Caption = "DS"
        Exit Sub
     End If
     JOLD = j
   Next
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
Private Sub FindSound()
  '
  ' Find Picture
  '
    File1.Path = Dir1.Path
    File1.Pattern = "*.txt"
    File1.Path = Dir1.Path
    File1.Pattern = "*.wav"
End Sub
Private Sub Image1_Click(Index As Integer)
 Dim i As Integer
 Label4.Caption = Index
 If strName(Index) = "" Then Exit Sub
 If Image1(Index).BorderStyle = 1 Then
    Image1(Index).BorderStyle = 0
     bApicture = False
 Else
    If bApicture = False Then
        Image1(Index).BorderStyle = 1
        'HScroll1.Value = Index
        i = Image1(Index).Width - 500
        HScroll2.Value = Int(i / 75)
        i = Image1(Index).Height - 500
        HScroll3.Value = Int(i / 75)
        bApicture = True
        Command1_Click
   End If
 End If
 Label5.Caption = strName(Index) & " Width"
 Label6.Caption = strName(Index) & " Height"
 Frame1.Caption = "Add Picture"
 Label5.Visible = True
 Label6.Visible = True
 HScroll2.Visible = True
 HScroll3.Visible = True
 Frame1.Visible = True
 cmdChange.Visible = True
 '
 If TxtStory = "" Then
     For i = 0 To iPicture
      strMovie(i, 0) = "00000000000000000000" & strName(iPicture)
      strMovie(i, 1) = Image1(i).Width
      strMovie(i, 2) = Image1(i).Height
      strMovie(i, 3) = Image1(i).Top
      strMovie(i, 4) = Image1(i).Left
    Next
 End If
End Sub

Private Sub TxtLabel_KeyDown(KeyCode As Integer, Shift As Integer)
  'Label1(iLabelChg).Caption = TxtLabel.Text
End Sub

Private Sub TxtLabel_KeyUp(KeyCode As Integer, Shift As Integer)
  Label1(iLabelChg).Caption = TxtLabel.Text
  Label5.Caption = TxtLabel.Text & " Color"
End Sub
