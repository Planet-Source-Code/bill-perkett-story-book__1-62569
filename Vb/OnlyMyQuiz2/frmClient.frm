VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmClient 
   Caption         =   "Create a Play"
   ClientHeight    =   9195
   ClientLeft      =   270
   ClientTop       =   1185
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   11385
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
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
      Left            =   1560
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame PositionFrame 
      Caption         =   "Position"
      Enabled         =   0   'False
      Height          =   600
      Left            =   4680
      TabIndex        =   36
      Top             =   0
      Width           =   1890
      Begin VB.TextBox CharPosn 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   38
         Top             =   240
         Width           =   570
      End
      Begin VB.TextBox CharPosn 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   37
         Top             =   240
         Width           =   570
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&Y:"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1000
         TabIndex        =   40
         Top             =   240
         Width           =   270
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&X:"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9960
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   7800
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4560
      TabIndex        =   28
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
      Height          =   1650
      Left            =   6240
      TabIndex        =   27
      Top             =   7080
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frmClient.frx":0000
      Left            =   3360
      List            =   "frmClient.frx":0002
      TabIndex        =   26
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdThink 
      Caption         =   "Think"
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
      Left            =   10680
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.FileListBox lstAgents 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Insert"
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
      Left            =   10320
      TabIndex        =   13
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtSay 
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
      Left            =   7320
      TabIndex        =   12
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdSay 
      Caption         =   "&Say"
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
      Left            =   10800
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Play Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      Begin VB.CommandButton CmdRecord 
         Caption         =   "Record"
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
         Left            =   3960
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExecute 
         Caption         =   "Execute"
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
         Left            =   5640
         TabIndex        =   33
         Top             =   5520
         Width           =   855
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
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5640
         TabIndex        =   25
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdMoveScriptUp 
         Caption         =   "Up"
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
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdMoveScriptDown 
         Caption         =   "Down"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdStep 
         Caption         =   "Step"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdAuto 
         Caption         =   "Play"
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
         Left            =   5880
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.ListBox lstScript 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   5640
         TabIndex        =   9
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox txtCommand 
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
         Left            =   360
         TabIndex        =   8
         Top             =   5280
         Width           =   5175
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   5880
         TabIndex        =   7
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label LblStatus 
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
         TabIndex        =   52
         Top             =   5280
         Width           =   135
      End
      Begin VB.Label LblRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "     OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   4920
         TabIndex        =   48
         Top             =   120
         Width           =   600
      End
      Begin VB.Label LblSave 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PlayName"
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
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Step Command Line:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   5040
         Width           =   2415
      End
   End
   Begin VB.ListBox lstActive 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   0
      TabIndex        =   4
      Top             =   4920
      Width           =   4455
   End
   Begin VB.ListBox lstGestures 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
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
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
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
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Command Status      D-Done       N-Not Done E-Error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   3360
      TabIndex        =   51
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label LblGreen 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label12"
      Height          =   255
      Left            =   10680
      TabIndex        =   50
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblRed 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label11"
      Height          =   375
      Left            =   9840
      TabIndex        =   49
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Say/ Think"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   46
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "1"
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
      Left            =   90
      TabIndex        =   45
      Top             =   3120
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "2"
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
      Left            =   90
      TabIndex        =   44
      Top             =   3720
      Width           =   120
   End
   Begin VB.Label LblChr2 
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
      Left            =   240
      TabIndex        =   43
      Top             =   3720
      Width           =   75
   End
   Begin VB.Label LblChr1 
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
      Left            =   240
      TabIndex        =   42
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label LblAgent 
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
      Left            =   2400
      TabIndex        =   34
      Top             =   240
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Play Files"
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
      Left            =   6600
      TabIndex        =   29
      Top             =   6840
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Script Commands"
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
      TabIndex        =   24
      Top             =   4680
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Characters"
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
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   930
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   5040
      Top             =   7560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Action Commands"
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
      Left            =   1590
      TabIndex        =   0
      Top             =   600
      Width           =   1545
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim strUserID As String
Dim strUserName As String
Dim strAgentList() As String
Dim strCurrentAgentID As String
Const MYVERSION = 1#
Dim blRefreshing As Boolean
Private agAgent As IAgentCtlCharacterEx
Public OK As Boolean
Dim GenieRequest As Object
Dim MerlinRequest As Object
Dim Request As Object
Dim Genie As Object
Dim Merlin As Object

'Private Sub cboActiveServers_Click()
'    Me.txtHost.Text = Me.cboActiveServers.Text
'    tcpClient(1).Close
'    If tcpClient(1).State <> sckConnected Then 'It's connected
'        tcpClient(1).RemotePort = Val(txtPort.Text)
'    End If
'End Sub

Private Sub cmdAdd_Click()
    Dim intItem As Integer
    
    If Me.txtCommand.Text = "" Then Exit Sub
    intItem = Me.lstScript.ListIndex
    Me.lstScript.AddItem Me.txtCommand.Text, IIf(intItem < 0, 0, intItem)
    Me.lstScript.ListIndex = intItem
End Sub

Private Sub cmdAuto_Click()
    Dim intLine As Integer
    List1.Clear
    LblStatus.Caption = ""
    For intLine = 0 To Me.lstScript.ListCount - 1
        Me.lstScript.ListIndex = intLine
'        If UCase(Left(Me.lstScript.List(intLine), 6)) = "/PAUSE" Then
'          Sleep 5000
'         Else
           ExecuteCommand txtCommand.Text
       ' End If
'        If tcpClient(1).State = sckConnected Then
'            tcpClient(1).SendData Me.lstScript.List(intLine) & vbCrLf
'        End If
        'DoEvents
    Next intLine
End Sub


Private Sub cmdDelete_Click()
    Dim intItem As Integer
    
    intItem = Me.lstScript.ListIndex
    If intItem < 0 Then Exit Sub
    If intItem < Me.lstScript.ListCount - 2 Then
        Me.lstScript.ListIndex = intItem + 1
    Else
        If intItem > 0 Then Me.lstScript.ListIndex = intItem - 1
    End If
    Me.lstScript.RemoveItem (intItem)
End Sub

Private Sub CmdExecute_Click()
'LblStatus.Caption = ""
  ExecuteCommand txtCommand.Text
End Sub

Private Sub cmdHide_Click()
      If Me.lstAgents.FileName <> "" Then
        LblStatus.Caption = ""
        ExecuteCommand "/HIDE " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        txtCommand.Text = "/HIDE " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        If bRecord Then Command3_Click
    End If
 End Sub


Private Sub cmdMoveScriptDown_Click()
    Dim intNew As Integer
    If lstScript.ListIndex = Me.lstScript.ListCount - 1 Then Exit Sub
    intNew = Me.lstScript.ListIndex
    
    Me.lstScript.AddItem Me.lstScript.Text, intNew + 2
    Me.lstScript.RemoveItem (intNew)
    Me.lstScript.ListIndex = intNew + 1
End Sub

Private Sub cmdMoveScriptUp_Click()
    Dim intNew As Integer
    If lstScript.ListIndex = 0 Then Exit Sub
    intNew = Me.lstScript.ListIndex
    
    Me.lstScript.AddItem Me.lstScript.Text, intNew - 1
    Me.lstScript.RemoveItem (intNew + 1)
    Me.lstScript.ListIndex = intNew - 1
End Sub

Private Sub cmdPause_Click()
    Dim intItem As Integer
    
    intItem = Me.lstScript.ListIndex
    Me.lstScript.AddItem "/PAUSE", IIf(intItem < 0, 0, intItem)
    Me.lstScript.ListIndex = intItem
    
End Sub

Private Sub CmdRecord_Click()
  If bRecord Then
      bRecord = False
    LblRecord.Caption = "      OFF"
    LblRecord.BackColor = LblRed.BackColor
  Else
    bRecord = True
    LblRecord.Caption = "      ON"
    LblRecord.BackColor = LblGreen.BackColor
  End If
End Sub

Private Sub CmdSave_Click()
   Dim cFile As String '
' Update quiz text file
   'If TxtFile.Text = "" Then Exit Sub
   If TxtName.Text = "" Then Exit Sub
   File1.Path = Dir1.Path
   File1.Pattern = "*.txt"
   'If Len(cQuizfile) > 0 Then
   'Dim MyString As String
   cFile = "c:/kids/" & TxtName.Text & "msa.txt"
   Open cFile For Output As #1 ' Open file for output.
   Dim intLine As Integer
    For intLine = 0 To Me.lstScript.ListCount - 1
        Me.lstScript.ListIndex = intLine
'        If UCase(Left(Me.lstScript.List(intLine), 6)) = "/PAUSE" Then
'          Sleep 5000
'         Else
            Print #1, txtCommand.Text
       ' End If
'        If tcpClient(1).State = sckConnected Then
'            tcpClient(1).SendData Me.lstScript.List(intLine) & vbCrLf
'        End If
        'DoEvents
    Next intLine
  ' Print #1, TxtFile.Text
   Close #1    ' Close file.
  'End If
   File1.Path = Dir1.Path
   File1.Pattern = "*msa.txt"
   LblSave.Caption = TxtName.Text & " Saved."
   LblSave.Visible = True
    
End Sub

Private Sub cmdSay_Click()
   ' ExecuteCommand "/SAY " & Me.lstActive & " " & Me.txtSay
    If Me.lstAgents.FileName <> "" Then
        ExecuteCommand "/SAY " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.txtSay
    End If
End Sub

Private Sub cmdShow_Click()
        If Me.lstAgents.FileName <> "" Then
        LblStatus.Caption = ""
        LblAgent.Caption = Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        LblChr2.Caption = LblChr1.Caption
        LblChr1.Caption = LblAgent.Caption
        ExecuteCommand "/SHOW " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        txtCommand.Text = "/SHOW " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        If bRecord Then Command3_Click
        'CharPosn(0) = agAgent.Left
        'CharPosn(1) = agAgent.Top
    End If
  ' AnimationFrame.Caption = "&Animations for " + Character.Name

'-- Disable the Play button to avoid trying to play a null animation selection


'-- Load the character's animation into the list box
    lstGestures.Clear
    For Each AnimationName In agAgent.AnimationNames
            lstGestures.AddItem AnimationName
    Next
    LblAgent.Caption = agAgent.Name
End Sub



Private Sub cmdStep_Click()
    Dim intLine As Integer
    intLine = Me.lstScript.ListIndex
    If intLine < Me.lstScript.ListCount - 1 Then
        Me.lstScript.ListIndex = intLine + 1
    End If
    If UCase(Left(Me.lstScript.List(intLine), 6)) = "/PAUSE" Then Exit Sub
      ExecuteCommand txtCommand.Text
End Sub

Private Sub CmdStop_Click()
         ExecuteCommand "/STOP " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        txtCommand.Text = "/STOP " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4)
        If bRecord Then Command3_Click
End Sub

Private Sub cmdThink_Click()
  If Me.lstAgents.FileName <> "" Then
        ExecuteCommand "/THINK " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.txtSay
    End If
End Sub

Private Sub cmdUpdate_Click()
    If Me.txtCommand.Text <> "" Then
        If Me.lstScript.ListIndex > 0 Then
            Me.lstScript.List(Me.lstScript.ListIndex) = Me.txtCommand.Text
        End If
    End If
End Sub


Private Sub Command1_Click()
    Dim cMsg As String
    lstActive.Clear
    cMsg = "/HIDE MERLIN      - Hide Merlin"
    lstActive.AddItem cMsg
    cMsg = "/MOVE Merlin 300,400   - Move Merlin to 300,400"
    lstActive.AddItem cMsg
    cMsg = "/MOVEDOWN Merlin Genie  - Move Merlin below Genie"
    lstActive.AddItem cMsg
    cMsg = "/MOVERIGHT Merlin Genie  - Move Merlin to right of Genie"
    lstActive.AddItem cMsg
    cMsg = "/MOVELEFT Merlin Genie  - Move Merlin to left Genie"
    lstActive.AddItem cMsg
    cMsg = "/MOVEUP Merlin Genie  - Move Merlin above Genie"
    lstActive.AddItem cMsg
    cMsg = "/PLAY MERLIN sad  - Merlin will play sad command"
    lstActive.AddItem cMsg
    'cMsg = "/POINT Merlin Genie  - Merlin will point to Genie"
   ' lstActive.AddItem cMsg
    cMsg = "/PAUSE            - Pause a short time"
    lstActive.AddItem cMsg
    cMsg = "/SAY MERLIN HI    - Merlin will say HI"
    lstActive.AddItem cMsg
    cMsg = "/SIZE MERLIN 1    - Merlin will be small"
    lstActive.AddItem cMsg
    cMsg = "/SIZE MERLIN 2    - Merlin will be normal"
    lstActive.AddItem cMsg
    cMsg = "/SIZE MERLIN 3    - Merlin will be big"
    lstActive.AddItem cMsg
    cMsg = "/SHOW MERLIN      - To show Merlin"
    lstActive.AddItem cMsg
    cMsg = "/SOUND MERLIN ON   - To turn sound effects on"
    lstActive.AddItem cMsg
    cMsg = "/SOUND MERLIN OFF  - To turn sound effects off"
    lstActive.AddItem cMsg
    cMsg = "/STOP MERLIN      - To stop Merlin action"
    lstActive.AddItem cMsg
    cMsg = "/THINK MERLIN IDEA  - Merlin think IDEA"
    lstActive.AddItem cMsg
    cMsg = "/UNLOAD MERLIN       - Unload Merlin"
    lstActive.AddItem cMsg
End Sub


Private Sub Command3_Click()
   Me.lstScript.AddItem txtCommand.Text
End Sub

Private Sub File1_Click()
 ' MyAgent.StopAll
 Dim strFilePath As String
 Dim cMsg As String
 Dim i As Integer
' If bHide Then
'    MyAgent.Show
'     bHide = False
' End If
LblSave.Visible = False
 strFilePath = File1.Path & "\" & File1.FileName
   Me.lstScript.Clear
    Open strFilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, strLine
        Me.lstScript.AddItem strLine
    Loop
    Close 1
  i = Len(File1.FileName)
  TxtName.Text = Mid(File1.FileName, 1, i - 7)
End Sub

Private Sub Form_Load()
Dim sDirName As String
  iCharCnt = 0
  CentreMe Me
     Me.Show
' The name of the Winsock control is tcpClient(1).
' Note: to specify a remote host, you can use
' either the IP address (ex: "121.111.1.1") or
' the computer's "friendly" name, as shown here.
    bRecord = False
    Me.lstActive.Clear
    'Me.lstAgents.Clear
    Me.lstGestures.Clear
    sDirName = GetWindowsDir() & "Msagent\Chars\"
    lstAgents.Path = sDirName
     Dir1.Path = App.Path
     File1.Path = Dir1.Path
     File1.Pattern = "*msa.txt"
    lstAgents.Pattern = "*.acs"
    Command1_Click
  '    Load tcpClient(2)
'    tcpClient(1).RemoteHost = tcpClient(0).LocalIP
'    Me.txtHost = tcpClient(0).LocalIP
'    tcpClient(1).RemotePort = 2000
'    tcpClient(1).LocalPort = 0
'    SetControls False
'    statbar.Panels(1).Text = "Status " & GetStatus()
'    Set frmSlave = New frmServer
'    Load frmSlave
End Sub



Private Sub Form_Unload(Cancel As Integer)
'    Dim intCount As Integer
'    For intCount = 0 To tcpClient.Count - 1
'        If tcpClient(intCount).State <> sckClosed Then
'            tcpClient(intCount).Close
'        End If
'    Next intCount
'    Unload frmSlave
End Sub

Private Sub LblRecord_Click()
If bRecord Then
      bRecord = False
    LblRecord.Caption = "      OFF"
    LblRecord.BackColor = LblRed.BackColor
  Else
    bRecord = True
    LblRecord.Caption = "      ON"
    LblRecord.BackColor = LblGreen.BackColor
  End If
End Sub

Private Sub List1_Click()
Dim cMsg As String
Dim i As Integer
  If Me.List1.Text <> "" Then
        i = Len(List1.Text)
        txtCommand.Text = Mid(List1.Text, 3, i - 2)
        LblStatus.Caption = Left(List1.Text, 1)
        If bRecord Then Command3_Click
    End If
End Sub

Private Sub lstActive_Click()
    Dim cMsg As String
    Dim cMsg2 As String
    Dim i As Integer
    Dim bSay As Boolean
    LblStatus.Caption = ""
'    If Me.lstActive.DataChanged Then
'        strCurrentAgentID = Me.lstActive
'        ExecuteCommand "/LISTGESTURES " & Me.lstActive
'    End If
    cMsg = Me.lstActive.Text
    cMsg2 = Me.lstActive.Text
     i = InStr(1, cMsg, " ")
    cMsg = Mid(cMsg, 1, i - 1)
    bSay = False
    Select Case cMsg
     Case "/SIZE"
        cMsg = cMsg & " " & LblAgent.Caption & " " & Mid(cMsg2, 14, 1)
    Case "/SOUND"
        cMsg = cMsg & " " & LblAgent.Caption & " " & Trim(Mid(cMsg2, 15, 3))
     Case "/HIDE"
        cMsg = cMsg & " " & LblAgent.Caption
     Case "/UNLOAD"
        cMsg = cMsg & " " & LblAgent.Caption
     Case "/SAY"
        bSay = True
        cMsg = cMsg & " " & LblAgent.Caption & " " & txtSay.Text
     Case "/THINK"
        bSay = True
        cMsg = cMsg & " " & LblAgent.Caption & " " & txtSay.Text
     Case "/STOP"
        cMsg = cMsg & " " & LblAgent.Caption
     Case "/SHOW"
        cMsg = cMsg & " " & LblAgent.Caption
     Case "/MOVE"
       cMsg = cMsg & " " & LblAgent.Caption & " " & CharPosn(0).Text & "," & CharPosn(1).Text '= agAgent.Top
    Case "/MOVEDOWN"
      cMsg = cMsg & " " & LblAgent.Caption & " " & LblChr2.Caption
    Case "/MOVERIGHT"
      cMsg = cMsg & " " & LblAgent.Caption & " " & LblChr2.Caption
    Case "/MOVELEFT"
      cMsg = cMsg & " " & LblAgent.Caption & " " & LblChr2.Caption
    Case "/MOVEUP"
      cMsg = cMsg & " " & LblAgent.Caption & " " & LblChr2.Caption
'    cMsg = "/MOVEDOWN Merlin Genie  - Move Merlin below Genie"
'    lstActive.AddItem cMsg
'    cMsg = "/MOVERIGHT Merlin Genie  - Move Merlin to right of Genie"
'    lstActive.AddItem cMsg
'    cMsg = "/MOVELEFT Merlin Genie  - Move Merlin to left Genie"
'    lstActive.AddItem cMsg
'    cMsg = "/MOVEUP Merlin Genie  - Move Merlin above Genie"
    
    End Select
    Me.txtCommand.Text = Trim(cMsg)
    If bRecord Then
     If bSay And txtSay.Text = "" Then
       txtSay.SetFocus
     Else
      Command3_Click
    End If
  End If
End Sub

Private Sub lstActive_DblClick()
 lstActive_Click
 CmdExecute_Click
End Sub

Private Sub lstAgents_Click()
    cmdShow.Visible = True
    cmdHide.Visible = True
    CmdStop.Visible = True
    Call cmdShow_Click
End Sub

Private Sub lstAgents_DblClick()
'    cmdShow.Visible = True
'    cmdHide.Visible = True
'    CmdStop.Visible = True
'    Call cmdShow_Click
End Sub

Private Sub lstGestures_Click()
LblStatus.Caption = ""
If Me.lstAgents.FileName <> "" Then
        'ExecuteCommand "/PLAY " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.lstGestures
        txtCommand.Text = "/PLAY " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.lstGestures
        If bRecord Then Command3_Click
    End If
       ' ExecuteCommand "/GESTURE " & Me.lstActive & " " & Me.lstGestures
End Sub

Private Sub lstGestures_DblClick()
    LblStatus.Caption = ""
    If Me.lstAgents.FileName <> "" Then
            'ExecuteCommand "/PLAY " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.lstGestures
            txtCommand.Text = "/PLAY " & Left(lstAgents.FileName, Len(lstAgents.FileName) - 4) & " " & Me.lstGestures
            If bRecord Then Command3_Click
        End If
        CmdExecute_Click
End Sub

Private Sub lstScript_Click()
    Me.txtCommand.Text = Me.lstScript.Text
   ' ExecuteCommand txtCommand.Text
End Sub

Private Sub lstScript_DblClick()
'    If Me.tcpClient(1).State = sckConnected Then
'        Me.tcpClient(1).SendData Me.lstScript.Text & vbCrLf
'    End If
End Sub







Private Function StripAgent(strFullPath As String) As String
    'Get last \ from path
    If strFullPath = "" Then
        StripAgent = ""
        Exit Function
    End If
    Dim strTemp As String
    strTemp = Mid(strFullPath, InStrRev(strFullPath, "\") + 1)
    StripAgent = Left(strTemp, Len(strTemp) - 4)
    
End Function
Private Sub ShutMeDown()
    End
    
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
    Dim PauseTime, Start, Finish, TotalTime
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
                Set agAgent = Me.MyAgent.Characters(strAgent)
                X = agAgent.Left ' / 15
                Y = agAgent.Top '- 150 '/ 15
                Text1.Text = strTemp & " " & X & " " & Y
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
                Set Request = agAgent.GestureAt(X, Y)
                bRequest = True
             Case "MOVELEFT"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Trim(Mid(strInstructions, InStr(1, strInstructions, " ") + 1))
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left + 100 ' / 15
                Y = agAgent.Top '- 150 '/ 15
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
               ' x = x + agAgent.Left / 15 + 120
               ' y = y + agAgent.Top / 15
               Set Request = agAgent.MoveTo(X, Y)
               bRequest = True
            Case "MOVERIGHT"
               strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left - 100 ' / 15
                Y = agAgent.Top '- 150 '/ 15
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
               ' x = x + agAgent.Left / 15 + 120
               ' y = y + agAgent.Top / 15
               Set Request = agAgent.MoveTo(X, Y)
               bRequest = True
            Case "MOVEDOWN"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left '+ 100 ' / 15
                Y = agAgent.Top + 150 '/ 15
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
               ' x = x + agAgent.Left / 15 + 120
               ' y = y + agAgent.Top / 15
                Set Request = agAgent.MoveTo(X, Y)
                bRequest = True
            Case "MOVEUP"
                strAgent = Left(strInstructions, InStr(1, strInstructions, " ") - 1)
                strTemp = Mid(strInstructions, InStr(1, strInstructions, " ") + 1)
                Set agAgent = Me.MyAgent.Characters(strTemp)
                X = agAgent.Left '+ 100 ' / 15
                Y = agAgent.Top - 150 '/ 15
                Set agAgent = Me.MyAgent.Characters(strAgent)
                agAgent.StopAll
               ' x = x + agAgent.Left / 15 + 120
               ' y = y + agAgent.Top / 15
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
              List1.AddItem "D " & strCommandIn
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
            List1.AddItem "D " & strCommandIn
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
          CharPosn(0) = agAgent.Left
          CharPosn(1) = agAgent.Top
           List1.AddItem "D " & strCommandIn
          Exit Sub
        End If
       Next
        'gAgent.StopAll
        List1.AddItem "N " & strCommandIn
        CharPosn(0) = agAgent.Left
        CharPosn(1) = agAgent.Top
    End If
    Exit Sub
ExitError:
      List1.AddItem "E " & strCommandIn
End Sub


Private Sub MyAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
  Set agAgent = Me.MyAgent.Characters(CharacterID)
CharPosn(0) = agAgent.Left
CharPosn(1) = agAgent.Top

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


Private Sub txtSay_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cMsg As String
If KeyCode = 13 Then
  If bRecord And txtSay.Text <> "" Then
     cMsg = Me.txtCommand.Text
     Me.txtCommand.Text = Me.txtCommand.Text & " " & txtSay.Text
     txtSay.Text = ""
     Command3_Click
      Me.txtCommand.Text = cMsg
   End If
 End If
End Sub


