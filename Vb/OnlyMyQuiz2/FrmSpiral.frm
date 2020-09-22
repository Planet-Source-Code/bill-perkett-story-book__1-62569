VERSION 5.00
Begin VB.Form FrmSpiral 
   Caption         =   "Spiral"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox RichTextBox1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "FrmSpiral.frx":0000
      Top             =   600
      Width           =   8775
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Left            =   9360
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   5520
      Width           =   4215
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Print"
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox TxtInput 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "THIS.IS.A.TEST...."
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CmdSpiral 
      Caption         =   "Create Spiral"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label LblHeight 
      Caption         =   "Height of Spiral"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label LblWidth 
      Caption         =   "Width of Spirial"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Text Here"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSpiral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    FrmSpiral.Hide
End Sub

Private Sub CmdPrint_Click()
    Printer.FontName = "Courier New"
     Printer.FontSize = "14"
     'Printer.Font = RichTextBox1.Font
     Printer.Print "  "
     Printer.Print "  "
     Printer.Print RichTextBox1.Text
     'Set Printer.FontName = RichTextBox1.Font.Name
     Printer.EndDoc
End Sub


Private Sub CmdSpiral_Click()
'C*******************************************
'C*                                         *
'C* SPIR - GENERATE A SPIRAL ARRAY          *
'C*                                         *
'C*W PERKETT                                *
'C*******************************************
        Dim RR(100, 100) As String
        Dim WRD(9005) As String
        Dim CNT As Integer
        Dim iHZ As Integer
        Dim iVT As Integer
        Dim iEND As Integer
        Dim i As Integer
        Dim j As Integer
        Dim iEHZ As Integer
        Dim iEVT As Integer
        Dim KVT  As Integer
        Dim KHZ As Integer
        Dim L  As Integer
        Dim m  As Integer
        Dim MQ As Integer
        Dim MD As Integer
        CNT = 0
        ICHG = 1
        iHZ = 50 '16 '20
        iVT = 40  '20 '24
        iHZ = HScroll1.Value
        iVT = VScroll1.Value
        iEND = Len(TxtInput.Text)
        If iEND > 9000 Then iEND = 9000
        For i = 1 To iEND
          WRD(i) = Mid(TxtInput.Text, i, 1)
          If WRD(i) = " " Then WRD(i) = "."
        Next
'        WRD(1) = "B"
'        WRD(2) = "I"
'        WRD(3) = "L"
'        WRD(4) = "L"
'        WRD(5) = "."
        
'        WRITE(6,601)
'601     FORMAT('  HZ VT SZ',/)
'C
'C       DO 2  BLANK OUT ARRAY
'C       RR(IHZ,IVT)=ARRAY         WRD(I)=INPUT CHARACTERS.
'C       IHZ=HORIZONTAL            IVT=VERTICAL
'C       IEND=# OF INPUT CHARACTERS
'C
        For i = 1 To 100
        For j = 1 To 100
        RR(i, j) = " "
        Next
        Next
'        READ(5,301) IHZ,IVT,IEND
'301     FORMAT (3I3)
'        DISPLAY "ENTER "
'        READ(5,302) (WRD(M),M=1,IEND)
'302     FORMAT (60A1)
        iEHZ = iHZ
        iEVT = iVT
        KVT = 1
        KHZ = 0
        L = -1
        For MQ = 1 To 60
        L = -1 * L
        If (iHZ) <= O Then GoTo MY51
'C
'C HORIZONTAL
'C       DO 10  SET RR=TO CHARACTER ARRAY   L=+ FORWARD   L=- BACKWARD
'C
        For i = 1 To iHZ
        KHZ = KHZ + L
        If (CNT - iEND) >= 0 Then CNT = 0
        CNT = CNT + 1
        RR(KHZ, KVT) = WRD(CNT)
        Next i
        If (iHZ - ICHG) < 0 Then GoTo MY51
        iHZ = iHZ - ICHG
        If (iVT - ICHG) <= 0 Then GoTo MY51 'GoTo MY21
        iVT = iVT - ICHG
'C
'C VERTICAL
'C       DO 20  SET RR=TO CHARACTER ARRAY   L=+ FORWARD   L=- BACKWARD
'C
        m = KVT
        If m = 7 Then
          m = KVT
        End If
        For m = 1 To iVT
        KVT = KVT + L
        If (CNT - iEND) >= 0 Then CNT = 0
        CNT = CNT + 1
        RR(KHZ, KVT) = WRD(CNT)
        Next m
MY21:   ICHG = 3
        If (iHZ - ICHG) < 0 Then GoTo MY51
        Next MQ
MY51:   RichTextBox1.Text = ""
        For MD = 1 To iEVT
        RichTextBox1.Text = RichTextBox1.Text & "    "
        For j = 1 To iEHZ
        RichTextBox1.Text = RichTextBox1.Text & RR(j, MD)
         Next
'60      WRITE(6,600) (RR(J,MD),J=1,IEHZ)
        RichTextBox1.Text = RichTextBox1.Text & Chr$(13) & Chr$(10)  '600     FORMAT (' ',80A1)
        Next MD
     
End Sub

Private Sub Form_Activate()
If strStory <> "" Then
     TxtInput.Text = strStory & "..."
  End If
End Sub

Private Sub Form_GotFocus()
If strStory <> "" Then
     TxtInput.Text = strStory & "..."
  End If
End Sub

Private Sub Form_Load()
  HScroll1.Min = 20    ' Set values of the scroll bar.
  HScroll1.Max = 60
  HScroll1.Value = 60
  VScroll1.Min = 20    ' Set values of the scroll bar.
  VScroll1.Max = 45
  VScroll1.Value = 45
  RichTextBox1.Text = ""
  If strStory <> "" Then
     TxtInput.Text = strStory & "..."
  End If
End Sub

Private Sub HScroll1_Change()
LblWidth.Caption = "Width of spiral = " & HScroll1.Value
End Sub

Private Sub VScroll1_Change()
LblHeight.Caption = "Height of spiral = " & VScroll1.Value
End Sub
