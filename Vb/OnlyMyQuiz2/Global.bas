Attribute VB_Name = "Module1"
Option Explicit
Public strButton As String
Public bQuiz As Boolean
Public bNoIdle As Boolean
Public bStory As Boolean
Public bStart As Boolean
Public bHide As Boolean
Public bFirstTime As Boolean
Public bAgentActive As Boolean
Public Success As Boolean
Public retval As Long
Public Anim As String
Public Pname As String
Public Ap As String
Public cStoryAgent As String
Public MyAgent As IAgentCtlCharacterEx
Public cQuizfile As String          'Quiz word file name
Public cQuiz(5000, 6) As String     'Quiz questions
Public iQuizcnt As Integer          'Quiz count
Public cJoke(5000, 6) As String     'Quiz questions
Public iJokecnt As Integer          'Quiz count
Public Const BalloonOn As Integer = 1
Public Question(40) As String
Public Answer(40, 5) As String
Public strMovie(2000, 5) As String
Public iMovie As Integer
Public MyAnswer As String
Public GAMEFILE As String            ' Game logfile
Public YesNo(5) As String
Public Ans(5) As Integer
Public Q As Integer
Public Search As Integer
Public cCharacters(500) As String
Public iCharCnt As Integer
Public bRequestDone As Boolean
Public bRecord As Boolean
Public soundpath As String
Public strStory As String
Public cWordFile As String      ' Word file name
Public cWordProb(5000) As String     ' Word Problems
Public iWordCnt As Integer           ' Counter for word Problems
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub CentreMe(P1 As Form)

If TypeOf P1 Is Form Then
    P1.Left = (Screen.Width - P1.Width) / 2
    P1.Top = (Screen.Height - P1.Height) / 2
End If

End Sub
Public Sub ReadWord(cMyfile As String)
 Dim cMyLetter As String
   'cMyFile = cWordFile
  If Len(cMyfile) > 0 Then
   Dim MyString As String
   Dim MyStr2, Mystr3 As String
   Dim i, iSeq, iMath As Integer
   Open cMyfile For Input As #1 ' Open file for input.
   iWordCnt = 0
   Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
    'Debug.Print MyString,  ' Print data to Debug window.
    If Len(MyString) > 0 Then
    iWordCnt = iWordCnt + 1
    cWordProb(iWordCnt) = ""
    For i = 1 To Len(MyString)
     cMyLetter = UCase(Mid(MyString, i, 1))
     ' read word file
       If Asc(cMyLetter) > 64 And Asc(cMyLetter) < 91 Then
           cWordProb(iWordCnt) = cWordProb(iWordCnt) & cMyLetter
       End If
     Next
     End If
   Loop
   Close #1    ' Close file.
  End If
End Sub
Public Function GetSound(ByVal FileName) As String
'------------------------------------------------------------
' Load a sound file into a string variable.
' Taken from:
'   Mark Pruett
'   Black Art of Visual Basic Game Programming
'   The Waite Group, 1995
'------------------------------------------------------------
Dim buffer As String
Dim F As Integer
Dim SoundBuffer As String
On Error GoTo NoiseGet_Error
buffer = Space(1024)
SoundBuffer = ""
F = FreeFile
Open soundpath + "\" + FileName For Binary As F
Do While Not EOF(F)
  Get #F, , buffer     ' Load in 1K chunks
  SoundBuffer = SoundBuffer & buffer
Loop
Close F
GetSound = Trim(SoundBuffer)
Exit Function
NoiseGet_Error:
  SoundBuffer = ""
  Exit Function
End Function
Public Sub Main()
 '
    Dim i, iFoundm, iFoundw, iFoundMaster As Integer
    Dim MyFile, MyPath, MyName, MyString  As String
    Dim cTxt As String
    Dim cMysource As String
    Dim cDestination As String
    soundpath = App.Path & "\"
    
    Close #2
    cStoryAgent = "Peedy"
    FrmStart.Show
    'FrmBook.Show
End Sub
Public Sub ReadQuiz(cMyfile As String)
' Dim cMyfile As String
 Dim cMyLetter As String
  ' cMyfile = cQuizfile
  If Len(cMyfile) > 0 Then
   Dim MyString As String
   Dim MyStr2, Mystr3 As String
   Dim i, iSeq, iMath As Integer
   Open cMyfile For Input As #1 ' Open file for input.
   iQuizcnt = 0
   Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
    Debug.Print MyString,  ' Print data to Debug window.
    If UCase(Mid(MyString, 1, 1)) = "Q" Then
      iQuizcnt = iQuizcnt + 1
      cQuiz(iQuizcnt, 0) = MyString
      cQuiz(iQuizcnt, 1) = "*"
      cQuiz(iQuizcnt, 2) = "*"
      cQuiz(iQuizcnt, 3) = "*"
      cQuiz(iQuizcnt, 4) = "*"
      cQuiz(iQuizcnt, 5) = "N"
      iSeq = 0
     Else
      iSeq = iSeq + 1
      If iSeq < 5 Then
      cQuiz(iQuizcnt, iSeq) = MyString
      End If
     End If
    Loop
   Close #1    ' Close file.
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
Public Sub ReadJoke(cMyfile As String)
' Dim cMyfile As String
 Dim cMyLetter As String
  ' cMyfile = cQuizfile
  If Len(cMyfile) > 0 Then
   Dim MyString As String
   Dim MyStr2, Mystr3 As String
   Dim i, iSeq, iMath As Integer
   Open cMyfile For Input As #1 ' Open file for input.
   iQuizcnt = 0
   Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, MyString ' Read data into two variables.
    Debug.Print MyString,  ' Print data to Debug window.
    If UCase(Mid(MyString, 1, 1)) = "Q" Then
      iJokecnt = iJokecnt + 1
      cJoke(iJokecnt, 0) = MyString
      cJoke(iJokecnt, 1) = "*"
      cJoke(iJokecnt, 2) = "*"
      cJoke(iJokecnt, 3) = "*"
      cJoke(iJokecnt, 4) = "*"
      cJoke(iJokecnt, 5) = "N"
      iSeq = 0
     Else
      iSeq = iSeq + 1
      If iSeq < 5 Then
      cJoke(iJokecnt, iSeq) = MyString
      End If
     End If
    Loop
   Close #1    ' Close file.
  End If
End Sub

