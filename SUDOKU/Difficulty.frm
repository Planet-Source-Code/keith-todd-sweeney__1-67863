VERSION 5.00
Begin VB.Form Difficulty 
   BackColor       =   &H0000FFFF&
   Caption         =   "Difficulty"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   8160
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   7335
      Left            =   480
      Negotiate       =   -1  'True
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   0
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "Difficulty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Try As Boolean
Public IsButtonClicked As Boolean
Public Easy As Boolean
Public Medium As Boolean
Public Hard As Boolean
Public Xpos, Ypos As Single
Public GridStr As String
Public hi As Integer
Public Max As Integer

Sub Form_Load()
Startpic.Show
Call PlaySnd(Starting)
Call Wait(4)
Startpic.Hide


Difficulty.WindowState = 0
Difficulty.Move (Screen.Width - Difficulty.Width) / 2, (Screen.Height - Difficulty.Height) / 2
FileName.Hide
Difficulty.Show
Picture1 = LoadPicture(App.Path & "\Image4.bmp")
End Sub

'***************************************************************************
'Get the degree of difficulty required for the puzzle and create the puzzle*
'***************************************************************************
Public Function IsMouseOver(hWnd As Long) As Boolean
    Dim Mouse As POINTAPI
    GetCursorPos Mouse
    Diff = ""
    Easy = False
    Medium = False
    Hard = False
    If (Mouse.X <= 360 And Mouse.X >= 260) And (Mouse.Y <= 360 And Mouse.Y >= 260) Then
        IsMouseOver = True 'If mouse is over easy then true
        Easy = True
    End If
    If (Mouse.X <= 455 And Mouse.X >= 355) And (Mouse.Y <= 465 And Mouse.Y >= 365) Then
        IsMouseOver = True 'If mouse is over medium than true
        Medium = True
    End If
    If (Mouse.X <= 555 And Mouse.X >= 495) And (Mouse.Y <= 560 And Mouse.Y >= 460) Then
        IsMouseOver = True 'If mouse is over hard than true
        Hard = True
    End If
End Function

Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     PositionBefore = PositionNow
     Call api.GetCursorPos(PositionNow)
     Try = IsMouseOver(hWnd)
     Diff = ""
     If Easy Then Diff = "EASY"
     If Medium Then Diff = "MEDIUM"
     If Hard Then Diff = "HARD"
     If Button = 2 Then
       Difficulty.Hide
'get name of file to store or retrieve an old file
       FileName.Show
       FileName.Move (Screen.Width - FileName.Width) / 2, (Screen.Height - FileName.Height) / 2
       FileName.Picture1 = LoadPicture(App.Path & "\Image3.Bmp")
       Ipos = 4950
       Jpos = 6000
       Infile = InputBox("FILENAME(Either new or a saved File)", , , Ipos, Jpos)
     End If
     I = Len(Infile) - 9
     If Right$(Infile, 9) = "Saved.txt" Then
       OldFile = True
       Sname = Left$(Infile, I)
     End If
     FileName.Hide
   '
   'make up a puzzle
   '
   If Not OldFile Then
     Call PuzzleMaker(Puzzle)
   
  'Store Complete Puzzle File
  '
     Infile = Infile + "(" + Diff + ")"
 
    Filenum = FreeFile
    Open App.Path & "\" & "\Puzzles\Solutions\" & Infile & ".txt" For Output As #Filenum
    For I = 1 To 9
      For J = 1 To 9
         Print #Filenum, Puzzle(I, J);
      Next J
    Next I
   I = 1
   Close #Filenum
  ' Remove some of the numbers to form Sudoku puzzle
  '
    Call CreateGrid(Puzzle)
  
  'Store sudoku puzzle
  '
    Filenum = FreeFile
    Open App.Path & "\" & "\Puzzles\" & Diff & "\" & Infile & ".txt" For Output As #Filenum
    K = 1
    For I = 1 To 9
      For J = 1 To 9
         Print #Filenum, Puzzle(I, J);
         SavePuzz(K) = Puzzle(I, J)
         K = K + 1
      Next J
    Next I
    I = 1
    Close #Filenum
 End If
 FileName.Hide
 FormMain.Show

End Sub
Function PuzzleMaker(Puzzle)
Dim I, J, K, L, M, N As Integer

Randomize (Timer)
'
'Put Numbers 1-9 in first line of puzzle
'in Random order
'
J = 0
For I = 1 To 9
   J = J + 1
A:   N = Int(Rnd(1) * 9 + 1)
   Do While N < 1 And N > 9
    N = Int(Rnd(1) * 9 + 1)
   Loop
  Puzzle(1, I) = N
  If I > 1 Then
    K = I - 1
    For L = K To 1 Step -1
      If Puzzle(1, I) = Puzzle(1, L) Then GoTo A:
    Next L
  End If
Next I
'
'Create remaining eight lines of puzzle
'
Puzzle(2, 1) = Puzzle(1, 9): Puzzle(3, 1) = Puzzle(1, 8): Puzzle(4, 1) = Puzzle(1, 4)
Puzzle(2, 2) = Puzzle(1, 6): Puzzle(3, 2) = Puzzle(1, 4): Puzzle(4, 2) = Puzzle(1, 3)
Puzzle(2, 3) = Puzzle(1, 7): Puzzle(3, 3) = Puzzle(1, 5): Puzzle(4, 3) = Puzzle(1, 6)
Puzzle(2, 4) = Puzzle(1, 2): Puzzle(3, 4) = Puzzle(1, 9): Puzzle(4, 4) = Puzzle(1, 7)
Puzzle(2, 5) = Puzzle(1, 3): Puzzle(3, 5) = Puzzle(1, 1): Puzzle(4, 5) = Puzzle(1, 2)
Puzzle(2, 6) = Puzzle(1, 8): Puzzle(3, 6) = Puzzle(1, 7): Puzzle(4, 6) = Puzzle(1, 5)
Puzzle(2, 7) = Puzzle(1, 1): Puzzle(3, 7) = Puzzle(1, 3): Puzzle(4, 7) = Puzzle(1, 8)
Puzzle(2, 8) = Puzzle(1, 5): Puzzle(3, 8) = Puzzle(1, 2): Puzzle(4, 8) = Puzzle(1, 9)
Puzzle(2, 9) = Puzzle(1, 4): Puzzle(3, 9) = Puzzle(1, 6): Puzzle(4, 9) = Puzzle(1, 1)

Puzzle(5, 1) = Puzzle(1, 5): Puzzle(6, 1) = Puzzle(1, 7): Puzzle(7, 1) = Puzzle(1, 3)
Puzzle(5, 2) = Puzzle(1, 1): Puzzle(6, 2) = Puzzle(1, 9): Puzzle(7, 2) = Puzzle(1, 5)
Puzzle(5, 3) = Puzzle(1, 8): Puzzle(6, 3) = Puzzle(1, 2): Puzzle(7, 3) = Puzzle(1, 9)
Puzzle(5, 4) = Puzzle(1, 6): Puzzle(6, 4) = Puzzle(1, 8): Puzzle(7, 4) = Puzzle(1, 1)
Puzzle(5, 5) = Puzzle(1, 9): Puzzle(6, 5) = Puzzle(1, 4): Puzzle(7, 5) = Puzzle(1, 8)
Puzzle(5, 6) = Puzzle(1, 3): Puzzle(6, 6) = Puzzle(1, 1): Puzzle(7, 6) = Puzzle(1, 2)
Puzzle(5, 7) = Puzzle(1, 4): Puzzle(6, 7) = Puzzle(1, 5): Puzzle(7, 7) = Puzzle(1, 6)
Puzzle(5, 8) = Puzzle(1, 7): Puzzle(6, 8) = Puzzle(1, 6): Puzzle(7, 8) = Puzzle(1, 4)
Puzzle(5, 9) = Puzzle(1, 2): Puzzle(6, 9) = Puzzle(1, 3): Puzzle(7, 9) = Puzzle(1, 7)

Puzzle(8, 1) = Puzzle(1, 6): Puzzle(9, 1) = Puzzle(1, 2)
Puzzle(8, 2) = Puzzle(1, 8): Puzzle(9, 2) = Puzzle(1, 7)
Puzzle(8, 3) = Puzzle(1, 4): Puzzle(9, 3) = Puzzle(1, 1)
Puzzle(8, 4) = Puzzle(1, 3): Puzzle(9, 4) = Puzzle(1, 5)
Puzzle(8, 5) = Puzzle(1, 7): Puzzle(9, 5) = Puzzle(1, 6)
Puzzle(8, 6) = Puzzle(1, 9): Puzzle(9, 6) = Puzzle(1, 4)
Puzzle(8, 7) = Puzzle(1, 2): Puzzle(9, 7) = Puzzle(1, 9)
Puzzle(8, 8) = Puzzle(1, 1): Puzzle(9, 8) = Puzzle(1, 3)
Puzzle(8, 9) = Puzzle(1, 5): Puzzle(9, 9) = Puzzle(1, 8)


End Function
Function CreateGrid(Puzzle)
Dim GridStr(3, 20)
Dim Pos2(9)
'
'Create new Puzzle
'Decimal numbers are to create binary patterns for grids
'Based on binary zero or one (Zero square empty,One square occupied)
'Each decimal number represents a horizontal row of the puzzle.

'Easy
  GridStr(1, 1) = "159,414,483,91,368,408,4,373,174 "
  
'Medium
  GridStr(2, 1) = "91,368,387,273,297,273,387,170,91 "
  
'Hard
  GridStr(3, 1) = "387,130,325,40,108,40,325,130,395 "
   
  If Diff = "EASY" Then
    A = 1
  End If
  If Diff = "MEDIUM" Then
    A = 2
  End If
  If Diff = "HARD" Then
    A = 3
  End If
 
   '
   'Using Gridstr(A,1) for decimal number to be converted to binary pattern
   '
   
   C = Len(GridStr(A, 1))
   F = 1: E = 0
   For D = 2 To C
     If Mid$(GridStr(A, 1), D, 1) = "," Or Mid$(GridStr(A, 1), D, 1) = " " Then
       E = E + 1
       Hold(E) = Mid$(GridStr(A, 1), F, D - F)
       F = D + 1
     End If
   Next D
   
'Shuffle decimal numbers to give multiple options
J = 0
For I = 1 To 9
   J = J + 1
A:   N = Int(Rnd(1) * 9 + 1)
   Do While N < 1 And N > 9
    N = Int(Rnd(1) * 9 + 1)
   Loop
  Pos2(I) = N
  If I > 1 Then
    K = I - 1
    For L = K To 1 Step -1
      If Pos2(I) = Pos2(L) Then GoTo A:
    Next L
  End If
Next I
Hold(2) = Val(Hold(Pos2(2)))
Hold(4) = Val(Hold(Pos2(4)))

Hold(7) = Val(Hold(Pos2(7)))

'Transpose decimal numbers into binary pattern
'To determine if square is 1 or 0 i.e. Blank or occupied
   
   Call BinaryPos(Hold)

End Function
Sub BinaryPos(HH)
   '
   'Converting decimal number into binary pattern
   '
   
  Dim W, V, X, Y As Integer
  Dim Decnum As Single
  Dim Remainder(9) As Single
  Dim Binnum(9), Workstr As String
  V = 0: W = 0: Y = 0
  For V = 1 To 9
    Decnum = HH(V)
    For W = 0 To 8
        Remainder(W) = Decnum Mod 2
        Decnum = Decnum / 2
        Decnum = Int(Decnum)
     Next W
     For Y = 8 To 0 Step -1
       Binnum(V) = Binnum(V) + Str$(Remainder(Y))
     Next Y
 Next V

 For V = 1 To 9
    X = 0
    For W = 1 To 18 Step 2
        X = X + 1
        Workstr = Mid$(Binnum(V), W, 2)
        If Workstr = "0" Or Workstr = " 0" Or Workstr = "0 " Then
          Puzzle(V, X) = 0
        End If
    Next W
 Next V
  
  
End Sub
Private Sub Wait(interval) 'Delay program for interval (1= 1 second)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

