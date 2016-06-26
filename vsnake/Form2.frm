VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   ScaleHeight     =   10995
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      Text            =   "40"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start game"
      Default         =   -1  'True
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   10575
      Left            =   0
      ScaleHeight     =   10515
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Matrix Width:"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ArrSnake(), mIndex As Integer, mWidth As Double, Running As Boolean
Dim Direction As String, Lost As Boolean
Dim n As Integer

Sub DrawSnake()
    Pic.Cls
    Dim X As Integer, Y As Integer
    For X = 1 To n
        For Y = 1 To n
            If ArrSnake(X, Y) <> 0 Then
                Pic.Line ((X - 1) * mWidth, (Y - 1) * mWidth)-(X * mWidth, Y * mWidth), IIf(ArrSnake(X, Y) = -1, vbRed, vbGreen), BF
            End If
        Next
    Next
End Sub




Private Sub Command1_Click()
    Form_Load
    Lost = False
    Command1.Enabled = False
    Dim Tr As Double
    Direction = "right"
    Pic.SetFocus
    Running = True
    Dim X As Integer, Y As Integer, HeadHandled As Boolean, X2 As Integer, Y2 As Integer, lTailX As Integer, lTailY As Integer
    Dim AteASeed As Boolean
Do While Running = True
    Tr = Timer
    Do Until Timer - Tr > 0.05:  DoEvents:  Loop
    HeadHandled = False: AteASeed = False: X2 = 0: Y2 = 0
    For X = 1 To n - 1
        For Y = 1 To n - 1
            If ArrSnake(X, Y) = 1 And HeadHandled = False Then
                X2 = X: Y2 = Y
                Select Case Direction
                    Case "left"
                        X2 = X - 1
                        Case "right"
                        X2 = X + 1
                    Case "up"
                        Y2 = Y - 1
                    Case "down"
                        Y2 = Y + 1
                End Select
                If ArrSnake(X2, Y2) > 1 Then Lost = True
                If X2 = 0 Or X2 >= n Or Y2 = 0 Or Y2 >= n Then Lost = True
                If ArrSnake(X2, Y2) = -1 Then AteASeed = True
                ArrSnake(X2, Y2) = 1
                ArrSnake(X, Y) = 2
                HeadHandled = True
            ElseIf ArrSnake(X, Y) <> 0 And Not (X = X2 And Y = Y2) And ArrSnake(X, Y) <> -1 Then
                If ArrSnake(X, Y) <> mIndex Then
                    ArrSnake(X, Y) = ArrSnake(X, Y) + 1
                Else
                    ArrSnake(X, Y) = 0
                    lTailX = X: lTailY = Y
                End If
            End If
        Next
    Next
    If Lost = True Then Call DrawLostSnake: Command1.Enabled = True: Exit Sub
    DrawSnake
    If AteASeed = True Then
        mIndex = mIndex + 1
        ArrSnake(lTailX, lTailY) = mIndex
        LoadNewSeed
    End If
    'Timer1.Enabled = False
    DoEvents
Loop
End
End Sub

Private Sub Form_Load()
    If txtWidth.Text < 15 Or IsNumeric(txtWidth.Text) = False Then MsgBox "Please Enter an integer greater than or equal to 15." & vbCrLf & "Default value 40 will be used instead.", vbExclamation, "Snake": txtWidth.Text = "40"
    n = txtWidth.Text
    mWidth = Me.Pic.Width / n
    ReDim ArrSnake(n, n)
    Dim i As Integer, St As Integer, En As Integer
    St = 1: En = 5
    For i = St To En
        ArrSnake(i, 5) = En - i + 1
    Next
    mIndex = En - St + 1
    ArrSnake(Int(n / 2), Int(n / 2)) = -1
    DrawSnake
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim mDirection As String
    mDirection = LCase(Direction)
    Select Case KeyCode
        Case vbKeyLeft
            If mDirection <> "right" Then Direction = "left"
        Case vbKeyRight
            If mDirection <> "left" Then Direction = "right"
        Case vbKeyUp
            If mDirection <> "down" Then Direction = "up"
        Case vbKeyDown
            If mDirection <> "up" Then Direction = "down"
    End Select
End Sub

Private Sub Pic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Running = False
End Sub

Sub LoadNewSeed()
    Randomize Timer
    Dim nX As Integer, nY As Integer
    
    Do
        nX = Int(Rnd * n)
        nY = Int(Rnd * n)
    Loop While ArrSnake(nX, nY) <> 0 Or nX = 0 Or nY = 0
    ArrSnake(nX, nY) = -1
    DrawSnake
End Sub

Sub DrawLostSnake()
    Dim X As Single, Y As Single
    Pic.Cls
    For X = 1 To n
        For Y = 1 To n
            If ArrSnake(X, Y) >= 1 Then Pic.Line ((X - 1) * mWidth, (Y - 1) * mWidth)-(X * mWidth, Y * mWidth), vbYellow, BF
        Next
    Next
    Pic.Refresh
    Dim mAns As VbMsgBoxResult
    mAns = MsgBox("You lost this game ;" & vbCrLf & "Do you want to start a new game?", vbQuestion + vbYesNo + vbDefaultButton1, "Lost")
    If mAns = vbYes Then Command1_Click
End Sub
