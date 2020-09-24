VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBNoid"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6615
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleMode       =   0  'User
   ScaleWidth      =   6514.032
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer BigTimer 
      Left            =   45
      Top             =   675
   End
   Begin VB.Timer TimerShorHide 
      Interval        =   20000
      Left            =   45
      Top             =   1710
   End
   Begin VB.Timer TimerHide 
      Left            =   45
      Top             =   1215
   End
   Begin VB.Timer TimerReverse 
      Left            =   540
      Top             =   675
   End
   Begin VB.Timer Timer2 
      Left            =   675
      Top             =   1170
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   1710
   End
   Begin VB.Image Paddle 
      Height          =   225
      Left            =   1620
      Picture         =   "Form1.frx":08E2
      Stretch         =   -1  'True
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label LabelPaused 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PAUSED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   825
      Left            =   45
      TabIndex        =   5
      Top             =   2475
      Width           =   6405
   End
   Begin VB.Image Ball 
      Height          =   150
      Index           =   4
      Left            =   2115
      Picture         =   "Form1.frx":16A4
      Stretch         =   -1  'True
      Top             =   495
      Width           =   195
   End
   Begin VB.Image Ball 
      Height          =   150
      Index           =   3
      Left            =   1845
      Picture         =   "Form1.frx":19E6
      Stretch         =   -1  'True
      Top             =   495
      Width           =   195
   End
   Begin VB.Image Ball 
      Height          =   150
      Index           =   2
      Left            =   1620
      Picture         =   "Form1.frx":1D28
      Stretch         =   -1  'True
      Top             =   495
      Width           =   195
   End
   Begin VB.Image Ball 
      Height          =   150
      Index           =   1
      Left            =   1395
      Picture         =   "Form1.frx":206A
      Stretch         =   -1  'True
      Top             =   495
      Width           =   195
   End
   Begin VB.Label LabelNumLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "LEVEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   825
      Left            =   45
      TabIndex        =   4
      Top             =   1440
      Width           =   6405
   End
   Begin VB.Label LabelLevel 
      BackColor       =   &H00000000&
      Caption         =   "LEVEL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   4680
      TabIndex        =   3
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label LabelPoints 
      BackColor       =   &H00000000&
      Caption         =   "POINTS:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   4635
      TabIndex        =   2
      Top             =   630
      Width           =   1500
   End
   Begin VB.Label LabelSpeed 
      BackColor       =   &H00000000&
      Caption         =   "SPEED:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3240
      TabIndex        =   1
      Top             =   585
      Width           =   1500
   End
   Begin VB.Label LabelLives 
      BackColor       =   &H00000000&
      Caption         =   "LIVES:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3330
      TabIndex        =   0
      Top             =   180
      Width           =   1500
   End
   Begin VB.Image Ball 
      Height          =   150
      Index           =   0
      Left            =   1215
      Picture         =   "Form1.frx":23AC
      Stretch         =   -1  'True
      Top             =   495
      Width           =   195
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "New"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Const SND_ASYNC = &H1            'Play asynchronously
Const SND_NODEFAULT = &H2        'Don't use default sound
Const SND_MEMORY = &H4           'lpszSoundName points to a memory file
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Private Type POINTAPI
        x As Long
        Y As Long
End Type
Dim a As POINTAPI



Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYCAPTION = 4

Dim LevelPaused As Long
Dim BX, BY As Long
Dim Lives As Integer
Dim Score As Integer
Dim Level As Integer
Dim Speed As Integer
Dim InGame As Integer
Dim InLevel As Integer
Dim BallReleased(5) As Integer
Dim BallType(5) As Integer
Dim Rows As Integer
Dim Cols As Integer
Dim BrickWidth As Integer
Dim BrickHeight As Integer
Dim bricks() As Image
Dim bricksOn() As Integer
Dim bricksHide() As Date
Dim TotalHided As Long
Dim bricksRoll(1 To 100) As Image
Dim bricksRollOn(1 To 100) As Long
Dim RollUsed, RollAction As Long
Dim dballx(5), dbally(5) As Long
Dim BallDirectionX(5), BallDirectionY(5), BallAngle(5) As Long
Dim BallX(5), BallY(5) As Double
Dim LevelFinished As Long
Dim FormErrorX As Long
Dim LevelMinSpeed, LevelMaxSpeed As Long
Dim StartLevelTime As Date
Dim NewBallAngleTime(5) As Date
Dim PlayerPoints As Long
Dim ReverseControls As Long
Dim ApplicationClosing As Long

Public Function app_path() As String
    Dim x As String
    x = App.Path
    If Right$(x, 1) <> "\" Then x = x + "\"
    app_path = UCase$(x)
End Function

Public Sub CenterForm(Frm As Form)
    Dim Left, Top As Integer

    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - Frm.Width / 2
    Top = Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) + GetSystemMetrics(SM_CYCAPTION)) / 2 - Frm.Height / 2

    Frm.Move Left, Top
End Sub

Private Sub Timer1_Timer()
    Dim i, Ret, rnd1, newcenter, newleft As Long
    Dim xx As Long
    Dim iball As Long
    
    rnd1 = Rnd(100)
    Ret = GetCursorPos(a)
    
    If (InGame = 1 And LevelPaused = 0) Then
        newcenter = a.x * Screen.TwipsPerPixelX
        newcenter = newcenter - Form1.Left - Paddle.Width / 2
        newleft = newcenter
        
        If (newleft < 0) Then
            newleft = 0
        End If
        If (newleft + Paddle.Width >= Form1.Width - FormErrorX) Then
            newleft = Form1.Width - Paddle.Width
        End If
        If (ReverseControls = 0) Then
            For iball = 0 To 4
                If (Ball(iball).Visible = True And BallReleased(iball) = 0) Then
                    xx = Ball(iball).Left - Paddle.Left
                    Ball(iball).Left = newleft + xx
                    BallX(iball) = Ball(iball).Left
                End If
            Next iball
            If (Abs(Paddle.Left - newleft) > 1) Then
                Paddle.Left = CLng(newleft)
            End If
        Else
            For iball = 0 To 4
                If (Ball(iball).Visible = True And BallReleased(iball) = 0) Then
                    xx = Ball(iball) - Paddle.Left
                    Ball(iball).Left = Form1.Width - Paddle.Width - newleft + xx
                    BallX(iball) = Ball(iball).Left
                End If
            Next iball
            Paddle.Left = Form1.Width - Paddle.Width - newleft
        End If
    End If
    
    If (InLevel = 1 And LevelPaused = 0) Then
        For i = 1 To RollUsed
            MoveBrickRoll (i)
        Next i
    End If
End Sub

Private Sub Timer2_Timer()
    Dim sec As Long
    Dim iball As Long
    
    If (InLevel = 1 And LevelPaused = 0) Then
        For iball = 0 To 4
            If (Ball(iball).Visible = True And BallReleased(iball) = 1) Then
                sec = DateDiff("s", StartLevelTime, Now)
                If (sec > 20) Then
                    dbally(iball) = dbally(iball) + 10
                    If dbally(iball) > LevelMaxSpeed Then
                        dbally(iball) = LevelMaxSpeed
                    End If
                    UpdatePanel
                    StartLevelTime = Now
                End If
                dballx(iball) = Abs(50 * Sin(BallAngle(iball) * 3.14159265358979 / 180))
                
                If (BallDirectionX(iball) = 1) Then
                    BallX(iball) = BallX(iball) - dballx(iball)
                Else
                    BallX(iball) = BallX(iball) + dballx(iball)
                End If
                
                If (BallDirectionY(iball) = 1) Then
                    BallY(iball) = BallY(iball) - dbally(iball)
                Else
                    BallY(iball) = BallY(iball) + dbally(iball)
                End If
                
                Ball(iball).Left = BallX(iball)
                Ball(iball).Top = BallY(iball)
                
                If (Ball(iball).Left < 0) Then
                    BallDirectionX(iball) = 2
                End If
                If (Ball(iball).Left > Form1.Width - Ball(iball).Width - FormErrorX) Then
                    BallDirectionX(iball) = 1
                End If
                
                If (Ball(iball).Top < 0) Then
                    BallDirectionY(iball) = 2
                End If
                If (Ball(iball).Top > Form1.Height - Ball(iball).Height) Then
                    BallDirectionY(iball) = 1
                End If
            End If
        Next iball
        

        DoEvents
        CheckHitbricks
        CheckHitPaddle
        CheckRollHitPaddle
        ClearRolls (0)
        CheckLostBall
        CheckLevelFinish
    End If
    
    If (InGame = 0 And mnuGameNew.Enabled = True) Then
        LabelNumLevel.Visible = True
    End If
    
    If (LevelPaused = 1 And LabelPaused.Visible = False) Then
        LabelPaused = "PAUSED..."
        LabelPaused.Left = Form1.Width / 2 - LabelPaused.Width / 2
        LabelPaused.Top = Form1.Height / 2 - LabelPaused.Height / 2
        LabelPaused.Visible = True
        DoEvents
    ElseIf (LevelPaused = 0 And LabelPaused.Visible = True) Then
        LabelPaused.Visible = False
    End If
End Sub

Private Sub BigTimer_Timer()
    Paddle.Width = 1000
    BigTimer.Interval = 0
End Sub

Private Sub TimerHide_Timer()
    Unhidebricks
    TimerHide.Interval = 0
End Sub

Private Sub TimerReverse_Timer()
    ReverseControls = 0
    TimerReverse.Interval = 0
End Sub

Private Sub TimerShorHide_Timer()
    On Error Resume Next
    Dim diff, i, j As Long
    
    If (InLevel = 1) Then
        For i = 1 To Rows
            For j = 1 To Cols
                If (bricksOn(i, j) = 19) Then
                    diff = DateDiff("s", bricksHide(i, j), Now)
                    If (diff > 20) Then
                        bricks(i, j).Visible = True
                        bricks(i, j).Picture = LoadPicture("./Images/Brick" & bricksOn(i, j) & ".bmp")
                        bricksHide(i, j) = DateAdd("yyyy", -1, Now)
                        DoEvents
                    End If
                End If
            Next j
        Next i
    End If
End Sub

Private Sub Form_Activate()
    ApplicationClosing = 0
    mnuGameNew_Click
End Sub

Private Sub Form_Load()
    BrickWidth = 480
    BrickHeight = 240
    Paddle.Width = 1000
    FormErrorX = 270
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iball As Long
    
    For iball = 0 To 4
        BallReleased(iball) = 1
    Next iball
    ApplicationClosing = 1
    Lives = 0
End Sub

Public Sub UpdatePanel()
    LabelLives = "LIVES: " & Lives
    LabelSpeed = "Speed: " & dbally(0)
    LabelLevel = "Level: " & Level
    LabelPoints = "Points: " & PlayerPoints
    'Form1.Caption = RollUsed
    DoEvents
End Sub

Public Sub SetNormalValues()
    Dim iball As Long
    
    LevelPaused = 0
    ReverseControls = 0
    Paddle.Width = 1000
    BigTimer.Interval = 0
    TimerReverse.Interval = 0
    TimerHide.Interval = 0
    
    For iball = 0 To 4
        Ball(iball).Picture = LoadPicture(app_path() & "/Images/ball.bmp")
        If (iball = 0) Then
            Ball(iball).Visible = True
            
        Else
            Ball(iball).Visible = False
        End If
        
        Ball(iball).Left = Paddle.Left + Paddle.Width / 2
        BallX(iball) = Ball(iball).Left
        Ball(iball).Top = Paddle.Top - 200
        BallY(iball) = Ball(iball).Top
        BallType(iball) = 1
        BallReleased(iball) = 0
        dbally(iball) = LevelMinSpeed
        BallType(iball) = 1
    Next iball
    ClearRolls (1)
End Sub

Public Sub RemoveBrickControls()
    On Error Resume Next
    Dim i, j As Long
    
    For i = 1 To Rows
        For j = 1 To Cols
            Controls.Remove ("Image" & i & "_" & j)
        Next j
    Next i
End Sub

Public Sub CalculateBallAngle(bb As Long)
    Dim x1, x2, x3, BL, BDX, BDY As Long
    Dim pc As Double
    
    
    If (Ball(bb).Visible = True) Then
        If (BallAngle(bb) <= 0 Or BallAngle(bb) >= 90) Then
            BallAngle(bb) = 45
        End If
        BL = BallAngle(bb)
        BDX = BallDirectionX(bb)
        BDY = BallDirectionY(bb)
        x1 = Ball(bb).Left + Ball(bb).Width / 2
        x2 = Paddle.Left + Paddle.Width / 2
        x3 = Abs(x2 - x1)
        pc = 100 * x3 / (Paddle.Width / 2)
        
        If (x2 <= x1) Then
            If (BDX = 1) Then
                If (BL < pc) Then
                    BallDirectionX(bb) = 2
                End If
            End If
        Else
            If (BDX = 2) Then
                If (BL < pc) Then
                    BallDirectionX(bb) = 1
                End If
            End If
        End If
        
        If (BDX <> BallDirectionX(bb)) Then
            BallAngle(bb) = BL
        Else
            If (x2 <= x1) Then
                If (BDX = 1) Then
                    BallAngle(bb) = BL - BL * pc / 100
                Else
                    BallAngle(bb) = BL + (90 - BL) * (pc - 10) / 100
                End If
            Else
                If (BDX = 1) Then
                    BallAngle(bb) = BL + (90 - BL) * (pc - 10) / 100
                Else
                    BallAngle(bb) = BL - BL * pc / 100
                End If
            End If
            If (BallAngle(bb) > 80) Then
                 BallAngle(bb) = BallAngle(bb) - 5
            End If
            If (BallAngle(bb) < 10) Then
                 BallAngle(bb) = BallAngle(bb) + 5
            End If
        End If
        
        If (BallReleased(bb) = 0) Then
            BallAngle(bb) = 15
        End If
        NewBallAngleTime(bb) = Now
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim x1, x2, x3 As Long
    Dim iball As Long
    
    If (Button = 2) Then
        EndLevel
    End If
        
    For iball = 0 To 4
        If (BallReleased(iball) = 0) Then
            If (Ball(iball).Visible = True) Then
                x1 = Paddle.Left
                x2 = Paddle.Left + Paddle.Width
                x3 = Ball(iball).Left + Ball(iball).Width / 2
                If (x3 > x1 And x3 < x2) Then
                    If (x3 < x1 + (x2 - x1) / 2) Then
                        BallDirectionX(iball) = 1 'left
                    Else
                        BallDirectionX(iball) = 2 'right
                    End If
                    BallDirectionY(iball) = 1
                    Call CalculateBallAngle(iball)
                    BallReleased(iball) = 1
                    BallType(iball) = 1
                    Ball(iball).Picture = LoadPicture(app_path() & "/Images/ball.bmp")
                End If
            End If
        End If
    Next iball
End Sub

Private Sub mnuGameNew_Click()
    On Error Resume Next
    Dim i, j As Long
    
    mnuGameNew.Enabled = False
    Lives = 3
    UpdatePanel
    InitializeGame
    While (Lives > 0)
        UpdatePanel
        InitializeLevel
        PlayLevel
        Level = Level + 1
    Wend
    
    InGame = 0
    mnuGameNew.Enabled = True
    LabelNumLevel = "GAME OVER"
    LabelNumLevel.Left = Form1.Width / 2 - LabelNumLevel.Width / 2
    LabelNumLevel.Top = Form1.Height / 2 - LabelNumLevel.Height / 2
    LabelNumLevel.Visible = True
    RemoveBrickControls
    LabelNumLevel.Visible = False
    If (ApplicationClosing = 1) Then
        End
    End If
End Sub

Public Sub InitializeGame()
    Score = 0
    Level = 1
    InGame = 1
    InLevel = 0
    PlayerPoints = 0
    ReverseControls = 0
    Timer1.Interval = 10
End Sub

Public Sub InitializeLevel()
    On Error Resume Next
    Dim i, j, impo As Long
    
    BX = 0
    BY = 0
    SetNormalValues
    RollUsed = 0
    RollAction = 0
    
    LabelNumLevel = "LEVEL " & Level
    LabelNumLevel.Left = Form1.Width / 2 - LabelNumLevel.Width / 2
    LabelNumLevel.Top = Form1.Height / 2 - LabelNumLevel.Height / 2
    LabelNumLevel.Visible = True
    DoEvents
    
    RemoveBrickControls
    
    Rows = 5 + (Level - 1)
    If (Rows > 10) Then
        Rows = 10
    End If
    Cols = 15 + (Level - 1)
    If (Cols > 13) Then
        Cols = 13
    End If
    Call CenterForm(Form1)
    impo = ImportLevel(Level)
    If (impo < 0) Then
        Drawbricks
    End If
    LevelMinSpeed = 10 + (Level - 1) * 5
    LevelMaxSpeed = 30 + Level * 5
    
    LevelMinSpeed = Level / 2
    LevelMaxSpeed = Level
    
    If (LevelMinSpeed < 10) Then
        LevelMinSpeed = 10
    End If
    If (LevelMinSpeed > 40) Then
        LevelMinSpeed = 40
    End If
    
    If (LevelMaxSpeed < 30) Then
        LevelMaxSpeed = 30
    End If
    If (LevelMaxSpeed > 70) Then
        LevelMaxSpeed = 70
    End If
    
    LabelNumLevel.Visible = False
    
    SetNormalValues
    
    Timer2.Interval = 1
End Sub

Public Sub Drawbricks()
    Dim i, j, rand1 As Long
    
    ReDim bricks(Rows + 1, Cols + 1)
    ReDim bricksOn(Rows + 1, Cols + 1)
    ReDim bricksHide(Rows + 1, Cols + 1)
    
    For i = 1 To Rows
        For j = 1 To Cols
            Set bricks(i, j) = Controls.Add("vb.image", "Image" & i & "_" & j)
            bricksHide(i, j) = DateAdd("yyyy", -1, Now)
            If (Level < 8) Then
                bricksOn(i, j) = 1 + Level * Rnd(100)
            Else
                bricksOn(i, j) = 1 + 8 * Rnd(100)
            End If
            If (bricksOn(i, j) > 8) Then
                bricksOn(i, j) = 8
            End If
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            rand1 = 10000 * Rnd(100)
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 9
                bricks(i, j).Picture = LoadPicture("./Images/Brick9.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 10
                bricks(i, j).Picture = LoadPicture("./Images/Brick10.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 11
                bricks(i, j).Picture = LoadPicture("./Images/Brick11.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 12
                bricks(i, j).Picture = LoadPicture("./Images/Brick12.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 13
                bricks(i, j).Picture = LoadPicture("./Images/Brick13.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 14
                bricks(i, j).Picture = LoadPicture("./Images/Brick14.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 15
                bricks(i, j).Picture = LoadPicture("./Images/Brick15.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 16
                bricks(i, j).Picture = LoadPicture("./Images/Brick16.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 17
                bricks(i, j).Picture = LoadPicture("./Images/Brick17.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 18
                bricks(i, j).Picture = LoadPicture("./Images/Brick18.bmp")
            End If
            If rand1 > 10000 - Level Then
                bricksOn(i, j) = 19
                bricks(i, j).Picture = LoadPicture("./Images/Brick19.bmp")
            End If
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricks(i, j).Top = i * BrickHeight
            bricks(i, j).Left = 10 + (j - 1) * BrickWidth
            bricks(i, j).Width = BrickWidth
            bricks(i, j).Height = BrickHeight
            bricks(i, j).Picture = LoadPicture("./Images/Brick" & bricksOn(i, j) & ".bmp")
            bricks(i, j).Stretch = True
            If (bricksOn(i, j) > 0) Then
                bricks(i, j).Visible = True
            Else
                bricks(i, j).Visible = False
            End If
        Next j
    Next i
    
    Form1.Width = (Cols + 1) * BrickWidth - FormErrorX
    Form1.Height = Screen.TwipsPerPixelY * 500
    Paddle.Top = Form1.Height - 1400
    
    Paddle.Left = Form1.Width / 2 - Paddle.Width / 2
    LabelLives.Top = Form1.Height - 1200
    LabelSpeed.Top = Form1.Height - 1200
    LabelPoints.Top = Form1.Height - 1200
    LabelLevel.Top = Form1.Height - 1200
    
    LabelLives.Left = 10
    LabelLevel.Left = LabelLives.Left + LabelLives.Width + 100
    LabelPoints.Left = LabelLevel.Left + LabelLevel.Width + 100
    LabelSpeed.Left = LabelPoints.Left + LabelPoints.Width + 100
    
    Call CenterForm(Form1)
End Sub

Public Sub PlayLevel()
    Dim iball As Long
    
    LevelFinished = 0
    While (Lives > 0 And LevelFinished = 0)
        DoEvents
        If (InLevel = 0) Then
            SetNormalValues
            StartLevelTime = Now()
            UpdatePanel
            InLevel = 1
        End If
    Wend
    InLevel = 0
End Sub

Public Sub CheckHitbricks()
    Dim i, j As Integer
    Dim iball As Long
    
    For i = 1 To Rows
        For j = 1 To Cols
            For iball = 0 To 4
                If (bricksOn(i, j) > 0) Then
                    If (Ball(iball).Visible = True) Then
                        Call CheckHitBrick(i, j, iball)
                    End If
                End If
            Next iball
        Next j
    Next i
End Sub

Public Sub CheckHitBrick(ByVal row As Long, ByVal col As Long, ByVal bb As Long)
    On Error Resume Next
    Dim BallX2, BallY2 As Double
    Dim x1, x2, x3, IsHided As Long
    Dim a1, a2, a3, a4 As Long
    Dim b1, b2, b3, b4 As Long
    Dim BounceX, BounceY As Long
    Dim rnd1, rnd2, sec As Long
    Dim UpExist, DownExist As Long
    Dim LeftExist, RightExist As Long
    Dim BDX, BDY As Long
    
    IsHided = 0
    x1 = bricks(row, col).Left
    x2 = bricks(row, col).Left + bricks(row, col).Width
    x3 = Ball(bb).Left
    
    a1 = bricks(row, col).Left
    a2 = bricks(row, col).Left + bricks(row, col).Width
    a3 = Ball(bb).Left
    a4 = Ball(bb).Left + Ball(bb).Width
    
    b1 = bricks(row, col).Top
    b2 = bricks(row, col).Top + bricks(row, col).Height
    b3 = Ball(bb).Top
    b4 = Ball(bb).Top + Ball(bb).Height
    
    BallX2 = BallX(bb) + Ball(bb).Width
    BallY2 = BallY(bb) + Ball(bb).Height
    
    If (bricks(row, col).Visible = True) Then
        If ((BallX(bb) >= bricks(row, col).Left And BallX(bb) <= bricks(row, col).Left + bricks(row, col).Width) Or (BallX2 >= bricks(row, col).Left And BallX2 <= bricks(row, col).Left + bricks(row, col).Width)) Then
            If ((BallY(bb) >= bricks(row, col).Top And BallY(bb) <= bricks(row, col).Top + bricks(row, col).Height) Or (BallY2 >= bricks(row, col).Top And BallY2 <= bricks(row, col).Top + bricks(row, col).Height)) Then
            
                If (bricks(row, col).Visible = True) Then
                    Call PlaySound(app_path() & "Sounds/sound2.wav", 0, SND_ASYNC Or SND_NOSTOP)
                End If
                
                If (bricksOn(row, col) <> 18 And bricksOn(row, col) <> 19 And bricksOn(row, col) > 0) Then
                    AddPlayerPoints (bricksOn(row, col))
                End If
                
                rnd1 = 90 * Rnd(100)
                If (rnd1 > BallAngle(bb)) Then
                    BallAngle(bb) = BallAngle(bb) + 1
                Else
                    BallAngle(bb) = BallAngle(bb) - 1
                End If
                
                'sec = Abs(DateDiff("s", Now, NewBallAngleTime(bb)))
                'If (sec > 20) Then
                '    rnd1 = 1 + 100 * Rnd(100)
                '    If (rnd1 > 95) Then
                '        dbally(bb) = dbally(bb) - 2
                '        If (dbally(bb) < 10) Then
                '            dbally(bb) = 10
                '        End If
                '        If (BallDirectionX(bb) = 1) Then
                '            BallAngle(bb) = BallAngle(bb) - 3
                '        Else
                '            BallAngle(bb) = BallAngle(bb) + 1
                '        End If
                '        If (BallAngle(bb) < 5) Then
                '            BallAngle(bb) = 6
                '        End If
                '        If (BallAngle(bb) > 85) Then
                '            BallAngle(bb) = 84
                '        End If
                '        If (sec > 120) Then
                '            BallAngle(bb) = 1
                '        End If
                '        UpdatePanel
                '    End If
                'End If
                
                bricks(row, col).Picture = LoadPicture("./Images/Brickflash.bmp")
                DoEvents
                bricks(row, col).Picture = LoadPicture("./Images/Brickflash.bmp")
                DoEvents
                bricks(row, col).Picture = LoadPicture("./Images/Brick" & bricksOn(row, col) & ".bmp")
                
                If (bricksOn(row, col) = 9) Then 'Bomb!
                    Call ExplodeBomb(row, col)
                ElseIf (bricksOn(row, col) = 10) Then 'Extra Life!
                    bricksOn(row, col) = 0
                    Lives = Lives + 1
                    UpdatePanel
                ElseIf (bricksOn(row, col) = 11) Then 'End of Level!
                    EndLevel
                ElseIf (bricksOn(row, col) = 12) Then 'Big Paddle!
                    bricksOn(row, col) = 0
                    Paddle.Width = 2000
                    BigTimer.Interval = 10000
                ElseIf (bricksOn(row, col) = 13) Then 'Slow down!
                    bricksOn(row, col) = 0
                    dbally(bb) = dbally(bb) - 10
                    If dbally(bb) < LevelMinSpeed Then
                        dbally(bb) = LevelMinSpeed
                    End If
                    UpdatePanel
                ElseIf (bricksOn(row, col) = 14) Then 'Reverse!
                    bricksOn(row, col) = 0
                    ReverseControls = 1
                    TimerReverse.Interval = 10000
                ElseIf (bricksOn(row, col) = 15) Then 'Hide bricks!
                    bricksOn(row, col) = 0
                    Hidebricks
                    TimerHide.Interval = 10000
                ElseIf (bricksOn(row, col) = 16) Then 'Decrease Values!
                    bricksOn(row, col) = 0
                    DecreaseValues
                ElseIf (bricksOn(row, col) = 17) Then 'Increase Values!
                    bricksOn(row, col) = 0
                    IncreaseValues
                ElseIf (bricksOn(row, col) = 18) Then 'Non Beatable block!
                    bricksOn(row, col) = bricksOn(row, col)
                    bricks(row, col).Picture = LoadPicture("./Images/Brickflash.bmp")
                    DoEvents
                    bricks(row, col).Picture = LoadPicture("./Images/Brick" & bricksOn(row, col) & ".bmp")
                ElseIf (bricksOn(row, col) = 19) Then 'Shortly hided block!
                    IsHided = CheckIfHided(row, col)
                    If (IsHided = 0) Then
                        bricksOn(row, col) = bricksOn(row, col)
                        bricks(row, col).Picture = LoadPicture("./Images/Brick0.bmp")
                        bricks(row, col).Visible = False
                        DoEvents
                        bricksHide(row, col) = Now
                    End If
                ElseIf (bricksOn(row, col) = 101 Or bricksOn(row, col) = 102 Or bricksOn(row, col) = 103 Or bricksOn(row, col) = 104 Or bricksOn(row, col) = 105 Or bricksOn(row, col) = 106 Or bricksOn(row, col) = 107 Or bricksOn(row, col) = 108 Or bricksOn(row, col) = 109 Or bricksOn(row, col) = 1110) Then 'Simple coloured brick
                    bricksOn(row, col) = 0
                    bricks(row, col).Picture = LoadPicture("./Images/Brickflash.bmp")
                Else
                    bricksOn(row, col) = bricksOn(row, col) - 1
                End If
                If (bricksOn(row, col) = 0) Then
                    rnd1 = 100 * Rnd(100)
                    rnd2 = 90 + Level
                    If (rnd2 > 90) Then
                        rnd2 = 90
                    End If
                    If (rnd1 > rnd2) Then
                        Call CreateNewBrickRoll(row, col)
                    End If
                    bricks(row, col).Visible = False
                    CheckLevelFinish
                Else
                    bricks(row, col).Picture = LoadPicture("./Images/Brick" & bricksOn(row, col) & ".bmp")
                End If
                
                
                'Ball movement
                If (IsHided = 0 And (bricksOn(row, col) = 18 Or BallType(bb) = 1)) Then
                    ' Bounce on Y axis
                    BounceX = 0
                    BounceY = 0
                    UpExist = 0
                    DownExist = 0
                    BDY = BallDirectionY(bb)
                    BDX = BallDirectionX(bb)
                    
                    If (row = 1) Then
                        UpExist = 1
                    Else
                        If ((bricksOn(row - 1, col) > 1 And bricksOn(row - 1, col) <> 19) Or (bricksOn(row - 1, col) = 19 And bricks(row - 1, col).Visible = True)) Then
                            UpExist = 1
                        End If
                    End If
                    If (row < Rows) Then
                        If ((bricksOn(row + 1, col) > 1 And bricksOn(row + 1, col) <> 19) Or (bricksOn(row + 1, col) = 19 And bricks(row + 1, col).Visible = True)) Then
                            DownExist = 1
                        End If
                    End If
                    
                    If (UpExist = 0 And BDY = 2) Then
                        If (b3 < b1 + 10 * Screen.TwipsPerPixelY And ((a3 < a1 And a4 > a1) Or (a3 > a1 And a4 < a2) Or (a3 < a2 And a4 > a2))) Then
                            BounceY = 1
                            BallDirectionY(bb) = 1
                        End If
                    End If
                    
                    If (DownExist = 0 And BDY = 1) Then
                        If (b3 > b2 - 10 * Screen.TwipsPerPixelY And ((a3 < a1 And a4 > a1) Or (a3 > a1 And a4 < a2) Or (a3 < a2 And a4 > a2))) Then
                            BounceY = 1
                            BallDirectionY(bb) = 2
                        End If
                    End If
                    
                    ' Bounce on X axis
                    LeftExist = 0
                    RightExist = 0
                    If (col = 1) Then
                        LeftExist = 1
                    Else
                        If ((bricksOn(row, col - 1) > 1 And bricksOn(row, col - 1) <> 19) Or (bricksOn(row, col - 1) = 19 And bricks(row, col - 1).Visible = True)) Then
                            LeftExist = 1
                        End If
                    End If
                    If (col = Cols) Then
                        RightExist = 1
                    Else
                        If ((bricksOn(row, col + 1) > 1 And bricksOn(row, col + 1) <> 19) Or (bricksOn(row, col + 1) = 19 And bricks(row, col + 1).Visible = True)) Then
                            RightExist = 1
                        End If
                    End If
                    
                    If (LeftExist = 0 And BDX = 2) Then
                        If (a3 < a1 + 13 * Screen.TwipsPerPixelX And ((b3 - 10 * Screen.TwipsPerPixelY < b1 And b4 > b1) Or (b3 > b1 And b4 < b2) Or (b3 + 10 * Screen.TwipsPerPixelY < b2 And b4 > b2))) Then
                            BounceX = 1
                            BallDirectionX(bb) = 1
                        End If
                    End If
                    
                    If (RightExist = 0 And BDX = 1) Then
                        If (a3 > a2 - 13 * Screen.TwipsPerPixelX And ((b3 - 10 * Screen.TwipsPerPixelY < b1 And b4 > b1) Or (b3 > b1 And b4 < b2) Or (b3 + 10 * Screen.TwipsPerPixelY < b2 And b4 > b2))) Then
                            BounceX = 1
                            BallDirectionX(bb) = 2
                        End If
                    End If
                    
                    If (BounceX = 1) Then
                        BX = BX + 1
                    End If
                    If (BounceY = 1) Then
                        BY = BY + 1
                    End If
                    
                    'Form1.Caption = "X:" & BounceX & " - Y:" & BounceY & " | BX:" & BX & " - BY:" & BY
                End If
                DoEvents
            End If
        End If
    End If
End Sub

Public Sub CheckHitPaddle()
    Dim BallX2, BallY2 As Double
    Dim x1, x2, x3 As Long
    Dim iball As Long
    
    For iball = 0 To 4
        If (Ball(iball).Visible = True) Then
            x1 = Paddle.Left
            x2 = Paddle.Left + Paddle.Width
            x3 = Ball(iball).Left
            
            BallX2 = BallX(iball) + Ball(iball).Width
            BallY2 = BallY(iball) + Ball(iball).Height
            
            If ((BallX(iball) >= Paddle.Left And BallX(iball) <= Paddle.Left + Paddle.Width) Or (BallX2 >= Paddle.Left And BallX2 <= Paddle.Left + Paddle.Width)) Then
                If ((BallY(iball) >= Paddle.Top And BallY(iball) <= Paddle.Top + Paddle.Height) Or (BallY2 >= Paddle.Top And BallY2 <= Paddle.Top + Paddle.Height)) Then
                    Call PlaySound(app_path() & "Sounds/sound1.wav", 0, SND_ASYNC Or SND_NOSTOP)
                    If (RollAction <> 21) Then
                        Call CalculateBallAngle(iball)
                        BallDirectionY(iball) = 1
                    Else
                        BallReleased(iball) = 0
                        BallDirectionY(iball) = 1
                    End If
                End If
            End If
        End If
    Next iball
End Sub

Public Sub CheckLostBall()
    Dim Ended As Long
    Dim iball As Long
    
    Ended = 1
    For iball = 0 To 4
        If (Ball(iball).Visible = True) Then
            If (BallY(iball) > Paddle.Top + Paddle.Height) Then
                Ball(iball).Visible = False
            End If
            
            If (Ball(iball).Visible = True) Then
                Ended = 0
            End If
        End If
    Next iball
    
    If (Ended = 1) Then
        Lives = Lives - 1
        UpdatePanel
        InLevel = 0
    End If
End Sub

Public Sub CheckLevelFinish()
    Dim i, j, Finish As Integer
    
    Finish = 1
    For i = 1 To Rows
        For j = 1 To Cols
            If (bricksOn(i, j) > 0 And bricksOn(i, j) <> 18 And bricksOn(i, j) <> 19) Then
                Finish = 0
                Exit For
            End If
        Next j
    Next i
    
    LevelFinished = Finish
End Sub

Sub ExplodeBomb(ByVal row As Long, ByVal col As Long)
    Dim r1, c1, IsHided As Long
    Dim iball As Long
    
    bricksOn(row, col) = 0
    bricks(row, col).Visible = False
    
    'Next row?
    For r1 = row - 1 To row + 1
        For c1 = col - 1 To col + 1
            If (r1 <= Rows And r1 > 0 And c1 <= Cols And c1 > 0) Then
                If (bricksOn(r1, c1) > 0) Then
                    If (bricksOn(r1, c1) = 9) Then
                        Call ExplodeBomb(r1, c1)
                    ElseIf (bricksOn(r1, c1) = 10) Then 'Extra Life
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        Lives = Lives + 1
                        UpdatePanel
                    ElseIf (bricksOn(r1, c1) = 11) Then 'End of Level!
                        AddPlayerPoints (bricksOn(r1, c1))
                        EndLevel
                    ElseIf (bricksOn(r1, c1) = 12) Then  'Big Paddle!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        Paddle.Width = 2000
                        BigTimer.Interval = 10000
                    ElseIf (bricksOn(r1, c1) = 13) Then 'Slow down!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        For iball = 0 To 4
                            dbally(iball) = dbally(iball) - 10
                            If dbally(iball) < LevelMinSpeed Then
                                dbally(iball) = LevelMinSpeed
                            End If
                        Next iball
                        UpdatePanel
                    ElseIf (bricksOn(r1, c1) = 14) Then 'Reverse!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        ReverseControls = 1
                        TimerReverse.Interval = 10000
                    ElseIf (bricksOn(r1, c1) = 15) Then 'Hide bricks!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        Hidebricks
                        TimerHide.Interval = 10000
                    ElseIf (bricksOn(r1, c1) = 16) Then 'Decrease Values!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        DecreaseValues
                    ElseIf (bricksOn(r1, c1) = 17) Then 'Increase Values!
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                        IncreaseValues
                    ElseIf (bricksOn(r1, c1) = 18) Then 'Non Beatable block!
                        bricksOn(r1, c1) = bricksOn(row, col)
                    ElseIf (bricksOn(row, col) = 19) Then 'Shortly hided block!
                        IsHided = CheckIfHided(r1, c1)
                        If (IsHided = 0) Then
                            bricksOn(r1, c1) = bricksOn(r1, c1)
                            bricks(r1, c1).Picture = LoadPicture("./Images/Brick0.bmp")
                            bricks(r1, c1).Visible = False
                            DoEvents
                            bricksHide(r1, c1) = Now
                        End If
                    ElseIf (bricksOn(r1, c1) = 101 Or bricksOn(r1, c1) = 102 Or bricksOn(r1, c1) = 103 Or bricksOn(r1, c1) = 104 Or bricksOn(r1, c1) = 105 Or bricksOn(r1, c1) = 106 Or bricksOn(r1, c1) = 107 Or bricksOn(r1, c1) = 108 Or bricksOn(r1, c1) = 109 Or bricksOn(r1, c1) = 1110) Then 'Simple coloured brick
                        bricksOn(r1, c1) = 0
                        AddPlayerPoints (bricksOn(r1, c1))
                    Else
                        bricksOn(r1, c1) = bricksOn(r1, c1) - 1
                        AddPlayerPoints (bricksOn(r1, c1))
                        If (bricksOn(r1, c1) = 0) Then
                            bricks(r1, c1).Visible = False
                        Else
                            bricks(r1, c1).Picture = LoadPicture("./Images/Brick" & bricksOn(r1, c1) & ".bmp")
                        End If
                    End If
                End If
            End If
        Next c1
    Next r1
        
    UpdatePanel
End Sub

Public Sub EndLevel()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricksOn(i, j) = 0
        Next j
    Next i
End Sub

Public Sub Hidebricks()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricks(i, j).Visible = False
        Next j
    Next i
End Sub

Public Sub Unhidebricks()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            If (bricksOn(i, j) > 0) Then
                bricks(i, j).Visible = True
            End If
        Next j
    Next i
End Sub

Public Sub DecreaseValues()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            If (bricksOn(i, j) > 0) Then
                If (bricksOn(i, j) <> 18 And bricksOn(i, j) <> 19) Then
                    AddPlayerPoints (bricksOn(i, j))
                    UpdatePanel
                End If
                If (bricksOn(i, j) < 8) Then
                    bricksOn(i, j) = bricksOn(i, j) - 1
                End If
                If (bricksOn(i, j) = 0) Then
                    bricks(i, j).Visible = False
                Else
                    bricks(i, j).Picture = LoadPicture("./Images/Brick" & bricksOn(i, j) & ".bmp")
                End If
            End If
        Next j
    Next i
End Sub

Public Sub IncreaseValues()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            If (bricksOn(i, j) > 0) Then
                If (bricksOn(i, j) < 7) Then
                    bricksOn(i, j) = bricksOn(i, j) + 1
                End If
                If (bricksOn(i, j) = 0) Then
                    bricks(i, j).Visible = False
                Else
                    bricks(i, j).Picture = LoadPicture("./Images/Brick" & bricksOn(i, j) & ".bmp")
                End If
            End If
        Next j
    Next i
End Sub

Public Function ImportLevel(Level) As Long
    On Error GoTo ErrorHandle
    
    Dim i, j, MaxCols, CurCols, BrickType As Long
    Dim fso As New FileSystemObject, fil As File, ts As TextStream
    Dim FileName, rowline As String
    Dim ReadRows As Long
    Dim BrickTypes() As String
    
    FileName = ".\Levels\Level" & Level & ".dat"
    ReadRows = 0
    MaxCols = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fso.OpenTextFile(FileName, ForReading, True)
    rowline = "Start"
    While (Trim(rowline) <> "")
        If (ts.AtEndOfLine = True) Then
            rowline = ""
        Else
            rowline = Trim(ts.ReadLine)
            If (Trim(rowline) <> "") Then
                ReadRows = ReadRows + 1
                BrickTypes = Split(rowline, ",")
                CurCols = 1 + UBound(BrickTypes, 1)
                If (CurCols > MaxCols) Then
                    MaxCols = CurCols
                End If
            End If
        End If
    Wend
    ts.Close
        
    If (ReadRows > 0 And MaxCols > 0) Then
        ReDim bricks(ReadRows + 1, MaxCols + 1)
        ReDim bricksOn(ReadRows + 1, MaxCols + 1)
        ReDim bricksHide(ReadRows + 1, MaxCols + 1)
        
        Rows = ReadRows
        Cols = MaxCols
        
        For i = 1 To ReadRows
            For j = 1 To MaxCols
                bricksOn(i, j) = 0
                bricksHide(i, j) = 0
                Set bricks(i, j) = Controls.Add("vb.image", "Image" & i & "_" & j)
            Next j
        Next i
        
        ReadRows = 0
        Set ts = fso.OpenTextFile(FileName, ForReading, True)
        rowline = "Start"
        While (Trim(rowline) <> "")
            If (ts.AtEndOfLine = True) Then
                rowline = ""
            Else
                rowline = Trim(ts.ReadLine)
                If (Trim(rowline) <> "") Then
                    ReadRows = ReadRows + 1
                    BrickTypes = Split(rowline, ",")
                    For i = 1 To 1 + UBound(BrickTypes, 1)
                        BrickType = Val(BrickTypes(i - 1))
                        If (BrickType < 0) Then
                            BrickType = 0
                        End If
                        If (BrickType > 19 And (BrickType <= 100 Or BrickType > 110)) Then
                            BrickType = 19
                        End If
                        
                        bricksOn(ReadRows, i) = BrickType
                        bricksHide(ReadRows, i) = DateAdd("yyyy", -1, Now)
                    Next
                End If
            End If
        Wend
        ts.Close
        
        For i = 1 To ReadRows
            For j = 1 To MaxCols
                bricks(i, j).Top = i * BrickHeight
                bricks(i, j).Left = 10 + (j - 1) * BrickWidth
                bricks(i, j).Width = BrickWidth
                bricks(i, j).Height = BrickHeight
                If (bricksOn(i, j) > 0) Then
                    bricks(i, j).Picture = LoadPicture("./Images/Brick" & bricksOn(i, j) & ".bmp")
                End If
                bricks(i, j).Stretch = True
                If (bricksOn(i, j) > 0) Then
                    bricks(i, j).Visible = True
                Else
                    bricks(i, j).Visible = False
                End If
            Next j
        Next i
    End If

    If (ReadRows < 1 Or MaxCols < 1) Then
        ImportLevel = -1
    Else
        ImportLevel = 1
        Form1.Width = (Cols + 1) * BrickWidth - FormErrorX
        Form1.Height = Screen.TwipsPerPixelY * 500
        Paddle.Top = Form1.Height - 1400
        
        Paddle.Left = Form1.Width / 2 - Paddle.Width / 2
        LabelLives.Top = Form1.Height - 1000
        LabelSpeed.Top = Form1.Height - 1000
        LabelPoints.Top = Form1.Height - 1000
        LabelLevel.Top = Form1.Height - 1000
        
        LabelLives.Left = 10
        LabelLevel.Left = LabelLives.Left + LabelLives.Width + 100
        LabelPoints.Left = LabelLevel.Left + LabelLevel.Width + 100
        LabelSpeed.Left = LabelPoints.Left + LabelPoints.Width + 100
        
        Call CenterForm(Form1)
        Form1.Refresh
    End If
    Exit Function
    
ErrorHandle:
    ImportLevel = -1
    MsgBox "Level failed to load!"
    MsgBox Err.Description
End Function

Public Function CheckIfHided(ByVal row As Long, ByVal col As Long) As Long
    Dim diff As Long
    
    If (bricks(row, col).Visible = False) Then
        CheckIfHided = 1
    Else
        CheckIfHided = 0
    End If
End Function

Public Sub CreateNewBrickRoll(ByVal row As Long, ByVal col As Long)
    On Error Resume Next
    Dim rnd1 As Long
    rnd1 = 20 + 10 * Rnd(100)
    If (rnd1 < 21) Then
        rnd1 = 21
    End If
    If (rnd1 > 29) Then
        rnd1 = 29
    End If
    
    If (RollUsed < 100) Then
        Set bricksRoll(RollUsed + 1) = Controls.Add("vb.image", CreateRandomName)
        bricksRollOn(RollUsed + 1) = rnd1
        
        bricksRoll(RollUsed + 1).Top = bricks(row, col).Top
        bricksRoll(RollUsed + 1).Left = bricks(row, col).Left
        bricksRoll(RollUsed + 1).Width = BrickWidth
        bricksRoll(RollUsed + 1).Height = BrickHeight
        bricksRoll(RollUsed + 1).Picture = LoadPicture("./Images/Brick" & rnd1 & ".bmp")
        bricksRoll(RollUsed + 1).Stretch = True
        bricksRoll(RollUsed + 1).Visible = True
        
        RollUsed = RollUsed + 1
    End If
    UpdatePanel
End Sub

Public Sub DestroyBrickRoll(ByVal Index As Long)
    On Error Resume Next
    Dim i As Long
    
    bricksRoll(Index).Visible = False
    Controls.Remove (bricksRoll(Index).Name)
    For i = Index To RollUsed - 1
        Set bricksRoll(i) = bricksRoll(i + 1)
        bricksRollOn(i) = bricksRollOn(i + 1)
    Next i
    RollUsed = RollUsed - 1
    UpdatePanel
End Sub

Public Sub MoveBrickRoll(ByVal Index As Long)
    On Error Resume Next
    bricksRoll(Index).Top = bricksRoll(Index).Top + dbally(0)
    DoEvents
End Sub

Public Function CreateRandomName() As String
    Dim rnd1, i As Long
    
    CreateRandomName = ""
    For i = 1 To 30
        rnd1 = 1 + 23 * Rnd(100)
        CreateRandomName = CreateRandomName & Chr(65 + rnd1)
    Next i
End Function

Public Sub CheckRollHitPaddle()
    On Error Resume Next
    Dim RollX, RollY, RollX2, RollY2 As Double
    Dim i, x1, x2, x3 As Long
    Dim iball As Long
    
    For i = 1 To RollUsed
        RollX = bricksRoll(i).Left
        RollY = bricksRoll(i).Top
        x1 = Paddle.Left
        x2 = Paddle.Left + Paddle.Width
        x3 = bricksRoll(i).Left
        
        RollX2 = RollX + bricksRoll(i).Width
        RollY2 = RollY + bricksRoll(i).Height
        
        If ((RollX >= Paddle.Left And RollX <= Paddle.Left + Paddle.Width) Or (RollX2 >= Paddle.Left And RollX2 <= Paddle.Left + Paddle.Width)) Then
            If ((RollY >= Paddle.Top And RollY <= Paddle.Top + Paddle.Height) Or (RollY2 >= Paddle.Top And RollY2 <= Paddle.Top + Paddle.Height)) Then
                ResetRollActions
                RollAction = bricksRollOn(i)
                If (RollAction = 22) Then 'Release 2nd ball
                    For iball = 0 To 1
                        If (Ball(iball).Visible = False) Then
                            Ball(iball).Visible = True
                            Ball(iball).Left = Paddle.Left + Paddle.Width / 2 - 3 * Paddle.Width / 8
                            Ball(iball).Top = Paddle.Top - 200
                            BallReleased(iball) = 1
                            BallAngle(iball) = 45
                            BallX(iball) = Ball(iball).Left
                            BallY(iball) = Ball(iball).Top
                            dbally(iball) = 20
                            BallDirectionX(iball) = 1
                            BallDirectionY(iball) = 1
                            BallType(iball) = 1
                            Ball(iball).Picture = LoadPicture(app_path() & "/Images/ball.bmp")
                        End If
                    Next iball
                End If
                If (RollAction = 23) Then 'Increase Paddle
                    If (Paddle.Width < 2000) Then
                        Paddle.Width = 2000
                    End If
                End If
                If (RollAction = 24) Then 'Super ball
                    For iball = 0 To 1
                        BallType(iball) = 2
                        Ball(iball).Picture = LoadPicture(app_path() & "/Images/ball2.bmp")
                    Next iball
                End If
                If (RollAction = 25) Then 'Slow down
                    For iball = 0 To 4
                        dbally(iball) = dbally(iball) - 10
                        If (dbally(iball) < 10) Then
                            dbally(iball) = 10
                        End If
                    Next iball
                End If
                If (RollAction = 26) Then 'End of Level
                    EndLevel
                End If
                If (RollAction = 27) Then 'Extra Life
                    Lives = Lives + 1
                    UpdatePanel
                End If
                If (RollAction = 28) Then 'Reverse Controls
                    ReverseControls = 1
                    TimerReverse.Interval = 10000
                End If
                If (RollAction = 29) Then 'Hide bricks
                    Hidebricks
                    TimerHide.Interval = 10000
                End If
                DestroyBrickRoll (i)
            End If
        End If
    Next i
End Sub

Public Sub ResetRollActions()
    Dim iball As Long
    
    If (Paddle.Width > 1000) Then
        Paddle.Width = 1000
    End If
    For iball = 0 To 1
        BallType(iball) = 1
        Ball(iball).Picture = LoadPicture(app_path() & "/Images/ball.bmp")
    Next iball
End Sub

Public Sub ClearRolls(ByVal All As Long)
    On Error Resume Next
    Dim i As Long
    
    For i = 1 To RollUsed
        If (All = 1) Then
            DestroyBrickRoll (i)
        ElseIf (bricksRoll(i).Top > Paddle.Top + Paddle.Height) Then
            DestroyBrickRoll (i)
        End If
    Next i
End Sub

Public Sub AddPlayerPoints(ByVal how As Long)
    PlayerPoints = PlayerPoints + how * (Level / 10) + (Level - 1)
    UpdatePanel
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iball As Long
    
    Select Case KeyCode
        Case vbKeyA
            For iball = 0 To 4
                If Ball(iball).Visible = True Then
                    BallAngle(iball) = 45
                End If
            Next iball
        Case vbKeyB
        Case vbKeyC
        Case vbKeyD
        Case vbKeyE
        Case vbKeyF
        Case vbKeyG
        Case vbKeyH
        Case vbKeyI
        Case vbKeyJ
        Case vbKeyK
        Case vbKeyL
            EndLevel
        Case vbKeyM
        Case vbKeyN
        Case vbKeyO
        Case vbKeyP
            If (LevelPaused = 0) Then
                LevelPaused = 1
            Else
                LevelPaused = 0
            End If
        Case vbKeyQ
        Case vbKeyR
        Case vbKeyS
        Case vbKeyT
        Case vbKeyU
        Case vbKeyV
            Lives = Lives + 1
        Case vbKeyW
        Case vbKeyX
        Case vbKeyY
        Case vbKeyZ
    End Select
End Sub

