VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBnoid Level Editor"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleMode       =   0  'User
   ScaleWidth      =   6514.032
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "Clear Level"
      End
      Begin VB.Menu mnuFileGlobal 
         Caption         =   "Global Change"
      End
      Begin VB.Menu mnuFileLevelSize 
         Caption         =   "Set Level Size"
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

Private Type POINTAPI
        x As Long
        Y As Long
End Type
Dim a As POINTAPI

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYCAPTION = 4


Dim Rows As Integer
Dim Cols As Integer
Dim NumTools As Integer
Dim SelectedTool As Integer
Dim brickWidth As Integer
Dim brickHeight As Integer
Dim toolbox() As Image
Dim bricks() As Image
Dim bricksOn() As Integer
Dim FormErrorX As Long
Dim ApplicationClosing As Long
Dim CdlgEx1 As New CdlgEx

Function File_Exists(ByVal PathName As String, Optional Directory As Boolean) As Boolean
    'Returns True if the passed pathname exist
    'Otherwise returns False

    If PathName <> "" Then
        If IsMissing(Directory) Or Directory = False Then
            File_Exists = (Dir$(PathName) <> "")
        Else
            File_Exists = (Dir$(PathName, vbDirectory) <> "")
        End If
    End If
End Function

Public Function app_path() As String
    Dim x As String
    x = App.path
    If Right$(x, 1) <> "\" Then x = x + "\"
    app_path = UCase$(x)
End Function

Public Sub CenterForm(Frm As Form)
    Dim Left, Top As Integer

    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - Frm.Width / 2
    Top = Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) + GetSystemMetrics(SM_CYCAPTION)) / 2 - Frm.Height / 2

    Frm.Move Left, Top
End Sub

Private Sub Form_Activate()
    ApplicationClosing = 0
End Sub

Private Sub Form_Load()
    FormErrorX = 270
    NumTools = 111
    SelectedTool = -1
    Rows = 20
    Cols = 13
    brickWidth = 480
    brickHeight = 240
    DrawDummybricks
    DrawToolBox
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim x1, x2, y1, y2, i, j As Long
    
    For i = 0 To NumTools - 1
        If (i <= 19 Or (i > 100 And i < 111)) Then
            x1 = toolbox(i).Left
            x2 = x1 + toolbox(i).Width
            y1 = toolbox(i).Top
            y2 = y1 + toolbox(i).Height
            
            If (x >= x1 And x <= x2 And Y >= y1 And Y <= y2) Then
                If (SelectedTool >= 0) Then
                    toolbox(SelectedTool).BorderStyle = vbBSNone
                End If
                
                SelectedTool = i
                toolbox(i).BorderStyle = vbFixedSingle
                Exit For
            End If
        End If
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            x1 = bricks(i, j).Left
            x2 = x1 + bricks(i, j).Width
            y1 = bricks(i, j).Top
            y2 = y1 + bricks(i, j).Height
            
            If (x >= x1 And x <= x2 And Y >= y1 And Y <= y2 And SelectedTool >= 0) Then
                If (mnuFileGlobal.Checked = True) Then
                    Call GlobalReplace(bricksOn(i, j), SelectedTool)
                Else
                    bricksOn(i, j) = SelectedTool
                    bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick" & bricksOn(i, j) & ".bmp")
                End If
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ApplicationClosing = 1
End Sub

Public Sub UpdatePanel()
    DoEvents
End Sub

Public Sub RemovebrickControls()
    On Error Resume Next
    Dim i, j As Long
    
    For i = 1 To Rows
        For j = 1 To Cols
            Controls.Remove ("Image" & i & "_" & j)
        Next j
    Next i
End Sub

Public Sub DrawDummybricks()
    Dim i, j, rand1 As Long
    
    ReDim bricks(Rows + 1, Cols + 1)
    ReDim bricksOn(Rows + 1, Cols + 1)
    
    For i = 1 To Rows
        For j = 1 To Cols
            Set bricks(i, j) = Controls.Add("vb.image", "Image" & i & "_" & j)
            bricksOn(i, j) = 1
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricks(i, j).Top = i * brickHeight
            bricks(i, j).Left = 10 + (j - 1) * brickWidth
            bricks(i, j).Width = brickWidth
            bricks(i, j).Height = brickHeight
            bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick" & bricksOn(i, j) & ".bmp")
            bricks(i, j).Stretch = True
            bricks(i, j).Visible = True
            bricks(i, j).Enabled = False
        Next j
    Next i
    
    Form1.Width = (Cols + 1) * brickWidth - FormErrorX
    Form1.Height = Screen.TwipsPerPixelY * 500
    
    Call CenterForm(Form1)
End Sub

Public Sub DrawToolBox()
    Dim i, j, tooly As Long
    
    tooly = 1800
    ReDim toolbox(NumTools + 1)
    
    For i = 0 To NumTools - 1
        If (i <= 19 Or (i > 100 And i < 111)) Then
            Set toolbox(i) = Controls.Add("vb.image", "Tool" & i)
        End If
    Next i
    
    For i = 0 To NumTools - 1
        If (i <= 19 Or (i > 100 And i < 111)) Then
            If (i < 10) Then
                toolbox(i).Top = Form1.Height - tooly
                toolbox(i).Left = 10 + (i + 1) * brickWidth
            ElseIf (i < 20) Then
                toolbox(i).Top = Form1.Height - tooly + brickHeight
                toolbox(i).Left = 10 + (i - 10 + 1) * brickWidth
            ElseIf (i < 111) Then
                toolbox(i).Top = Form1.Height - tooly + 2 * brickHeight
                toolbox(i).Left = 10 + (i - 101 + 1) * brickWidth
            End If
        
            toolbox(i).Width = brickWidth
            toolbox(i).Height = brickHeight
            If (i > 0) Then
                toolbox(i).Picture = LoadPicture(app_path() & "/Images/brick" & i & ".bmp")
            Else
                toolbox(i).Picture = LoadPicture(app_path() & "/Images/brick00.bmp")
            End If
            toolbox(i).Stretch = True
            toolbox(i).Visible = True
            toolbox(i).Enabled = False
        End If
    Next i
    
    Call CenterForm(Form1)
End Sub


Public Function ImportLevel(FileName As String) As Long
    On Error GoTo ErrorHandle
    
    Dim i, j, MaxCols, CurCols, brickType As Long
    Dim fso As New FileSystemObject, fil As File, ts As TextStream
    Dim rowline As String
    Dim ReadRows As Long
    Dim brickTypes() As String
    
    'FileName = ".\Levels\Level" & Level & ".dat"
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
                brickTypes = Split(rowline, ",")
                CurCols = 1 + UBound(brickTypes, 1)
                If (CurCols > MaxCols) Then
                    MaxCols = CurCols
                End If
            End If
        End If
    Wend
    ts.Close
        
    If (ReadRows > 0 And MaxCols > 0) Then
        Call ResizeLevel(ReadRows, MaxCols)
        
        RemovebrickControls
        ReDim bricks(Rows + 1, Cols + 1)
        ReDim bricksOn(Rows + 1, Cols + 1)
        
        For i = 1 To Rows
            For j = 1 To Cols
                bricksOn(i, j) = 0
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
                    brickTypes = Split(rowline, ",")
                    For i = 1 To 1 + UBound(brickTypes, 1)
                        brickType = Val(brickTypes(i - 1))
                        If (brickType < 0) Then
                            brickType = 0
                        End If
                        If (brickType > 19 And (brickType < 101 Or brickType > 110)) Then
                            brickType = 19
                        End If
                        
                        bricksOn(ReadRows, i) = brickType
                    Next
                End If
            End If
        Wend
        ts.Close
        
        For i = 1 To Rows
            For j = 1 To Cols
                bricks(i, j).Top = i * brickHeight
                bricks(i, j).Left = 10 + (j - 1) * brickWidth
                bricks(i, j).Width = brickWidth
                bricks(i, j).Height = brickHeight
                bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick" & bricksOn(i, j) & ".bmp")
                bricks(i, j).Stretch = True
                bricks(i, j).Visible = True
                bricks(i, j).Enabled = False
            Next j
        Next i
    End If

    If (ReadRows < 1 Or MaxCols < 1) Then
        ImportLevel = -1
    Else
        ImportLevel = 1
        Form1.Width = (Cols + 1) * brickWidth - FormErrorX
        Form1.Height = Screen.TwipsPerPixelY * 500
    
        Call CenterForm(Form1)
    End If
    Exit Function
    
ErrorHandle:
    ImportLevel = -1
    MsgBox "Level failed to load!"
End Function

Private Sub mnuFileClear_Click()
    Dim i, j As Integer
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricksOn(i, j) = 0
            bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick0.bmp")
        Next j
    Next i
End Sub

Private Sub mnuFileGlobal_Click()
    If (mnuFileGlobal.Checked = False) Then
        mnuFileGlobal.Checked = True
    Else
        mnuFileGlobal.Checked = False
    End If
End Sub

Private Sub mnuFileLevelSize_Click()
    Dim LevelX, LevelY As Long
    
    Form2.Show vbModal
    LevelX = Form2.LevelX
    LevelY = Form2.LevelY
    If (LevelX > 0 And LevelY > 0) Then
        Call ResizeLevel(LevelY, LevelX)
    End If
    Set Form2 = Nothing
End Sub

Private Sub mnuFileLoad_Click()
    CdlgEx1.CancelError = False
    CdlgEx1.Filter = "Level files|*.dat|All files|*.*"
    CdlgEx1.ShowOpen
    If (Trim(CdlgEx1.FileName) <> "") Then
        ImportLevel (CdlgEx1.FileName)
    End If
End Sub

Private Sub mnuFileSave_Click()
    Dim save As Boolean
    Dim ret As Long
    
    save = True
    CdlgEx1.CancelError = False
    CdlgEx1.Filter = "Level files|*.dat|All files|*.*"
    CdlgEx1.ShowSave
    If (Trim(CdlgEx1.FileName) <> "") Then
        If (File_Exists(CdlgEx1.FileName)) Then
            ret = MsgBox("Level already exists. Overwrite it?", vbOKCancel)
            If (ret = vbOK) Then
                save = True
            Else
                save = False
            End If
        End If
        If (save = True) Then
            ExportLevel (CdlgEx1.FileName)
        End If
    End If
End Sub

Public Function ExportLevel(FileName As String) As Long
    On Error GoTo Error1
    
    Dim i, j As Long
    Dim fso As New FileSystemObject, fil As File, ts As TextStream
    Dim linestr As String
    
    ExportLevel = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(FileName, ForWriting, True)
    
    For i = 1 To Rows
        linestr = ""
        For j = 1 To Cols
            If (j < Cols) Then
                linestr = linestr & bricksOn(i, j) & ","
            Else
                linestr = linestr & bricksOn(i, j)
            End If
        Next j
        ts.WriteLine (linestr)
    Next i
    ts.Close
    MsgBox "Level saved succesfully!"
    Exit Function
    
Error1:
    MsgBox "Level failed to be saved!"
    ExportLevel = -1
End Function


Public Sub GlobalReplace(ByVal Oldbrick As Long, ByVal NewBrake As Long)
    Dim i, j As Long
    
    For i = 1 To Rows
        For j = 1 To Cols
            If (bricksOn(i, j) = Oldbrick) Then
                bricksOn(i, j) = NewBrake
                bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick" & bricksOn(i, j) & ".bmp")
            End If
        Next j
    Next i
End Sub

Public Sub ResizeLevel(ByVal LevelY As Long, ByVal LevelX As Long)
    Dim oldbricksOn() As Integer
    Dim i, j As Integer
    Dim OldRows, OldCols As Long
    
    ReDim oldbricksOn(Rows + 1, Cols + 1)
    For i = 1 To Rows
        For j = 1 To Cols
            oldbricksOn(i, j) = bricksOn(i, j)
            bricks(i, j).Visible = False
        Next j
    Next i
    
    OldRows = Rows
    OldCols = Cols
    
    Rows = LevelY
    Cols = LevelX
    
    RemovebrickControls
    ReDim bricks(Rows + 1, Cols + 1)
    ReDim bricksOn(Rows + 1, Cols + 1)
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricksOn(i, j) = 0
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            If (i < OldRows And j < OldCols) Then
                bricksOn(i, j) = oldbricksOn(i, j)
            End If
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            Set bricks(i, j) = Controls.Add("vb.image", "Image" & i & "_" & j)
        Next j
    Next i
    
    For i = 1 To Rows
        For j = 1 To Cols
            bricks(i, j).Top = i * brickHeight
            bricks(i, j).Left = 10 + (j - 1) * brickWidth
            bricks(i, j).Width = brickWidth
            bricks(i, j).Height = brickHeight
            bricks(i, j).Picture = LoadPicture(app_path() & "/Images/brick" & bricksOn(i, j) & ".bmp")
            bricks(i, j).Stretch = True
            bricks(i, j).Visible = True
            bricks(i, j).Enabled = False
        Next j
    Next i
    
    Form1.Width = (Cols + 1) * brickWidth - FormErrorX
    Form1.Height = Screen.TwipsPerPixelY * 500

    Call CenterForm(Form1)
End Sub
