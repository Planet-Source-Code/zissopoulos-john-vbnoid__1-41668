VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Level Size"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3105
      TabIndex        =   5
      Top             =   315
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2295
      TabIndex        =   4
      Top             =   315
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1215
      TabIndex        =   1
      Top             =   495
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1215
      TabIndex        =   0
      Top             =   90
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Columns:"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Rows:"
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   1005
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LevelX, LevelY As Long


Private Sub Command1_Click()
    LevelY = Val(Text1.Text)
    LevelX = Val(Text2.Text)
    Me.Visible = False
End Sub

Private Sub Form_Load()
    LevelY = 0
    LevelX = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LevelY = 0
    LevelX = 0
End Sub
