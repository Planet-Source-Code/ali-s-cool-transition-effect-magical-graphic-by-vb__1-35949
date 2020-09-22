VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Transition Effects"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   Icon            =   "frmTransEffects.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Wipe Horizontal"
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Wipe Veritcal"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Right"
      Height          =   495
      Index           =   3
      Left            =   5640
      TabIndex        =   12
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Left"
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Down"
      Height          =   495
      Index           =   1
      Left            =   4440
      TabIndex        =   10
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Wipe Up"
      Height          =   495
      Index           =   0
      Left            =   4440
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stretching Move"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Slide -Down"
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Lines Horizontal"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stretching Push"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Slide -Up"
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   5520
      Picture         =   "frmTransEffects.frx":0442
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   5520
      Picture         =   "frmTransEffects.frx":429E
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Lines Vertical"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5160
      Left            =   360
      Picture         =   "frmTransEffects.frx":922B
      ScaleHeight     =   5100
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   480
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Transition Effects By Mohammed Ali Sohrabi ,ali6236@yahoo.com
'   Cool Transition for your programs!
'   See Notes on the module
Public StopProgram As Boolean
Private Sub Command1_Click(Index As Integer)
'Random Lines
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical - Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Enable
'   Step             : Disable
'**********************************
'   Notes:
'   RefreshRate : number of lines in each refresh
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        RandomLines Picture1, Picture2, VerticalSide, 0
    Else
        'the speed is 1, but it is slow, we use RefreshRate for faster result...
        RandomLines Picture1, Picture2, HorizontalSide, 2
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command2_Click(Index As Integer)
'Slide
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : Yes
'   Sides            : All Sides (Up and Down are completed)
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes: Just use Up and Down,
'          I will complete other sides as soon as possible!
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        Slide Picture1, Picture3, Picture2, aUp, 3
    Else
        Slide Picture1, Picture3, Picture2, aDown, 3
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command3_Click(Index As Integer)
'Stretching
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : Yes
'   Sides            : Left - Right
'**********************************
'   Push Modes       : Enable
'   Refresh Rate     : Enable
'   Step             : Enable
'**********************************
'   Notes:
'   Stretch is a slow effect,
'   Just use for small pictures, with large steps
'   and use push mode just when you need,
If Not IsReady Then Exit Sub
    lngSpeed = 1
    If Index = 0 Then
        Stretching Picture1, Picture3, Picture2, sLeft, 5, 0, Pushing
    Else
        Stretching Picture1, Picture3, Picture2, sLeft, 5, 0, Moving
    End If
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command4_Click(Index As Integer)
'Wipe
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : All (Left,Right,Up,Down)
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
If Not IsReady Then Exit Sub
    lngSpeed = 1
    Wipe Picture1, Picture2, 2 ^ Index, 3 '!!!! i'm using ^ !!! it's better to use select case
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Command5_Click(Index As Integer)
'Wipe In
'**********************************
'   Need New Picture : Yes
'   Need Old Picture : No
'   Sides            : Vertical and Horizontal
'**********************************
'   Push Modes       : Disable
'   Refresh Rate     : Disable
'   Step             : Enable
'**********************************
'   Notes:
'   This is like two normal wipe.

If Not IsReady Then Exit Sub
    lngSpeed = 1
    Wipe_In Picture1, Picture2, Index + 1, 3  '!!!! i'm using ^ !!! it's better to use select case
    If StopProgram Then Exit Sub
    Set Picture2.Picture = Picture3.Picture
    Set Picture3.Picture = Picture1.Picture
End Sub

Private Sub Form_Load()
    StopProgram = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnRunning = False
    StopProgram = True
    Unload Me
End Sub

