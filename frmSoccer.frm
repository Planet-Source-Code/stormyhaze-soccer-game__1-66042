VERSION 5.00
Begin VB.Form frmSoccer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soccer"
   ClientHeight    =   11295
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSoccer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11295
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9750
      Top             =   11505
   End
   Begin VB.Timer tmrOpponent 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10230
      Top             =   11505
   End
   Begin VB.Timer tmrShoot 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10710
      Top             =   11505
   End
   Begin VB.Timer tmrBall 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11190
      Top             =   11505
   End
   Begin VB.ListBox lstHighscores 
      Enabled         =   0   'False
      Height          =   10000
      IntegralHeight  =   0   'False
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdShoot 
      Caption         =   "3"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   885
      Left            =   120
      Picture         =   "frmSoccer.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10290
      Width           =   1575
   End
   Begin VB.PictureBox picPitch 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      FillColor       =   &H8000000F&
      Height          =   10000
      Left            =   120
      MouseIcon       =   "frmSoccer.frx":22A4
      MousePointer    =   99  'Custom
      Picture         =   "frmSoccer.frx":23F6
      ScaleHeight     =   9945
      ScaleWidth      =   9945
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   10000
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(message)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   570
         Left            =   3702
         TabIndex        =   8
         Top             =   4715
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Image imgPlayer 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmSoccer.frx":B03AA
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgBall 
         Height          =   375
         Left            =   1440
         Picture         =   "frmSoccer.frx":B0C47
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgOpponent 
         Height          =   615
         Index           =   4
         Left            =   1440
         Picture         =   "frmSoccer.frx":B11F4
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgOpponent 
         Height          =   615
         Index           =   2
         Left            =   2880
         Picture         =   "frmSoccer.frx":B1A6F
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgOpponent 
         Height          =   615
         Index           =   3
         Left            =   2160
         Picture         =   "frmSoccer.frx":B22EA
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgOpponent 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmSoccer.frx":B2B65
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line linTarget 
         BorderColor     =   &H0000C000&
         Visible         =   0   'False
         X1              =   480
         X2              =   480
         Y1              =   1680
         Y2              =   1080
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   10200
      Width           =   10575
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         ItemData        =   "frmSoccer.frx":B33E0
         Left            =   240
         List            =   "frmSoccer.frx":B33F3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cboPlayers 
         Height          =   315
         ItemData        =   "frmSoccer.frx":B340B
         Left            =   3360
         List            =   "frmSoccer.frx":B342A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   495
         Width           =   2895
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ball speed: 50"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label lblPlayers 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of players:"
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   255
         Width           =   1665
      End
   End
   Begin VB.Image imgGoal 
      Height          =   615
      Left            =   11760
      Picture         =   "frmSoccer.frx":B344B
      Top             =   11400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameEnd 
         Caption         =   "&End game"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit game"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmSoccer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const pi = 3.1415926
Const ballDistance = 540
Const ballSize = 360 / 2
Const targetDistance = 6000
Const targetSize = 120 / 2

Dim intRotate As Integer
Dim intCounter As Integer
Dim strName As String
Dim i As Integer
Dim j As Integer

Dim intBallLeft As Integer
Dim intBallTop As Integer
Dim intTargetLeft As Integer
Dim intTargetTop As Integer
Dim sngStepLeft As Single
Dim sngStepTop As Single

Private Sub cboSpeed_Click()
    tmrBall.Interval = 50 - CInt(cboSpeed.Text) + 1
    lblInterval.Caption = "Ball speed: " & cboSpeed.Text
End Sub

Private Sub cmdShoot_Click()
    tmrBall.Enabled = False
    tmrShoot.Enabled = True
    cmdShoot.Enabled = False
    
    intCounter = 1
    
    PlayWave App.Path & "\sfx\kick.wav"
End Sub

Private Sub BeginGame()
    LoadLevel Me
    
    mnuGameNew.Enabled = False
    mnuGameEnd.Enabled = True
    fraOptions.Enabled = False
    
    lblMessage.Visible = False
End Sub

Private Sub EndGame(trigger As Byte, Optional score As Integer)
    Select Case trigger
        Case 0 ' Manual stop.
            lblMessage.Caption = ""
        Case 1 ' Bad pass.
            lblMessage.Caption = "bad pass"
            PlayWave App.Path & "\sfx\miss.wav"
        Case 2 ' Ball lost to opponent.
            lblMessage.Caption = "ball lost to opponent"
            PlayWave App.Path & "\sfx\miss.wav"
        Case 3 ' Goal.
            lblMessage.Caption = "goal, score: " & score
            PlayWave App.Path & "\sfx\goal.wav"
    End Select
    
    lblMessage.Visible = True
    
    tmrBall.Enabled = False
    tmrShoot.Enabled = False
    tmrTime.Enabled = False
    tmrOpponent.Enabled = False
    
    mnuGameNew.Enabled = True
    mnuGameEnd.Enabled = False
    cmdShoot.Enabled = False
    fraOptions.Enabled = True
    
    Sleep 1000
    
    For i = 2 To cboPlayers.Text + 1
        Unload imgPlayer(i)
    Next
    
    For i = 1 To 4
        imgOpponent(i).Visible = False
    Next
    
    imgPlayer(1).Visible = False
    imgBall.Visible = False
End Sub

Private Sub Form_Load()
    cboPlayers.ListIndex = 2
    cboSpeed.ListIndex = 4
    
    LoadScores Me, App.Path & "\highscores.hsc"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveScores Me, App.Path & "\highscores.hsc"
End Sub

Private Sub mnuGameEnd_Click()
    EndGame 0
End Sub

Private Sub mnuGameExit_Click()
    Unload Me
End Sub

Private Sub mnuGameNew_Click()
    BeginGame
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.LegalCopyright, vbInformation
End Sub

Private Sub tmrBall_Timer()
    If intRotate = 100 Then
        intRotate = 0
    Else
        intRotate = intRotate + 1
    End If
        
    intBallLeft = Math.Sin(intRotate * pi / 50) * ballDistance + (imgPlayer(intPlayerIndex).Left + (imgPlayer(intPlayerIndex).Width / 2))
    intBallTop = Math.Cos(intRotate * pi / 50 + pi) * ballDistance + (imgPlayer(intPlayerIndex).Top + (imgPlayer(intPlayerIndex).Height / 2))
    
    intTargetLeft = Math.Sin(intRotate * pi / 50) * targetDistance + (imgPlayer(intPlayerIndex).Left + (imgPlayer(intPlayerIndex).Width / 2))
    intTargetTop = Math.Cos(intRotate * pi / 50 + pi) * targetDistance + (imgPlayer(intPlayerIndex).Top + (imgPlayer(intPlayerIndex).Height / 2))
    
    imgBall.Left = intBallLeft - ballSize
    imgBall.Top = intBallTop - ballSize
    
    linTarget.X1 = (imgPlayer(intPlayerIndex).Left + (imgPlayer(intPlayerIndex).Width / 2))
    linTarget.Y1 = (imgPlayer(intPlayerIndex).Top + (imgPlayer(intPlayerIndex).Height / 2))
    linTarget.X2 = intTargetLeft - targetSize
    linTarget.Y2 = intTargetTop - targetSize
End Sub

Private Sub tmrOpponent_Timer()
    imgOpponent(1).Top = imgOpponent(1).Top + intOpponentSpeed
    imgOpponent(2).Top = imgOpponent(2).Top - intOpponentSpeed
    imgOpponent(3).Left = imgOpponent(3).Left + intOpponentSpeed
    imgOpponent(4).Left = imgOpponent(4).Left - intOpponentSpeed
    
    If (imgOpponent(1).Top > 9500 - imgOpponent(1).Height - 120) Or (imgOpponent(1).Top < 620) Then
        intOpponentSpeed = intOpponentSpeed * -1
    End If
End Sub

Private Sub tmrShoot_Timer()
    If intCounter < 30 Then
        imgBall.Left = imgBall.Left + (linTarget.X2 - (imgBall.Left + (imgBall.Width / 2))) / (30 - intCounter)
        imgBall.Top = imgBall.Top + (linTarget.Y2 - (imgBall.Top + (imgBall.Height / 2))) / (30 - intCounter)
    Else
        EndGame 1
        Exit Sub
    End If
    
    CheckContact 0, imgBall, imgPlayer(intPlayerIndex + 1), 120
    
    For i = 1 To 4
        CheckContact 2, imgBall, imgOpponent(i), 120
    Next
    
    intCounter = intCounter + 1
End Sub

Public Sub CheckContact(mode As Byte, firstObject As Object, secondObject As Object, span As Integer)
    Dim intFirstCenterLeft As Integer
    Dim intFirstCenterTop As Integer
    
    intFirstCenterLeft = firstObject.Left + (firstObject.Width / 2)
    intFirstCenterTop = firstObject.Top + (firstObject.Height / 2)

    If ((intFirstCenterLeft >= secondObject.Left - span) And (intFirstCenterLeft <= secondObject.Left + secondObject.Width + span)) And _
    ((intFirstCenterTop >= secondObject.Top - span) And (intFirstCenterTop <= secondObject.Top + secondObject.Height + span)) Then
        Select Case mode
            Case 0
                ProcessContact secondObject
            Case 1
                blnReposition = True
            Case 2
                EndGame 2
                Exit Sub
        End Select
    End If
End Sub

Private Sub ProcessContact(secondObject As Object)
    Dim intScore As Integer

    intPlayerIndex = secondObject.Index
    
    If intPlayerIndex = cboPlayers.Text + 1 Then
        intScore = (CInt(cboPlayers.Text) * 10) + CInt(CInt(cboSpeed.Text) / 5) - intTime
        If intScore < 0 Then intScore = 0
        
        EndGame 3, intScore
        
        strName = InputBox("What is your name?", , strName)
        If strName = "" Then strName = "Anonymous"
        
        lstHighscores.AddItem intScore & " | " & strName
        SortList lstHighscores
        
        Exit Sub
    End If
    
    imgPlayer(intPlayerIndex - 1).Visible = False
    imgPlayer(intPlayerIndex + 1).Visible = True
    
    tmrShoot.Enabled = False
    tmrBall.Enabled = True
    cmdShoot.Enabled = True
End Sub

Private Sub tmrTime_Timer()
    intTime = intTime + 1
    cmdShoot.Caption = Abs(intTime)
    
    If intTime = 0 Then
        cmdShoot.Enabled = True
        PlayWave App.Path & "\sfx\whistle.wav"
    ElseIf intTime < 0 Then
        PlayWave App.Path & "\sfx\tick.wav"
    End If
End Sub

Public Sub SortList(listToSort As ListBox)
    Dim strTemp As String

    For i = 0 To listToSort.ListCount - 1
        For j = i + 1 To listToSort.ListCount - 1
            If Val(listToSort.List(i)) < Val(listToSort.List(j)) Then
                strTemp = listToSort.List(i)
                listToSort.List(i) = listToSort.List(j)
                listToSort.List(j) = strTemp
            End If
        Next
    Next
End Sub
