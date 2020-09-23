Attribute VB_Name = "mdlLevelManager"
Option Explicit

Public intPlayerIndex As Integer
Public blnReposition As Boolean
Public intOpponentSpeed As Integer
Public intTime As Integer

Public Sub LoadLevel(sender As Form)
    Dim i As Integer
    Dim intPrevBallLeft As Integer
    Dim intPrevBallTop As Integer

    Randomize

    sender.imgPlayer(1).Left = Int((6000 - 4000 + 1) * Rnd + 4000)
    sender.imgPlayer(1).Top = Int((6000 - 4000 + 1) * Rnd + 4000)
    
    For i = 2 To sender.cboPlayers.Text + 1
        Load sender.imgPlayer(i)
        
        intPrevBallLeft = sender.imgPlayer(i - 1).Left + (sender.imgPlayer(i - 1).Width / 2)
        intPrevBallTop = sender.imgPlayer(i - 1).Top + (sender.imgPlayer(i - 1).Height / 2)

Repo:
        Randomize
        
        blnReposition = False

        sender.imgPlayer(i).Left = Int(((sender.imgPlayer(i - 1).Left + 5000) - (sender.imgPlayer(i - 1).Left - 5000) + 1) * Rnd + (sender.imgPlayer(i - 1).Left - 5000))
        sender.imgPlayer(i).Top = Int(((sender.imgPlayer(i - 1).Top + 5000) - (sender.imgPlayer(i - 1).Top - 5000) + 1) * Rnd + (sender.imgPlayer(i - 1).Top - 5000))
        
        If sender.imgPlayer(i).Left < 1000 Then sender.imgPlayer(i).Left = 1000
        If sender.imgPlayer(i).Left > 9000 Then sender.imgPlayer(i).Left = 9000
        If sender.imgPlayer(i).Top < 1000 Then sender.imgPlayer(i).Top = 1000
        If sender.imgPlayer(i).Top > 9000 Then sender.imgPlayer(i).Top = 9000
        
        sender.CheckContact 1, sender.imgPlayer(i), sender.imgPlayer(i - 1), 615
        
        If blnReposition Then GoTo Repo
    Next

    sender.imgPlayer(sender.cboPlayers.Text + 1).Picture = sender.imgGoal.Picture
    sender.tmrBall.Interval = 50 - CInt(sender.cboSpeed.Text) + 1
    
    Initialize sender
End Sub

Private Sub Initialize(sender As Form)
    Dim i As Integer
    
    intTime = -3
    sender.cmdShoot.Caption = "3"
    
    sender.tmrBall.Enabled = True
    sender.tmrOpponent.Enabled = True
    sender.tmrTime.Enabled = True
    
    sender.imgBall.Visible = True
    
    sender.imgOpponent(1).Top = 620
    sender.imgOpponent(1).Left = 2500
    
    sender.imgOpponent(2).Top = 8760
    sender.imgOpponent(2).Left = 7500 - sender.imgOpponent(2).Width
    
    sender.imgOpponent(3).Top = 7500 - sender.imgOpponent(3).Height
    sender.imgOpponent(3).Left = 620
    
    sender.imgOpponent(4).Top = 2500
    sender.imgOpponent(4).Left = 8760
    
    For i = 1 To 4
        sender.imgOpponent(i).Visible = True
    Next
    
    sender.imgPlayer(1).Visible = True
    sender.imgPlayer(2).Visible = True
    
    intPlayerIndex = 1
    intOpponentSpeed = 30
End Sub
