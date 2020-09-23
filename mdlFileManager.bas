Attribute VB_Name = "mdlFileManager"
Option Explicit

Public Sub LoadScores(sender As Form, filename As String)
    Dim strData As String
    
    Open filename For Input As #1
        Do While Not EOF(1)
            Input #1, strData
            sender.lstHighscores.AddItem strData
        Loop
    Close #1
    
    sender.SortList sender.lstHighscores
End Sub

Public Sub SaveScores(sender As Form, filename As String)
    Dim i As Integer
    
    Open filename For Output As #1
        For i = 0 To sender.lstHighscores.ListCount - 1
            Write #1, sender.lstHighscores.List(i)
        Next
    Close #1
End Sub
