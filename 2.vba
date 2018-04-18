''''''''第一环节第一题19
Private Sub BTN_RET_Click()

End Sub

Public Sub Show()
With Me
    .Shapes("answer").Visible = msoCTrue
End With
End Sub

Public Sub Return_1()
    With Me
        .Shapes("answer").Visible = msoCFalse
    End With
    SlideShowWindows(1).View.Previous
End Sub

Public Sub Next_Page()
    Call GStop_Time
    With Me
        .Shapes("answer").Visible = msoCTrue
    End With
    SlideShowWindows(1).View.Next
End Sub

Public Sub GStart_Time()
Call Start_Time(Slide18, 30)
End Sub

Public Sub GStop_Time()
Call Reset_Time
Me.Shapes("timer").TextFrame2.TextRange.Text = "30"
End Sub

Public Sub GPause_Time()
Call Pause_Time
End Sub
