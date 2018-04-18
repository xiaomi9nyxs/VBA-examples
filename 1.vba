''''''''第一环节选题  第18
Option Base 1
Dim Questions(48) As String
Dim Answers(48) As String
Dim flag_init As Boolean

''''''''公共函数
Public Function Load_Question()
    Dim app_excel As Excel.Application
    Dim current_wb As Excel.Workbook
    Set app_excel = New Excel.Application
    Dim str As String
    str = ActivePresentation.Path & "\xxx.xls"
    If Dir(str) <> "" Then
            Set current_wb = app_excel.Workbooks.Open(str)
            n = current_wb.Sheets("1").UsedRange.Rows.Count
            For i = 1 To 48
                Questions(i) = current_wb.Sheets("1").Cells(i + 1, 2) ''''''''''''题目
                Answers(i) = current_wb.Sheets("1").Cells(i + 1, 3) ''''''''''''答案
            Next
    End If
    current_wb.Close 1
    app_excel.Quit
    Set current_wb = Nothing
    Set app_excel = Nothing
End Function
Private Sub choose_Q(ByVal num As Integer)
    Dim index1 As Integer
    Dim index2 As Integer
    Dim index3 As Integer
    index1 = 1 + (num - 1) * 3
    index2 = 2 + (num - 1) * 3
    index3 = 3 + (num - 1) * 3
    If flag_init Then
    With Slide18
        .Shapes("question").TextFrame2.TextRange.Text = Questions(index1)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(index1)
        .Shapes("answer").Visible = msoCFalse
    End With
    With Slide19
        .Shapes("question").TextFrame2.TextRange.Text = Questions(index2)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(index2)
        .Shapes("answer").Visible = msoCFalse
    End With
    With Slide20
        .Shapes("question").TextFrame2.TextRange.Text = Questions(index3)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(index3)
        .Shapes("answer").Visible = msoCFalse
    End With
    If Len(Questions(index1)) > 80 Then
        Slide18.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide18.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    If Len(Questions(index2)) > 80 Then
        Slide19.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide19.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    If Len(Questions(index3)) > 80 Then
        Slide20.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide20.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    SlideShowWindows(1).Activate
    SlideShowWindows(1).View.GotoSlide 19
    End If
End Sub
Sub choose_Q_1()
    With Me
        .Shapes("Q1").Visible = msoCFalse
    End With
    Call choose_Q(1)
End Sub
Sub choose_Q_2()
With Me
        .Shapes("Q2").Visible = msoCFalse
    End With
    Call choose_Q(2)
End Sub
Sub choose_Q_3()
With Me
        .Shapes("Q3").Visible = msoCFalse
    End With
    Call choose_Q(3)
End Sub
Sub choose_Q_4()
With Me
        .Shapes("Q4").Visible = msoCFalse
    End With
    Call choose_Q(4)
End Sub
Sub choose_Q_5()
    With Me
        .Shapes("Q5").Visible = msoCFalse
    End With
    Call choose_Q(5)
End Sub
Sub choose_Q_6()
    With Me
        .Shapes("Q6").Visible = msoCFalse
    End With
    Call choose_Q(6)
End Sub
Sub choose_Q_7()
    With Me
        .Shapes("Q7").Visible = msoCFalse
    End With
    Call choose_Q(7)
End Sub
Sub choose_Q_8()
    With Me
        .Shapes("Q8").Visible = msoCFalse
    End With
    Call choose_Q(8)
End Sub
Sub choose_Q_9()
    With Me
        .Shapes("Q9").Visible = msoCFalse
    End With
    Call choose_Q(9)
End Sub
Sub choose_Q_10()
    With Me
        .Shapes("Q10").Visible = msoCFalse
    End With
    Call choose_Q(10)
End Sub
Sub choose_Q_11()
    With Me
        .Shapes("Q11").Visible = msoCFalse
    End With
    Call choose_Q(11)
End Sub
Sub choose_Q_12()
    With Me
        .Shapes("Q12").Visible = msoCFalse
    End With
    Call choose_Q(12)
End Sub
Sub choose_Q_13()
    With Me
        .Shapes("Q13").Visible = msoCFalse
    End With
    Call choose_Q(13)
End Sub
Sub choose_Q_14()
    With Me
        .Shapes("Q14").Visible = msoCFalse
    End With
    Call choose_Q(14)
End Sub
Sub choose_Q_15()
    With Me
        .Shapes("Q15").Visible = msoCFalse
    End With
    Call choose_Q(15)
End Sub
Sub choose_Q_16()
    With Me
        .Shapes("Q16").Visible = msoCFalse
    End With
    Call choose_Q(16)
End Sub


Private Sub Choose_Q_1_1()
If flag_init Then
    With Slide18
        .Shapes("question").TextFrame2.TextRange.Text = Questions(1)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(1)
        .Shapes("answer").Visible = msoCFalse
    End With
    If Len(Questions(index)) > 50 Then
        Slide18.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide18.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    SlideShowWindows(1).Activate
    SlideShowWindows(1).View.GotoSlide 19
End If
End Sub

Private Sub Choose_Q_1_2()
If flag_init Then
    With Slide19
        .Shapes("question").TextFrame2.TextRange.Text = Questions(2)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(2)
        .Shapes("answer").Visible = msoCFalse
    End With
    If Len(Questions(index)) > 50 Then
        Slide19.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide19.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    SlideShowWindows(1).Activate
    SlideShowWindows(1).View.GotoSlide 20
End If
End Sub

Private Sub Choose_Q_1_3()
If flag_init Then
    With Slide20
        .Shapes("question").TextFrame2.TextRange.Text = Questions(3)
        .Shapes("answer").TextFrame2.TextRange.Text = Answers(3)
        .Shapes("answer").Visible = msoCFalse
    End With
    If Len(Questions(index)) > 50 Then
        Slide20.Shapes("question").TextFrame2.TextRange.Font.Size = 20
    Else
        Slide20.Shapes("question").TextFrame2.TextRange.Font.Size = 24
    End If
    SlideShowWindows(1).Activate
    SlideShowWindows(1).View.GotoSlide 21
End If
End Sub

Private Sub BTN_INIT_Click()
    Call init_1
End Sub

Public Sub Next_round()
SlideShowWindows(1).Activate
SlideShowWindows(1).View.GotoSlide 22
End Sub

Public Sub Return_1()
SlideShowWindows(1).Activate
SlideShowWindows(1).View.Previous
End Sub

Public Sub init_1()
    Call Load_Question
    With Me
        .Shapes.Range(Array("Q1", "Q2", "Q3", _
        "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10", _
        "Q11", "Q12", "Q13")).Visible = msoCTrue
    End With
    Slide18.Shapes("Timer").TextFrame2.TextRange.Text = "30"
    Slide19.Shapes("Timer").TextFrame2.TextRange.Text = "30"
    Slide20.Shapes("Timer").TextFrame2.TextRange.Text = "30"
    flag_init = True
End Sub
