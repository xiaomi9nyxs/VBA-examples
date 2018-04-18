'Dim bofang As Boolean
Dim counter As Integer
Dim counter_my As Integer
Dim second As Boolean
Option Explicit
Private TimerID As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As _
        Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As _
        Long, ByVal nIDEvent As Long) As Long
'''''''''''''''''''''''''''''''
' m用于保存计时时间（分钟），n用于保存秒，c保存原始计时
Private m As Long, n As Long, c As String
Dim sld_show As Slide

Private Function CreateTimer(ByVal Interval As Long) As Long
    ' 建立一个时间间隔为Interval微秒的定时器
    Dim tID As Long
    tID = SetTimer(0, 0, Interval, AddressOf TimerProc)
    CreateTimer = tID
End Function

' 终止tID标识的定时器
Private Sub TerminateTimer(ByVal tID As Long)
    Call KillTimer(0, tID)
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal idEvent As Long, ByVal dwTime As Long)
    ' 此处放入要执行的代码
    CounterNumber
End Sub

' 开始定时器
'计时长度由counter决定
Private Sub BeginTimer(sld As Slide, time_period As Integer)
'    Call Play    '''播放音乐，模块2中
    Set sld_show = sld
    If TimerID <> 0 Then TerminateTimer TimerID
    ' 获取tm形状内所设置的计时时间
    c = sld_show.Shapes("Timer").TextFrame2.TextRange.Text
    sld_show.Shapes("Timer").Visible = msoTrue
    counter = time_period
    second = False ''''''''''自定义
    counter_my = 5 '''''''''自定义
    TimerID = CreateTimer(100)   '''''500ms
End Sub

' 结束定时器
Private Sub EndTimer()
    Call Pause    '''结束音乐，模块2中
    If TimerID <> 0 Then
        Call TerminateTimer(TimerID)
        TimerID = 0
        ' 恢复原始计时时间
'        sld_show.Shapes("Timer").TextFrame2.TextRange.Text = c
        sld_show.Shapes("Timer").TextFrame2.TextRange.Text = "00"
        c = ""
    End If
'bofang = True
End Sub

' 暂停定时器
Private Sub PauseTimer()
    Call Pause    '''结束音乐，模块2中
    If TimerID <> 0 Then
        Call TerminateTimer(TimerID)
        TimerID = 0
        ' 恢复原始计时时间
        sld_show.Shapes("Timer").Visible = msoCTrue
    End If
'bofang = True
End Sub

Private Sub T()
    Call Pause    '''结束音乐，模块2中
    If TimerID <> 0 Then
        Call TerminateTimer(TimerID)
        TimerID = 0
        ' 恢复原始计时时间
        sld_show.Shapes("Timer").Visible = msoCTrue
    End If
'bofang = True
End Sub


''''显示时间变化
Private Sub CounterNumber()
'If second = False Then
'    second = True
'Else
'    second = False
'End If
If counter_my = 5 Then '''''''''''''''''''''
    second = Not second
    If counter < 10 Then
        sld_show.Shapes("Timer").TextFrame2.TextRange.Text = "0" & counter
    Else
        sld_show.Shapes("Timer").TextFrame2.TextRange.Text = counter
    End If
    If counter = 5 Then
        Call Play    '''播放音乐，模块2中
    End If
    
    If counter > 5 Then
        If second Then
            counter = counter - 1
        End If
    ElseIf counter > 0 Then
        If second Then
            sld_show.Shapes("Timer").Visible = msoTriStateToggle
            counter = counter - 1
        End If
    Else
        counter = 0
        sld_show.Shapes("Timer").Visible = msoTrue
        Call EndTimer
    End If
    counter_my = 0
Else
    counter_my = counter_my + 1
    If counter < 6 Then
        sld_show.Shapes("Timer").Visible = msoTriStateToggle
    End If
End If
End Sub

Sub Start_Time(ByVal sld As Slide, ByVal time_period As Integer)
    If time_period > 5 Then
        Call BeginTimer(sld, time_period)
    End If
End Sub

Sub Reset_Time()
Call EndTimer
'sld_show.Shapes("Timer").Visible = msoFalse
End Sub

Sub Pause_Time()
Call PauseTimer
End Sub

'''显示音乐模块
Option Explicit
Public Declare Function mciSendString Lib "winmm.dll" _
    Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    
Private Function ConvShortFilename(ByVal strLongPath$) As String
    Dim strShortPath$
    '''$ 表示字符串
    '''%表示Integer
    '''&表示Long
    '''!表示Single
    '''@表示Currency
    '''#表示Double
    If InStr(1, strLongPath, " ") Then
        strShortPath = String(LenB(strLongPath), Chr(0))
        GetShortPathName strLongPath, strShortPath, Len(strShortPath)
        ConvShortFilename = Left(strShortPath, InStr(1, strShortPath, Chr(0)) - 1)
    Else
        ConvShortFilename = strLongPath
    End If
End Function

Public Sub MMPlay(ByRef FileName As String)
    FileName = ConvShortFilename(FileName)
    mciSendString "close " & FileName, vbNullString, 0, 0
    mciSendString "open " & FileName, vbNullString, 0, 0
    mciSendString "play " & FileName, vbNullString, 0, 0
End Sub

Public Sub MMStop(ByRef FileName As String)
    FileName = ConvShortFilename(FileName)
    mciSendString "stop " & FileName, vbNullString, 0, 0
    mciSendString "close " & FileName, vbNullString, 0, 0
End Sub

Sub Play()
    MMPlay (ActivePresentation.Path & "\5.mp3")
End Sub

Sub Pause()
    MMStop (ActivePresentation.Path & "\5.mp3")
End Sub


