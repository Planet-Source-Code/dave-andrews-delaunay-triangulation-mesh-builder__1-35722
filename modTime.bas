Attribute VB_Name = "modTime"

Function Long2Time(MyTime As Long) As String
Dim a As Integer
Dim B As Integer
Dim c As Integer
a = MyTime / 1000 'Seconds of play
B = a
c = 0
'If b > 60 Then
'    c = 1
'Else
'    c = 0
'End If
Do While (B / 60) >= 1
    B = B - 60
    c = c + 1
Loop
Long2Time = Format(c, "0#") & ":" & Format(B Mod 60, "0#")
End Function
Sub WaitFor(Secs As Single)
Dim XT
XT = Timer
Do While Timer - XT < Secs
    DoEvents
Loop
End Sub


