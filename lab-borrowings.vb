Option Explicit

' Version       : 0.2.0
' Author        : lxvs <jn.apsd@gmail.com>
' Last Updated  : 2021-08-24

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim col_num_status As Integer
    Dim col_num_history As Integer
    Dim col_status As String
    Dim col_history As String
    Dim date_now As String
    Dim r As Integer
    Dim i As Integer

    col_num_status = 5
    col_num_history = 7
    col_status = Choose(col_num_status, _
        "A", "B", "C", "D", "E", "F", "G", "H", "I")
    col_history = Choose(col_num_history, _
        "A", "B", "C", "D", "E", "F", "G", "H", "I")

    If Intersect(Target, Sheet1.Range(col_status & ":" & col_status)) _
        Is Nothing Then
        Exit Sub
    End If

    For i = 0 To Target.Rows.Count - 1
        r = Target.Row + i
        If Sheet1.Range(col_status & r) = "" _
            Or Sheet1.Range("A" & r) = "" Then
            GoTo continue
        End If
        date_now = Format(Now, "yyyy-MM-dd")
        If Sheet1.Range(col_history & r) = "" Then
            Sheet1.Range(col_history & r) = _
                date_now & " " & Sheet1.Range(col_status & r)
            GoTo continue
        End If
        If InStr(Sheet1.Range(col_history & r), vbLf) = 0 _
            And Sheet1.Range(col_history & r) _
            <> date_now & " " & Sheet1.Range(col_status & r) Then
            Sheet1.Range(col_history & r) _
                = date_now & " " & Sheet1.Range(col_status & r) _
                & vbLf & Sheet1.Range(col_history & r)
            GoTo continue
        End If
        Dim logs() As String
        logs = Split(Sheet1.Range(col_history & r), vbLf)
        If logs(0) <> date_now & " " & Sheet1.Range(col_status & r) Then
            Sheet1.Range(col_history & r) _
                = date_now & " " & Sheet1.Range(col_status & r) _
                & vbLf & Sheet1.Range(col_history & r)
        End If
continue:
    Next i
End Sub
