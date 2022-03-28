Option Explicit

' Version       : 0.3.0
' Author        : lxvs <jn.apsd@gmail.com>
' Last Updated  : 2022-03-28

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim sheet As Worksheet
    Dim col_num_status As Integer
    Dim col_num_history As Integer
    Dim col_status As String
    Dim col_history As String
    Dim date_now As String
    Dim r As Integer
    Dim i As Integer

    Set sheet = Sheet1
    col_num_status = 5
    col_num_history = 7
    col_status = Choose(col_num_status, _
        "A", "B", "C", "D", "E", "F", "G", "H", "I")
    col_history = Choose(col_num_history, _
        "A", "B", "C", "D", "E", "F", "G", "H", "I")

    If Intersect(Target, sheet.Range(col_status & ":" & col_status)) _
        Is Nothing Then
        Exit Sub
    End If

    For i = 0 To Target.Rows.Count - 1
        r = Target.Row + i
        If sheet.Range(col_status & r) = "" _
            Or sheet.Range("A" & r) = "" Then
            GoTo continue
        End If
        date_now = Format(Now, "yyyy-MM-dd")
        If sheet.Range(col_history & r) = "" Then
            sheet.Range(col_history & r) = _
                date_now & " " & sheet.Range(col_status & r)
            GoTo continue
        End If
        Dim logs() As String
        logs = Split(sheet.Range(col_history & r), vbLf)
        If logs(0) <> date_now & " " & sheet.Range(col_status & r) Then
            sheet.Range(col_history & r) _
                = date_now & " " & sheet.Range(col_status & r) _
                & vbLf & sheet.Range(col_history & r)
        End If
continue:
    Next i
End Sub
