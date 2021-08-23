Option Explicit

' Version       : 0.1.0
' Author        : lxvs <jn.apsd@gmail.com>
' Last Updated  : 2021-04-14

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim col_num_status As Integer
    Dim col_num_history As Integer

    Dim col_status As String
    Dim col_history As String

    col_num_status = 5
    col_num_history = 7
    col_status = Choose(col_num_status, "A", "B", "C", "D", "E", "F", "G", "H", "I")
    col_history = Choose(col_num_history, "A", "B", "C", "D", "E", "F", "G", "H", "I")

    If Target.Column = col_num_status And Sheet1.Range(col_status & Target.Row) <> "" And Sheet1.Range("A" & Target.Row) <> "" Then
        If Sheet1.Range(col_history & Target.Row) = "" Then
            Sheet1.Range(col_history & Target.Row) = Format(Now, "yyyy-MM-dd") & " " & Target
        ElseIf InStr(Sheet1.Range(col_history & Target.Row), "|") = 0 Then
            If Sheet1.Range(col_history & Target.Row) <> Format(Now, "yyyy-MM-dd") & " " & Target Then Sheet1.Range(col_history & Target.Row) = Format(Now, "yyyy-MM-dd") & " " & Target & "  |  " & Sheet1.Range(col_history & Target.Row)
        Else
            Dim logs() As String
            logs = split(Sheet1.Range(col_history & Target.Row), "  |  ")
            If logs(0) <> Format(Now, "yyyy-MM-dd") & " " & Target Then Sheet1.Range(col_history & Target.Row) = Format(Now, "yyyy-MM-dd") & " " & Target & "  |  " & Sheet1.Range(col_history & Target.Row)
        End If
    End If
End Sub
