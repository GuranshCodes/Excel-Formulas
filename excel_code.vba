Private Sub Worksheet_Change(ByVal Target As Range)
    Dim row As Long
    Dim mValue As Double
    Dim pFormula As String
    Dim dayNum As Integer
    Dim currentMonth As String
    Dim currentYear As String

    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False


    ' =========================
    ' M19:M30
    ' =========================
    If Not Intersect(Target, Me.Range("M19:M30")) Is Nothing Then

        row = Target.Row

        If Target.Value = "" Then

            Me.Cells(row, 15).Value = 4.49
            pFormula = "=M" & row & " * O" & row & " / 100"
            Me.Cells(row, 16).Formula = pFormula

        ElseIf IsNumeric(Target.Value) Then

            mValue = Target.Value
            Target.Value = mValue * 2.20462262

            ' =========================
            ' PRICE SYSTEM (UPDATED)
            ' =========================
            If Target.Value > 15000 Then
                Me.Cells(row, 15).Value = 2.99

            ElseIf Target.Value > 10000 Then
                Me.Cells(row, 15).Value = 3.29

            ElseIf Target.Value >= 5000 Then
                Me.Cells(row, 15).Value = 3.89

            Else
                Me.Cells(row, 15).Value = 4.49
            End If

            Me.Cells(row, 16).Formula = "=M" & row & " * O" & row & " / 100"

            If Target.Value < 1450 Then
                Me.Cells(row, 16).Value = 65.5
            End If
        End If
    End If


    ' =========================
    ' O19:O30
    ' =========================
    If Not Intersect(Target, Me.Range("O19:O30")) Is Nothing Then

        row = Target.Row

        pFormula = "=M" & row & " * O" & row & " / 100"
        Me.Cells(row, 16).Formula = pFormula

        If Me.Cells(row, 13).Value < 1450 And IsNumeric(Me.Cells(row, 13).Value) Then
            Me.Cells(row, 16).Value = 65.5
        End If

        If IsNumeric(Me.Cells(row, 13).Value) Then

            If Me.Cells(row, 13).Value > 15000 Then
                Me.Cells(row, 15).Value = 2.99

            ElseIf Me.Cells(row, 13).Value > 10000 Then
                Me.Cells(row, 15).Value = 3.29

            ElseIf Me.Cells(row, 13).Value >= 5000 Then
                Me.Cells(row, 15).Value = 3.89
            End If

        End If
    End If


    ' =========================
    ' G19:G30
    ' =========================
    If Not Intersect(Target, Me.Range("G19:G30")) Is Nothing Then

        row = Target.Row

        If Trim(Target.Value) = "" Then
            Target.ClearContents

        ElseIf IsNumeric(Target.Value) And InStr(Target.Value, "/") = 0 Then

            dayNum = CInt(Target.Value)

            If dayNum > 0 And dayNum <= 31 Then
                currentMonth = UCase(Format(Date, "MMM"))
                currentYear = Right(Format(Date, "YYYY"), 2)
                Target.Value = currentMonth & "/" & Format(dayNum, "00") & "/" & currentYear
            Else
                MsgBox "Invalid day number. Please enter 1–31.", vbExclamation
                Target.ClearContents
            End If
        End If
    End If


SafeExit:
    Application.EnableEvents = True
End Sub
