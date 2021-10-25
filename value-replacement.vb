Private Sub Worksheet_Change(ByVal Target As Range)
    For Each cell In Target
        If cell.Row < 40 And Not IsEmpty(cell.Value) Then
            Set found = Range("40:40").Find(cell.Value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                cell.Value = found.Offset(1, 0).Value
            End If
        End If
    Next cell
End Sub
