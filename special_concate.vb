Option Explicit
Function SpecConcat(headRng As Range, dataRng As Range) As String
    Dim loopcnt As Integer, i As Integer
    Dim outputArr() As String, txt As String
    ReDim outputArr(Application.WorksheetFunction.CountIf(Range(dataRng.Address), "Yes"))
    Dim cell As Variant
    
    loopcnt = 1
    For Each cell In dataRng
        If cell = "Yes" Then
            outputArr(loopcnt) = cell.Offset(-(dataRng.Row - headRng.Row), 0).Value
            loopcnt = loopcnt + 1
        End If
    Next cell

    For i = 1 To UBound(outputArr)
        If i = 1 Then
            txt = outputArr(i) & ", "
        ElseIf i = UBound(outputArr) Then
            txt = txt & outputArr(i)
        Else
            txt = txt & outputArr(i) & ", "
        End If
    Next i
    SpecConcat = txt
End Function
