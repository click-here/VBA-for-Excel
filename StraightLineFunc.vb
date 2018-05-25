Option Explicit
Function StraightLineFunc(headRng As Range, dataRng As Range) As Double
    Application.Volatile True 'Forces update if values are changed/rows inserted
    Dim arrCntr As Integer 'Declare our array counter
    Dim arr() As Variant 'Declare our array
    Dim cntr As Integer
    Dim stdvTotal As Double
    Dim cell As Variant
    
    stdvTotal = 0
    cntr = 0
    arrCntr = 1

    For Each cell In headRng
        'If cell contains text and it's neither "Response" nor "Open-Ended Response" then
        'we know that it's the text for a question group
        If cell <> "Response" And cell <> "Open-Ended Response" And cell <> "" Then
            'If the row above the question text contains text and cntr > 0 then
            'we know that a question group has just completed
            If cell.Offset(-1, 0) <> "" And cntr > 0 Then
                stdvTotal = stdvTotal + WorksheetFunction.StDev(arr)
            End If
            'If the row above contains text than
            'we've started a new question grouping
            If cell.Offset(-1, 0) <> "" Then
                cntr = cntr + 1
                Erase arr
                'Declaring the size of our array large enough fit the coming array
                ReDim arr(headRng.Columns.Count)
                arrCntr = 1
                'Save the value of the cell containing the question value
                'to be checked for straight-lining
                arr(arrCntr) = cell.Offset(dataRng.Row - headRng.Row, 0).Value
                arrCntr = arrCntr + 1
            End If
        End If
    Next cell
    stdvTotal = stdvTotal + WorksheetFunction.StDev(arr)
    StraightLineFunc = stdvTotal
End Function
