Attribute VB_Name = "Module1"
Sub CountLettersAndUpdate()
    Dim lastRow As Long
    Dim counter As Long
    Dim i As Long
    
    Dim userInput As String
    userInput = InputBox("Моля въведете номер на задание до което сте стигнали:", "User Input")
    
    ' Check if F2 is empty, if so initialize counter to 0
    If IsEmpty(Range("E2").Value) Then
        counter = userInput
    Else
        counter = Range("E2").Value ' Use the existing value in F2
    End If
    
    ' Get the last row in column H (assuming H is the column with data)
    lastRow = Cells(Rows.Count, "G").End(xlUp).Row
    
    ' Loop through each row in column H starting from row 2
    For i = 2 To lastRow
        ' Check if the value in column H is not empty
        If Not IsEmpty(Cells(i, "G").Value) Then
            ' Check if the value in column H is either "?" or "?"
            If Cells(i, "G").Value = "П" Or Cells(i, "G").Value = "К" Then
                ' Increment the counter
                ' Update column F with the new counter value (skip row 1)
                Cells(i, "E").Value = counter
                counter = counter + 1
            End If
        End If
        If counter = 101 Then
            counter = 1
        End If
    Next i
End Sub
