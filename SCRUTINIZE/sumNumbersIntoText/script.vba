'==============================
' Function: SumNumbersInText
' Description: Extracts and sums all numeric sequences found in a given Excel cell.
' Example: If cell A1 contains "abc123def45", this function returns 168 (123 + 45).
'==============================
Function SumNumbersInText(cell As Range) As Double
    Dim total As Double         ' Accumulates the sum of found numbers
    Dim temp As String          ' Temporarily holds digits to form a complete number
    Dim i As Integer            ' Loop counter
    Dim char As String          ' Holds the current character from the cell value

    total = 0
    temp = ""
    
    ' Exit early if the cell is empty to avoid unnecessary processing
    If cell.Value = "" Then
        SumNumbersInText = 0
        Exit Function
    End If

    ' Loop through each character in the cell's text
    For i = 1 To Len(cell.Value)
        char = Mid(cell.Value, i, 1)
        
        ' Check if the character is numeric (0-9)
        If IsNumeric(char) Then
            temp = temp & char  ' Build the number as a string
        Else
            ' Non-numeric character encountered, so finalize the current number
            If temp <> "" Then
                total = total + Val(temp)  ' Convert and add the number to the total
                temp = ""                 ' Reset for the next number
            End If
        End If
    Next i
    
    ' Add any number that may be left in temp after the loop ends
    If temp <> "" Then
        total = total + Val(temp)
    End If
    
    ' Return the final sum
    SumNumbersInText = total
End Function

'==============================
' Subroutine: num
' Description: Demonstrates how to call the SumNumbersInText function.
' It reads the value from cell A1, calculates the sum of embedded numbers,
' and displays the result in a message box.
'==============================
Sub num()
    Dim result As Double  ' Holds the result returned by the function
    result = SumNumbersInText(Range("A1"))  ' Call the function with cell A1 as input
    MsgBox "The total is " & result         ' Display the result to the user
End Sub
