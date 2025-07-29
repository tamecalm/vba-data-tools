' ==========================================================================================
' Subroutine: InsertRegionBasedOnState
' Purpose   : Automatically assign a geopolitical region (North-Central, North-East,
'             or North-West) to each row based on the state name found in Column A.
' Output    : Writes the region name to Column J (9 columns to the right of Column A).
' Assumptions:
'     - Worksheet name is "Pivot (Old)"
'     - States are stored in Column A (starting from A2)
'     - Region output is written in Column J
' ==========================================================================================

Sub InsertRegionBasedOnState()
    Dim ws As Worksheet
    Dim stateColumn As Range
    Dim stateCell As Range
    Dim state As String
    Dim region As String
    Dim matchFound As Boolean

    ' Set the worksheet (adjust if your sheet name is different)
    Set ws = ThisWorkbook.Sheets("Pivot (Old)")

    ' Define the range containing state data in Column A
    Set stateColumn = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Loop through each cell in the range
    For Each stateCell In stateColumn
        state = stateCell.Value
        matchFound = False

        ' Process only if the cell is not empty
        If state <> "" Then
            ' Match the state against known regions using partial, case-insensitive match
            Select Case True

                ' NORTH-CENTRAL States
                Case InStr(1, state, "Benue", vbTextCompare) > 0, _
                     InStr(1, state, "FCT", vbTextCompare) > 0, _
                     InStr(1, state, "Kaduna", vbTextCompare) > 0, _
                     InStr(1, state, "Kogi", vbTextCompare) > 0, _
                     InStr(1, state, "Nasarawa", vbTextCompare) > 0, _
                     InStr(1, state, "Niger", vbTextCompare) > 0
                    region = "NORTH-CENTRAL"
                    matchFound = True

                ' NORTH-EAST States
                Case InStr(1, state, "Adamawa", vbTextCompare) > 0, _
                     InStr(1, state, "Bauchi", vbTextCompare) > 0, _
                     InStr(1, state, "Borno", vbTextCompare) > 0, _
                     InStr(1, state, "Gombe", vbTextCompare) > 0, _
                     InStr(1, state, "Plateau", vbTextCompare) > 0, _
                     InStr(1, state, "Taraba", vbTextCompare) > 0, _
                     InStr(1, state, "Yobe", vbTextCompare) > 0
                    region = "NORTH-EAST"
                    matchFound = True

                ' NORTH-WEST States
                Case InStr(1, state, "Jigawa", vbTextCompare) > 0, _
                     InStr(1, state, "Kano", vbTextCompare) > 0, _
                     InStr(1, state, "Katsina", vbTextCompare) > 0, _
                     InStr(1, state, "Kebbi", vbTextCompare) > 0, _
                     InStr(1, state, "Sokoto", vbTextCompare) > 0, _
                     InStr(1, state, "Zamfara", vbTextCompare) > 0
                    region = "NORTH-WEST"
                    matchFound = True

            End Select

            ' Write the region to Column J if a match was found
            If matchFound Then
                stateCell.Offset(0, 9).Value = region ' Column J
            Else
                stateCell.Offset(0, 9).ClearContents ' Clear region if unmatched
            End If
        Else
            ' Clear the output if source cell is empty
            stateCell.Offset(0, 9).ClearContents
        End If
    Next stateCell

    MsgBox "REGION HAVE BEEN INSERTED SUCCESSFULLY!"
End Sub
