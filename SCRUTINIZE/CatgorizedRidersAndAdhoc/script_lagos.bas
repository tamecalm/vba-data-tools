' ====================================================================================
' Subroutine: CategorizeRidersAdhocAndHubOfficers
' Purpose   : Categorizes entries in column A of the "Pivot (Old)" sheet into
'             HUB OFFICER, RIDER, ADHOC, IN HOUSE CLEANER, or Not In List, based on name matching
'             against predefined arrays.
' Output    : Writes the category into Column K (10 columns to the right of Column A)
' ====================================================================================

Sub CategorizeRidersAdhocAndHubOfficers()
    Dim ws As Worksheet
    Dim stateColumn As Range
    Dim stateCell As Range
    Dim state As String
    Dim ridersList As Variant
    Dim adhocList As Variant
    Dim hubOfficerList As Variant
    Dim ipdsList As Variant
    Dim psbList As Variant
    Dim cleanerList As Variant
    Dim statesList As Variant
    Dim category As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Pivot (Old)")

    ' Set the range for the column to be checked (assumes names in Column A)
    Set stateColumn = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)


    ' Adhoc list
    adhocList = Array("ADEBAYO AYOMIDE", "Adebayo Christianah", "ADESANWO ADEYEMI", _
                      "ALIU JAMIU", "BABATUNDE DELE", "Charles Okafor(T)", _
                      "Dairo Titilola", "DANIEL OMASORO", "FAITH ONOJA", _
                      "Gbajumo Oladipo", "GODFREY NEKABARI", "Jubreal Hamzat", _
                      "Kingsley Okoye", "Ocheli Joyce", "OLAWALE ODEYINGBO", _
                      "GRACE OKPARA", "Lasaki Nimotallahi(T)", "Oteikwu Godfrey", _
                      "Ronald Usoro", "SHAKIRAT ISHOLA", "OGUNLEYE JEREMIAH", _
                      "BASIL ULUGBE")

    ' In House Cleaner list
    cleanerList = Array("Zainab Uthman")

    ' Riders list
    ridersList = Array("BAMIDELE OJO", "Cares Akohamen", "David Edegbo", _
                       "ESSU CHUKS", "HENRY TOBY", "JOHNSON ALARAPE", _
                       "MONSURU AJINIKIRUN", "MUHAMMED HASSAN", "Njiaka Onyemaechi", _
                       "ODEH BENEDICT", "OJOR PAUL", "Raphael Adebowale", _
                       "SULEIMAN PETERS", "TAIRU YUSUF")

    ' Ipds list (Update with Lagos names if applicable)
    ipdsList = Array()         

    ' Psb list (Update with Lagos names if applicable)
    psbList = Array()                 

    ' Hub Officer list (Update with Lagos hub officers if applicable)
    hubOfficerList = Array()

    ' List of states / region keywords to ignore
    statesList = Array("Lagos", "Lagos State")

    ' Loop through each cell in column A
    For Each stateCell In stateColumn
        state = Trim(stateCell.Value) ' Clean up leading/trailing spaces
        category = "" ' Reset category

        ' Skip categorization if value is a state in statesList
        If Not IsState(state, statesList) Then
            ' Priority: HUB OFFICER > RIDER > ADHOC > IN HOUSE CLEANER > IPDS > PSB
            If IsInList(state, hubOfficerList) Then
                category = "HUB OFFICER"
            ElseIf IsInList(state, ridersList) Then
                category = "RIDER"
            ElseIf IsInList(state, adhocList) Then
                category = "ADHOC"
            ElseIf IsInList(state, cleanerList) Then
                category = "IN HOUSE CLEANER"
            ElseIf IsInList(state, ipdsList) Then
                category = "IPDS"
            ElseIf IsInList(state, psbList) Then
                category = "PSB"
            Else
                category = "Not In List"
            End If
        End If

        ' Write result to Column K (Offset by 10 columns from Column A)
        If category <> "" Then
            stateCell.Offset(0, 10).Value = category
        End If
    Next stateCell

    MsgBox "CATEGORIZATION COMPLETED!"
End Sub

' =====================================================
' Function: IsState
' Purpose : Checks if a given value contains any known state names
' Returns : True if match found; otherwise False
' =====================================================
Function IsState(state As String, statesList As Variant) As Boolean
    Dim i As Integer
    If (Not statesList) = -1 Then Exit Function
    For i = LBound(statesList) To UBound(statesList)
        If InStr(1, state, statesList(i), vbTextCompare) > 0 Then
            IsState = True
            Exit Function
        End If
    Next i
    IsState = False
End Function

' =====================================================
' Function: IsInList
' Purpose : Checks if a given name contains any string in the provided list
' Returns : True if match found; otherwise False
' =====================================================
Function IsInList(name As String, nameList As Variant) As Boolean
    Dim i As Integer
    If (Not nameList) = -1 Then Exit Function
    For i = LBound(nameList) To UBound(nameList)
        If InStr(1, name, nameList(i), vbTextCompare) > 0 Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function