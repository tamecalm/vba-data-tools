' ====================================================================================
' Subroutine: CategorizeRidersAdhocAndHubOfficers
' Purpose   : Categorizes entries in column A of the "Pivot (Old)" sheet into
'             HUB OFFICER, RIDER, ADHOC, or Not In List, based on name matching
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
    Dim statesList As Variant
    Dim category As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Pivot (Old)")

    ' Set the range for the column to be checked (assumes names in Column A)
    Set stateColumn = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' ----------------------------
    ' Populate your own name lists
    ' ----------------------------

    ' Adhoc list (example: replace with actual names)
    adhocList = Array("Abubakar Ishaq", "DEBORAH FAITH OMOSEYE", "Ibrahim Isyaku", "Otukhagua Benjamin", "Ojeikhoa Emmanuel Abiodun", _
                      "Shitu Aliyu", "Sani Abdurahman", "Ubangida Abdurrahman", "PETER STEPHEN", "Sani Aminu", "Mahmud   Ibrahim", _
                      "Yahaya Ibrahim", "Yakubu Ibrahim", "Andrew Vawe", "Shuaibu Saidu", "Yusuf Nasirdeen", "Usman Yahaya", "Lukman Ishaq")


    ipdsList = Array("ABDUSSAMAD MUHAMMAD ALMAJIR", "Abel John")         

    ' Hub Officer list (takes priority if name appears in multiple lists)
    hubOfficerList = Array("UDUMA EMMANUEL", "DANLANDI EUGENE", "AMINU AUDU", "PETER DANKARO", "INNOCENT  SIMON", _
                           "BUKAR LAWAL", "TOYOSI ADEYEMI", "SOMI KADIRI", "HARUNA YELWA", "HARUNA ABDULLAHI", _
                           "GABRIEL GEORGE", "GIDEON DANIEL", "RAKIYA MUSA", "SULE UMARU", "HASHIMA ADAM", _
                           "DANLADI EUGENE", "Abdullahi Haruna", _
                           "AUSTIN MODI", "ABDULSALAR SULEIMAN", "SULEIMAN AMBALI", "ABDULSALAM SULEIMAN")

    ' Rider list (should exclude names already in hubOfficerList)
    ridersList = Array("ABUBAKAR ISAH", "ABUBAKAR SARKI", "ALIYU MUHAMMED", "OLOCHE NGBEDE", "Yusuf Sunday", _
                       "HASSAN SULEIMAN", "Ibrahim Mohammed", "IBRAHIM UMAR", "IKELLE NAILS", "James Godwin", _
                       "KURAH IBRAHIM", "Mark Misi", "PAUL IKYOOR", "Sadiq Maidugu", "Suleiman Abdullahi", "Sadiq Mijinyawa", _
                       "KUM TIMOTHY", "MATHIAS GBASONGON", "Musa Wada", "RAYMOND SAMUEL", "PHILIP OLORUNSAIYE", _
                       "EMMANUEL DANLADI", "UMAR IBN ISAH", "ABDULLAHI AHMED USMAN", "Yusuf Isah", "Rotimi Elisha", "Christian Odeh", _
                       "ALEXANDER MATHIAS", "BINCHAK BINTUR NANKPAK", "KAPCHANG DANLADI", "Hassan Ali Gambo", "AMANG FRANCIS FELIX", _
                       "SANI ALHAJI BAKARI", "Umoru Odilihi", "ZAHARADDEEN ADAMU", "Usman Buhari", "LAWAL LURWANU")

    ' List of states to ignore
    statesList = Array("Benue", "FCT", "Kaduna", "Kogi", "Nasarawa", "Niger", _
                       "Adamawa", "Bauchi", "Borno", "Gombe", "Plateau", "Taraba", _
                       "Yobe", "Jigawa", "Kano", "Katsina", "Kebbi", "Sokoto", "Zamfara")

    ' Loop through each cell in column A
    For Each stateCell In stateColumn
        state = Trim(stateCell.Value) ' Clean up leading/trailing spaces
        category = "" ' Reset category

        ' Skip categorization if value is a state in statesList
        If Not IsState(state, statesList) Then
            ' Priority: HUB OFFICER > RIDER > ADHOC
            If IsInList(state, hubOfficerList) Then
                category = "HUB OFFICER"
            ElseIf IsInList(state, ridersList) Then
                category = "RIDER"
            ElseIf IsInList(state, adhocList) Then
                category = "ADHOC"
            Else
                category = "Not In List"
            End If
        End If

        ' Write result to Column K (Offset by 10 columns from Column A)
        If category <> "" Then
            stateCell.Offset(0, 10).Value = category
        End If
    Next stateCell

    MsgBox "CATEGORIZATION COMPLETED FOR HUB OFFICERS, RIDERS, AND ADHOC!"
End Sub

' =====================================================
' Function: IsState
' Purpose : Checks if a given value contains any known state names
' Returns : True if match found; otherwise False
' =====================================================
Function IsState(state As String, statesList As Variant) As Boolean
    Dim i As Integer
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
    For i = LBound(nameList) To UBound(nameList)
        If InStr(1, name, nameList(i), vbTextCompare) > 0 Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function


