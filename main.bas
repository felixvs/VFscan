Attribute VB_Name = "main"
'Main program
Sub main()
    Dim WK As Workbook, i As Long, cells_data As String, cells_name As String
    Set WK = ThisWorkbook
    
    'Create the dictionay
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    'Preposition sheets cells counts
    prepositions_lastrow = WK.Worksheets("prepositions").Range("A" & Rows.Count).End(xlUp).Row

    'Populated the dictionay with the prepositions sheet data
    For i = 2 To prepositions_lastrow
        dict.Add Key:=(WK.Worksheets("prepositions").Range("A" & i).Value), Item:=WK.Worksheets("prepositions").Range("B" & i).Value
    Next i
    'VF sheets cells counts
    VF_lastrow = WK.Worksheets("VF").Range("A" & Rows.Count).End(xlUp).Row

    'Calling the replace prepositions funcion an writing the data in the AR column
    For i = 2 To VF_lastrow
        'replace preposition
        prepositions_name = replace_prepositions(WK.Worksheets("VF").Range("A" & i).Value, dict)

        'If prepositions_name is different of empry
        If prepositions_name <> vbNullString Then
            WK.Worksheets("VF").Range("A" & i).Value = prepositions_name
            WK.Worksheets("VF").Range("A" & i).Interior.Color = vbYellow
        End If
    Next i

    'Destroy a dictionary
    Set dict = Nothing
    
    lastrow = WK.Worksheets("VF").Range("A" & Rows.Count).End(xlUp).Row
    
    'Lower Case Words : Hightlight lower case words in the vivial force file.
    For i = lastrow To 1 Step -1

        cells_data = WK.Worksheets("VF").Range("A" & i).Value2

        cells_data = proper_case(cells_data)

        If cells_data = "lower" Then
            Worksheets("VF").Range("A" & i).Interior.Color = vbYellow
        End If
    Next i
    
    'Designations : Looking for designations that were wrongly placed.
    For i = lastrow To 1 Step -1

        cells_name = WK.Worksheets("VF").Range("A" & i).Value2
        'replace preposition
        cells_name = designation_correct_place(cells_name)

        If cells_name = "found" Then
            Worksheets("VF").Range("A" & i).Interior.Color = vbYellow
        End If
    Next i
    
End Sub
