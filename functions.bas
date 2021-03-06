Attribute VB_Name = "functions"
'Find words that arent proper case trough the VF documents
Function proper_case(cells As String) As String
    Dim arr As Variant
    'Replace dash for a space
    cells = Replace(cells, "-", " ")
    
    'Convert the name in an array
    arr = Split(cells, " ")
    
    'Check if the first letter of each word is lower case.
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Left(arr(i), 1) Like "*[a-z]*" Then
            proper_case = "lower"
            Exit For
        End If
    Next i    
End Function

'Replace prepositions in the VF file
Function replace_prepositions(nombre As String, dict As Object) As String
    For Each Key In dict.keys
        
        'If there is a key in the name, replace it with the item of the key.
        If InStr(nombre, Key) <> 0 Then
            replace_prepositions = Replace(nombre, Key, dict(Key), 1)
            Exit For
        End If
    Next Key
End Function

'Looking for designations that were wrongly placed.
Function designation_correct_place(nombre As String) As String
    ' Create designation array
    Dim arr() As Variant
    'Arrays are key sensitive in VBA, please add the designation accordingly.
    arr = Array("Jr", "Mr", "Mrs", "Dr", "Ii", "Iii", "Sr")
    'If the first three characters of the name is equal to the array.
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Trim(Left(nombre, 3)) = arr(i) Then
            MsgBox "hello"
            designation_correct_place = "found"
        End If
    Next i
End Function
