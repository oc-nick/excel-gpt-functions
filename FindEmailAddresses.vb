' Intended to be used within TRANSPOSE() to spill multiple addresses to the right
Function FindEmailAddresses(text As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result() As Variant
    Dim i As Long
    
    ' Create a new RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern for email addresses (a simple pattern to match most common email formats)
    regex.Pattern = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
    
    ' Set the Global property to True to find all matches in the text
    regex.Global = True
    
    ' Set the ignore case property to True for case-insensitive matching
    regex.ignorecase = True
    
    ' Get all matches in the text
    Set matches = regex.Execute(text)
    
    ' Resize the result array to the number of matches
    ReDim result(1 To matches.Count, 1 To 1)
    
    ' Loop through each match and add the email addresses to the result array
    For Each match In matches
        i = i + 1
        result(i, 1) = match.Value
    Next match
    
    ' Return the result array
    FindEmailAddresses = result
End Function
