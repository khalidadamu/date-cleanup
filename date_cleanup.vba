Sub CleanUpDates()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dateValue As Variant
    
    ' Set the worksheet where your data is located  Replace "Sheet1" with your sheet name
    Set ws = ThisWorkbook.Sheets("Sheet_name") 
    
    ' Assuming your date data is in column H, adjust accordingly if it's in a different column
    Set rng = ws.Range("H:H")
    
    ' Convert referenced cells to their values
    rng.Value = rng.Value
    
    ' Loop through each cell in the specified range
    For Each cell In rng
        ' Check if the cell is not empty and contains text
        If Not IsEmpty(cell) And IsDate(cell.Value) Then
            ' Attempt to convert text representations of dates to Excel date values using DATEVALUE function
            On Error Resume Next 
            cell.Value = Application.WorksheetFunction.dateValue(cell.Value)
            On Error GoTo 0 
            
            ' Check if the conversion was successful
            If IsDate(cell.Value) Then
                ' Format the cell as a date
                cell.NumberFormat = "mm/dd/yyyy" 
            Else
                ' Display a message if the conversion failed
                MsgBox "Unable to clean date in cell " & cell.Address
            End If
        ElseIf Not IsEmpty(cell) Then
            ' If the cell is not empty but doesn't contain a date, attempt further cleaning
            ' Remove leading and trailing spaces using TRIM function
            cell.Value = Application.WorksheetFunction.Trim(cell.Value)
            
            ' Extract date portion from within text using custom function ExtractDate
            dateValue = ExtractDate(cell.Value)
            
            ' Check if a valid date value was extracted
            If IsDate(dateValue) Then
                ' Format the cell as a date
                cell.Value = dateValue
                cell.NumberFormat = "mm/dd/yyyy" ' Adjust the date format as needed
            Else
                ' Display a message if a valid date couldn't be extracted
                MsgBox "Unable to clean date in cell " & cell.Address
            End If
        End If
    Next cell
End Sub

Function ExtractDate(ByVal text As String) As Variant
    ' Custom function to extract date portion from within text
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim dateString As String
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the pattern to match date-like strings
    regex.Pattern = "\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b"
    
    ' Find all matches of the pattern in the text
    Set matches = regex.Execute(text)
    
    ' Iterate through the matches and return the first valid date found
    For Each match In matches
        dateString = match.Value
        ' Attempt to convert the matched string to a date value
        If IsDate(dateString) Then
            ExtractDate = CDate(dateString)
            Exit Function ' Exit function after finding the first valid date
        End If
    Next match
    
    ' Return an error value if no valid date was found
    ExtractDate = CVErr(xlErrValue)
End Function

