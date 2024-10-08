Option Compare Database

Sub ReplaceSiteCodeAndLogChanges()
    Dim db As DAO.Database
    Dim tblNames As Variant
    Dim findWord As String
    Dim replaceWord As String
    Dim i As Integer
    Dim sqlUpdate As String
    Dim rs As DAO.Recordset
    Dim csvContent As String
    Dim csvPath As String
    Dim fNum As Integer
    Dim currentDate As String
    Dim userResponse As Integer
    
    ' Warning to the user
    userResponse = MsgBox("Have you changed the necessary backfill information before using this script?", vbYesNo + vbQuestion, "Confirm Information Change")
    If userResponse = vbNo Then
        MsgBox "Please update the necessary information before proceeding.", vbExclamation, "Action Required"
        Exit Sub
    End If
    
    ' Initialize database
    Set db = CurrentDb()
    
    ' Tables to be accessed
    ' CHANGE AS NEEDED!
    ' EXTRA TABLES CAN BE ADDED
    tblNames = Array("tbl_AnnualStats_Archived19952017", "tbl_DailyStats_Archived19952017", "tbl_HourlyData_Archived19952017")
    
    ' Prompt for words
    findWord = InputBox("Enter the word to find:", "Find Word")
    replaceWord = InputBox("Enter the replacement word:", "Replace Word")
    
    ' Initialize CSV content
    csvContent = "Table,Old SiteCode,New SiteCode" & vbCrLf
    
    ' Loop through each table
    For i = 0 To UBound(tblNames)
        ' Update query
        sqlUpdate = "UPDATE " & tblNames(i) & " SET SiteCode = Replace(SiteCode, '" & findWord & "', '" & replaceWord & "') WHERE SiteCode LIKE '*" & findWord & "*'"
        
        ' Execute update
        db.Execute sqlUpdate, dbFailOnError
        
        ' Log changes
        Set rs = db.OpenRecordset("SELECT SiteCode FROM " & tblNames(i) & " WHERE SiteCode LIKE '*" & replaceWord & "*'", dbOpenSnapshot)
        Do While Not rs.EOF
            csvContent = csvContent & tblNames(i) & "," & findWord & "," & rs!SiteCode & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    Next i
    
    ' Current date in YYYY-MM-DD_HHNN format for better readability
    currentDate = Format(Now, "yyyy-mm-dd_hhnn")
    
    ' Save changes to CSV, including the current date in the file name
    csvPath = Application.CurrentProject.Path & "\SiteCodeChanges_" & currentDate & ".csv"
    fNum = FreeFile
    Open csvPath For Output As #fNum
    Print #fNum, csvContent
    Close #fNum
    
    ' Clean up
    Set db = Nothing
    
    MsgBox "SiteCode replacements and logging completed. CSV saved to: " & csvPath, vbInformation
End Sub