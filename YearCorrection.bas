Option Compare Database

Sub UpdateSiteYears()
    Dim db As DAO.Database
    Dim rsSites As DAO.Recordset
    Dim rsYears As DAO.Recordset
    Dim siteCode As String
    Dim yearList As String
    Dim sqlQuery As String
    
    Set db = CurrentDb()
    
    ' Open the recordset for tbl_Sites
    Set rsSites = db.OpenRecordset("tbl_Sites", dbOpenDynaset, dbSeeChanges)
    
    ' Loop through each site in tbl_Sites
    Do While Not rsSites.EOF
        siteCode = rsSites!siteCode
        yearList = ""
        
        ' Construct the SQL query to get distinct years for the current site from both tables
        sqlQuery = "SELECT DISTINCT Year([DateTime]) AS Year FROM tbl_HourlyData_Archived19952017 WHERE SiteCode = '" & siteCode & "' " & _
                   "UNION SELECT DISTINCT Year([DateTime]) AS Year FROM tbl_HourlyData_CurrentMaster2018ToPresent WHERE SiteCode = '" & siteCode & "' " & _
                   "ORDER BY Year"
        
        ' Open the recordset for the years query
        Set rsYears = db.OpenRecordset(sqlQuery, dbOpenSnapshot)
        
        ' Loop through each year and build the yearList string
        Do While Not rsYears.EOF
            If yearList = "" Then
                yearList = CStr(rsYears!Year)
            Else
                yearList = yearList & ", " & CStr(rsYears!Year)
            End If
            rsYears.MoveNext
        Loop
        
        ' Update the Years column for the current site
        rsSites.Edit
        rsSites!Years = yearList
        rsSites.Update
        
        ' Move to the next site
        rsSites.MoveNext
        
        ' Clean up the years recordset
        rsYears.Close
        Set rsYears = Nothing
    Loop
    
    ' Clean up
    rsSites.Close
    Set rsSites = Nothing
    Set db = Nothing
    
    MsgBox "Site years updated successfully.", vbInformation
End Sub

