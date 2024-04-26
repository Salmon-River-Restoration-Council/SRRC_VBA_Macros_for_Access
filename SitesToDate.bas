Option Compare Database

Sub ExportSitesToDate()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim siteCode As String
    Dim yearFilter As String
    Dim sqlQuery As String
    Dim exportPath As String
    Dim fileName As String
    Dim wsName As String
    Dim i As Integer
    Dim lastCol As Integer
    Dim pareColumns As String
    Dim siteCodes As Variant
    Dim code As Variant

    ' Prompt the user for input
    siteCode = InputBox("Enter a site code or 'All' for all sites:")
    yearFilter = InputBox("Enter a year or 'All' for all years:")
    pareColumns = InputBox("Would you like to pare columns A, C, D, G, H, and I? Enter 'Yes' to pare:")

    ' Get the current database
    Set db = CurrentDb()

    ' Define the export path
    exportPath = CurrentProject.Path & "\ExportsFromAccess\SitesToDate\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    ' Create an Excel application object
    Set xlApp = CreateObject("Excel.Application")

    If siteCode = "All" Then
        ' Retrieve all distinct site codes from the database
        Set rs = db.OpenRecordset("SELECT DISTINCT [SiteCode] FROM tbl_HourlyData_Archived19952017 WHERE [SiteCode] Is Not Null AND [SiteCode] <> '' UNION SELECT DISTINCT [SiteCode] FROM tbl_HourlyData_CurrentMaster2018ToPresent WHERE [SiteCode] Is Not Null AND [SiteCode] <> ''", dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            siteCodes = rs.GetRows(rs.RecordCount)
        End If
        rs.Close
    Else
        ' If a specific site code is provided, create an array with just that code
        siteCodes = Array(siteCode)
    End If

    For Each code In siteCodes
        If Not IsNull(code) And Trim(code) <> "" Then
            ' Sanitize the site code to create a valid worksheet name
            wsName = SanitizeWorksheetName(CStr(code))

            ' Construct the SQL query based on user input
            If yearFilter = "All" Then
                sqlQuery = "SELECT * FROM tbl_HourlyData_Archived19952017 UNION ALL SELECT * FROM tbl_HourlyData_CurrentMaster2018ToPresent"
            Else
                sqlQuery = "SELECT * FROM tbl_HourlyData_Archived19952017 WHERE Year([DateTime]) = " & yearFilter & _
                           " UNION ALL SELECT * FROM tbl_HourlyData_CurrentMaster2018ToPresent WHERE Year([DateTime]) = " & yearFilter
            End If

            If wsName <> "All" Then
                sqlQuery = "SELECT * FROM (" & sqlQuery & ") AS CombinedData WHERE [SiteCode] = '" & wsName & "'"
            End If

            ' Open the recordset
            Set rs = db.OpenRecordset(sqlQuery, dbOpenDynaset, dbSeeChanges)

            ' Continue only if the recordset is not empty
            If Not (rs.EOF And rs.BOF) Then
                ' Define the filename
                fileName = wsName
                If yearFilter <> "All" Then
                    fileName = fileName & "_" & yearFilter
                End If
                If pareColumns = "Yes" Then
                    ' only for testing. Add "_p" if necessary
                    fileName = fileName & ""
                End If
                fileName = fileName & ".xlsx"

                ' Add a new workbook
                Set xlBook = xlApp.Workbooks.Add

                ' Reference the first sheet
                Set xlSheet = xlBook.Sheets(1)

                ' Set the worksheet name
                xlSheet.Name = wsName

                ' Write headers to the first row
                For i = 0 To rs.Fields.Count - 1
                    xlSheet.Cells(1, i + 1).Value = rs.Fields(i).Name
                Next i

                ' Copy the data from the recordset to the worksheet, starting from the second row
                xlSheet.Cells(2, 1).CopyFromRecordset rs

                ' Find the last column with data
                lastCol = xlSheet.Cells(1, xlSheet.Columns.Count).End(-4159).Column ' -4159 corresponds to xlToLeft

                ' Sort the entire data range based on the date column (assumed to be column E)
                With xlSheet
                    .Range(.Cells(1, 1), .Cells(.Cells(.Rows.Count, "E").End(-4162).Row, lastCol)).Sort Key1:=.Range("E2"), Order1:=xlAscending, Header:=xlYes ' -4162 corresponds to xlUp
                    .Columns("E:E").NumberFormat = "MM/DD/YYYY HH:MM AM/PM"
                    .Rows(1).AutoFilter
                    .Cells.EntireColumn.AutoFit
                End With

                ' Pare columns if chosen by the user, after sorting and before saving
                If pareColumns = "Yes" Then
                    Dim colsToPare As Variant
                    colsToPare = Array("I", "H", "G", "D", "C", "A") ' Reverse order to avoid shifting
                    Dim col As Variant
                    For Each col In colsToPare
                        xlSheet.Columns(col & ":" & col).Delete
                    Next col
                End If

                ' Turn off alerts to overwrite existing files without asking
                xlApp.DisplayAlerts = False

                ' Save and close the Excel workbook
                xlBook.SaveAs exportPath & fileName
                xlBook.Close SaveChanges:=False

                ' Turn alerts back on
                xlApp.DisplayAlerts = True
            End If

            ' Clean up the recordset
            rs.Close
        End If
    Next code

    ' Clean up
    Set rs = Nothing
    db.Close
    Set db = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    ' Notify the user
    MsgBox "Export complete.", vbInformation

End Sub

Function SanitizeWorksheetName(name As String) As String
    ' This function removes invalid characters and trims the worksheet name to a valid length
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", "*", "?", ":", "[", "]")
    Dim char As Variant
    
    For Each char In invalidChars
        name = Replace(name, char, "")
    Next char
    
    ' Trim the name to the maximum length if necessary
    If Len(name) > 31 Then
        name = Left(name, 31)
    End If
    
    ' Ensure the name is not empty
    If Len(name) = 0 Then
        name = "Sheet"
    End If
    
    SanitizeWorksheetName = name
End Function