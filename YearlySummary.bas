Option Compare Database
Option Explicit ' Ensure all variables are declared

' remember to IMPORT EXCEL OBJECT LIBRARY
' Tools > References > Microsoft excel etc

Sub ExportYearlySummary()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim xlApp As Excel.Application
    Dim includeSites As VbMsgBoxResult
    Dim allYears As VbMsgBoxResult
    Dim exportPath As String
    Dim yearFilter As String

    ' Ask the user if they want to include the "Sites" sheet
    includeSites = MsgBox("Do you want to add tbl_Sites?", vbQuestion + vbYesNo)

    ' Ask the user if they want to export all years
    allYears = MsgBox("Do you want to export data for all years?", vbQuestion + vbYesNo)

    ' Get the current database
    Set db = CurrentDb()

    ' Define the export path
    exportPath = CurrentProject.Path & "\ExportsFromAccess\YearlySummaries\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    ' Create an Excel application object
    Set xlApp = New Excel.Application
    xlApp.Visible = False ' Hide Excel application
    xlApp.ScreenUpdating = False ' Disable screen updating to improve performance

    If allYears = vbYes Then
        ' Determine the range of years available in your data
        Dim rsYears As DAO.Recordset
        Set rsYears = db.OpenRecordset("SELECT DISTINCT Year([DateTime]) AS Year FROM tbl_HourlyData_Archived19952017 " & _
                                        "UNION SELECT DISTINCT Year([DateTime]) AS Year FROM tbl_HourlyData_CurrentMaster2018ToPresent " & _
                                        "ORDER BY Year", dbOpenSnapshot)

        ' Loop through each year and perform the export process
        Do While Not rsYears.EOF
            If Not IsNull(rsYears!Year) Then
                yearFilter = CStr(rsYears!Year)
                Call ExportYearlyData(db, xlApp, yearFilter, includeSites, exportPath)
            End If
            rsYears.MoveNext
        Loop
        rsYears.Close
        Set rsYears = Nothing
    Else
        ' Prompt the user for the year if not exporting all years
        yearFilter = InputBox("Enter a year for the Continuous WQ Summary:")
        If yearFilter <> "" Then
            Call ExportYearlyData(db, xlApp, yearFilter, includeSites, exportPath)
        Else
            MsgBox "No year was entered. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End If

    ' Notify the user
    MsgBox "Export complete.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    ' Perform any cleanup operations here
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
End Sub

Sub ExportYearlyData(db As DAO.Database, xlApp As Excel.Application, yearFilter As String, includeSites As VbMsgBoxResult, exportPath As String)
    ' Ensure yearFilter is not empty before proceeding
    If yearFilter = "" Then
        MsgBox "Year filter is empty. Cannot proceed with export.", vbCritical
        Exit Sub
    End If

    Dim rs As DAO.Recordset
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim sqlQuery As String
    Dim fileName As String
    Dim lastRow As Long

    ' Construct the SQL query for the specified year
    sqlQuery = "SELECT SiteCode, DateTime, TempC, Matrix FROM tbl_HourlyData_Archived19952017 WHERE Year([DateTime]) = " & yearFilter
    sqlQuery = sqlQuery & " UNION ALL SELECT SiteCode, DateTime, TempC, Matrix FROM tbl_HourlyData_CurrentMaster2018ToPresent WHERE Year([DateTime]) = " & yearFilter

    ' Open the recordset
    Set rs = db.OpenRecordset(sqlQuery, dbOpenSnapshot) ' Use dbOpenSnapshot for read-only access

    ' Add a new workbook
    Set xlBook = xlApp.Workbooks.Add

    ' Reference the first sheet for data export
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.name = "WQ Summary " & yearFilter

    ' Write headers to the first row
    xlSheet.Cells(1, 1).Value = "SiteCode"
    xlSheet.Cells(1, 2).Value = "DateTime"
    xlSheet.Cells(1, 3).Value = "TempC"
    xlSheet.Cells(1, 4).Value = "Matrix"

    ' Copy the data from the recordset to the worksheet, starting from the second row
    xlSheet.Cells(2, 1).CopyFromRecordset rs

    ' Format the DateTime column
    xlSheet.Columns("B:B").NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"

    ' Sort the entire data range based on the DateTime column, then SiteCode
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "B").End(Excel.xlUp).Row

    ' First sort by DateTime
    With xlSheet.Range("A1:D" & lastRow)
        .Sort Key1:=xlSheet.Range("B2:B" & lastRow), Order1:=Excel.xlAscending, _
              Orientation:=Excel.xlTopToBottom, Header:=Excel.xlYes
    End With

    ' Then sort by SiteCode while maintaining the DateTime sort
    With xlSheet.Range("A1:D" & lastRow)
        .Sort Key1:=xlSheet.Range("A2:A" & lastRow), Order1:=Excel.xlAscending, _
              Key2:=xlSheet.Range("B2:B" & lastRow), Order2:=Excel.xlAscending, _
              Orientation:=Excel.xlTopToBottom, Header:=Excel.xlYes
    End With

    ' Apply filter to the data range
    xlSheet.Range("A1:D" & lastRow).AutoFilter

    ' If the user chose to include the "Sites" sheet, add it
    If includeSites = vbYes Then
        Dim xlSheetSites As Excel.Worksheet
        Set xlSheetSites = xlBook.Sheets.Add(After:=xlBook.Sheets(xlBook.Sheets.Count))
        xlSheetSites.name = "Sites"
        Dim rsSites As DAO.Recordset
        Set rsSites = db.OpenRecordset("tbl_Sites", dbOpenSnapshot)
        xlSheetSites.Cells(1, 1).CopyFromRecordset rsSites
        rsSites.Close
        Set rsSites = Nothing
    End If

    ' Define the filename
    fileName = yearFilter & "_SRRC_Continuous_WQ_Summary" & ".xlsx"

    ' Save and close the Excel workbook
    xlBook.SaveAs exportPath & fileName
    xlBook.Close SaveChanges:=False

    ' Clean up
    rs.Close
    Set rs = Nothing
    Set xlBook = Nothing

    ' Re-enable screen updating
    xlApp.ScreenUpdating = True
End Sub
