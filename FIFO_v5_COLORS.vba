Option Explicit

' ========================================
' FIFO ULTIMATE - ONE CLICK SOLUTION
' Combines VBA #1 (Master Generator) + VBA #2 (RSM Splitter)
' Run FIFO_ULTIMATE_OneClick() to do everything!
' ========================================

Public Sub FIFO_ULTIMATE_OneClick()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    Dim masterFile As String
    
    startTime = Timer
    
    MsgBox "Starting FIFO Ultimate Process..." & vbCrLf & vbCrLf & _
           "Step 1/2: Generating master report...", vbInformation, "FIFO Ultimate"
    
    ' STEP 1: Generate master report
    Call SegmentExpiryItems
    
    ' STEP 2: Find the generated master report
    Dim todayStamp As String
    todayStamp = Format(Date, "dd-mmm-yyyy")
    
    masterFile = "FIFO Expiry Report - " & todayStamp & ".xlsx"
    
    If Dir(ThisWorkbook.Path & "\" & masterFile) = "" Then
        masterFile = "FIFO_Expiry_Report_-_" & todayStamp & ".xlsx"
        
        If Dir(ThisWorkbook.Path & "\" & masterFile) = "" Then
            MsgBox "Could not find generated master report!" & vbCrLf & vbCrLf & _
                   "Looked for:" & vbCrLf & _
                   "- FIFO Expiry Report - " & todayStamp & ".xlsx" & vbCrLf & _
                   "- FIFO_Expiry_Report_-_" & todayStamp & ".xlsx", vbCritical
            GoTo CleanupAndExit
        End If
    End If
    
    MsgBox "Step 2/2: Splitting by RSM...", vbInformation, "FIFO Ultimate"
    
    Dim wbMaster As Workbook
    Set wbMaster = Workbooks.Open(ThisWorkbook.Path & "\" & masterFile)
    
    ' STEP 3: Split by RSM
    Call SplitFIFOByRSM_Complete_Internal(wbMaster)
    
    wbMaster.Close SaveChanges:=False
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    MsgBox "✅ FIFO Ultimate Complete!" & vbCrLf & vbCrLf & _
           "Master report generated" & vbCrLf & _
           "Split into RSM files" & vbCrLf & vbCrLf & _
           "Time: " & Format(elapsedTime, "0.0") & " seconds", vbInformation, "Success!"
    
    Exit Sub
    
CleanupAndExit:
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in FIFO Ultimate: " & Err.Description, vbCritical
End Sub

' ========================================
' VBA #1: MASTER REPORT GENERATOR
' ========================================

Sub SegmentExpiryItems()
    On Error GoTo ErrorHandler
    
    Dim wsSource As Worksheet
    Dim wsExpired As Worksheet, ws1Month As Worksheet, ws2Month As Worksheet
    Dim ws3Month As Worksheet, wsTotal As Worksheet
    Dim lastRow As Long, i As Long
    Dim expiryCol As Long
    Dim expiryValue As Variant
    Dim expiryDate As Date
    Dim todayDate As Date
    Dim daysRemaining As Long
    Dim sourceFilePath As String, outputFilePath As String
   
    Dim finalFileName As String
    
    Dim rowExpired As Long, row1Month As Long, row2Month As Long
    Dim row3Month As Long, rowTotal As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    todayDate = Date
    sourceFilePath = ThisWorkbook.Path
    
    ' Load SAP codes from SAP_Code.xlsx
    Dim wbSAP As Workbook, wsSAP As Worksheet
    Dim sapDict As Object
 
    Set sapDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wbSAP = Workbooks.Open(sourceFilePath & "\SAP Code VBA-0.xlsx")
    If Not wbSAP Is Nothing Then
        Set wsSAP = wbSAP.Sheets(1)
        Dim sapLastRow As Long
        sapLastRow = wsSAP.Cells(wsSAP.Rows.Count, 1).End(xlUp).Row
        
        Dim distName As String, sapCode As String
        
        For i = 2 To sapLastRow
            distName = Trim(CStr(wsSAP.Cells(i, 1).Value))
            sapCode = Trim(CStr(wsSAP.Cells(i, 2).Value))
            If Len(distName) > 0 And Len(sapCode) > 0 Then
                sapDict.Add distName, sapCode
            End If
        Next i
     
        wbSAP.Close SaveChanges:=False
    End If
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("Total")
    On Error GoTo ErrorHandler
    
    If wsSource Is Nothing Then
        MsgBox "Error: Could not find sheet named 'Total'.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    If wsSource.Cells(1, 1).Value = "" Then
        MsgBox "No data found in 'Total' sheet.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data to process!", vbExclamation
        GoTo CleanupAndExit
    End If
    
    expiryCol = 14
    
    Dim wbOutput As Workbook
    
    Set wbOutput = Workbooks.Add
    
    Application.DisplayAlerts = False
    Do While wbOutput.Worksheets.Count > 1
        wbOutput.Worksheets(wbOutput.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    Set wsExpired = wbOutput.Worksheets(1)
    wsExpired.Name = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    wsExpired.Tab.Color = RGB(255, 102, 102)
    
    Set ws1Month = wbOutput.Worksheets.Add(After:=wsExpired)
    ws1Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws1Month.Tab.Color = RGB(255, 153, 153)
    
    Set ws2Month = wbOutput.Worksheets.Add(After:=ws1Month)
    ws2Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    ws2Month.Tab.Color = RGB(255, 192, 128)
    
    Set ws3Month = wbOutput.Worksheets.Add(After:=ws2Month)
    ws3Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws3Month.Tab.Color = RGB(255, 230, 153)
    
    Set wsTotal = wbOutput.Worksheets.Add(After:=ws3Month)
    wsTotal.Name = ChrW(&H6A9) & ChrW(&H627) & ChrW(&H645) & ChrW(&H644)
    wsTotal.Tab.Color = RGB(146, 208, 80)
    
    Call SetupReorganizedHeaders(wsExpired, sapDict)
    Call SetupReorganizedHeaders(ws1Month, sapDict)
    Call SetupReorganizedHeaders(ws2Month, sapDict)
    Call SetupReorganizedHeaders(ws3Month, sapDict)
    Call SetupReorganizedHeaders(wsTotal, sapDict)
    
    rowExpired = 2
    row1Month = 2
    row2Month = 2
    row3Month = 2
    rowTotal = 2
    
    For i = 2 To lastRow
        ' FIXED: Category moved to column 20 (ItemValue added at col 18) instead of recalculating
        Dim categoryValue As String
        categoryValue = Trim(CStr(wsSource.Cells(i, 20).Value))
        
        ' Get expiry date and calculate days for display
        expiryValue = wsSource.Cells(i, expiryCol).Value
        daysRemaining = 999
        
        If expiryValue <> "" And Not IsEmpty(expiryValue) Then
            On Error Resume Next
            expiryDate = DateValue(expiryValue)
            daysRemaining = expiryDate - todayDate
            On Error GoTo ErrorHandler
        End If
        
        ' Categorize using existing Category column (matches Summary!)
        Select Case categoryValue
            Case "Expired"
                Call CopyReorganizedRow(wsSource, wsExpired, i, rowExpired, daysRemaining, sapDict)
                rowExpired = rowExpired + 1
            Case "Less than 1 Month"
                Call CopyReorganizedRow(wsSource, ws1Month, i, row1Month, daysRemaining, sapDict)
                row1Month = row1Month + 1
            Case "Less than 2 Months"
                Call CopyReorganizedRow(wsSource, ws2Month, i, row2Month, daysRemaining, sapDict)
                row2Month = row2Month + 1
            Case "Less than 3 Months"
                Call CopyReorganizedRow(wsSource, ws3Month, i, row3Month, daysRemaining, sapDict)
                row3Month = row3Month + 1
        End Select
        
        ' Copy to Complete sheet if categorized
        If categoryValue <> "" Then
            Call CopyReorganizedRow(wsSource, wsTotal, i, rowTotal, daysRemaining, sapDict)
            rowTotal = rowTotal + 1
        End If
    Next i
    
    Dim ws As Worksheet
    Dim lastRowSheet As Long
    
    For Each ws In wbOutput.Worksheets
        Dim skipSheet As Boolean
    
        skipSheet = False
        
        If InStr(ws.Name, "10") > 0 Or InStr(ws.Name, "RSM") > 0 Or InStr(ws.Name, ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621)) > 0 Or InStr(ws.Name, ChrW(&H645) & ChrW(&H644) & ChrW(&H62E) & ChrW(&H635)) > 0 Or InStr(ws.Name, "Pivot") > 0 Then
            skipSheet = True
        End If
        
        If Not skipSheet Then
            With ws
                lastRowSheet = .Cells(.Rows.Count, 1).End(xlUp).Row
                
                If lastRowSheet > 1 Then
                    .Sort.SortFields.Clear
                     .Sort.SortFields.Add Key:=.Range("F2:F" & lastRowSheet), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    With .Sort
                        .SetRange ws.Range("A1:I" & lastRowSheet)
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                
                Dim totalQtySum As Long
                totalQtySum = Application.WorksheetFunction.Sum(.Range("F2:F" & lastRowSheet))
         
                lastRowSheet = lastRowSheet + 1
                
                ' Clear all cells in TOTAL row first
                .Range("A" & lastRowSheet & ":I" & lastRowSheet).ClearContents
                
                ' Set only the cells we need
                .Cells(lastRowSheet, 1).Value = ""
                .Cells(lastRowSheet, 6).Value = totalQtySum
                .Cells(lastRowSheet, 6).NumberFormat = "#,##0"
                .Cells(lastRowSheet, 6).Font.Bold = True
                .Cells(lastRowSheet, 6).Interior.Color = RGB(146, 208, 80)
                .Cells(lastRowSheet, 6).HorizontalAlignment = xlCenter
                .Cells(lastRowSheet, 6).VerticalAlignment = xlCenter
                .Cells(lastRowSheet, 6).Borders.LineStyle = xlContinuous
                .Cells(lastRowSheet, 6).Borders.Weight = xlMedium
                .Cells(lastRowSheet, 6).Borders.Color = RGB(0, 0, 0)
                
                ' NEW: Add ItemValue total (column 9)
                Dim totalValueSum As Double
                totalValueSum = Application.WorksheetFunction.Sum(.Range("I2:I" & (lastRowSheet - 1)))
                .Cells(lastRowSheet, 9).Value = totalValueSum
                .Cells(lastRowSheet, 9).NumberFormat = "#,##0"
                .Cells(lastRowSheet, 9).Font.Bold = True
                .Cells(lastRowSheet, 9).Interior.Color = RGB(146, 208, 80)
                .Cells(lastRowSheet, 9).HorizontalAlignment = xlCenter
                .Cells(lastRowSheet, 9).VerticalAlignment = xlCenter
                .Cells(lastRowSheet, 9).Borders.LineStyle = xlContinuous
                .Cells(lastRowSheet, 9).Borders.Weight = xlMedium
                .Cells(lastRowSheet, 9).Borders.Color = RGB(0, 0, 0)
                
                .Range("A" & lastRowSheet & ":E" & lastRowSheet).Interior.ColorIndex = xlNone
                .Range("A" & lastRowSheet & ":E" & lastRowSheet).Borders.LineStyle = xlNone
                .Range("G" & lastRowSheet & ":H" & lastRowSheet).Interior.ColorIndex = xlNone
                .Range("G" & lastRowSheet & ":H" & lastRowSheet).Borders.LineStyle = xlNone
                
                With .Range(.Cells(1, 1), .Cells(lastRowSheet - 1, 9))
                    .Font.Name = "Calibri"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .ShrinkToFit = False
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(200, 200, 200)
                End With
                
                With .Range("A1:I1")
                    .Font.Name = "Calibri"
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .ShrinkToFit = False
                    .Interior.Color = RGB(68, 114, 196)
                End With
                
                .Cells(1, 6).Interior.Color = RGB(255, 255, 0)
                .Cells(1, 6).Font.Color = RGB(0, 0, 0)
                .Cells(1, 6).Font.Bold = True
                
                .Cells(1, 8).Interior.Color = RGB(255, 165, 0)
                .Cells(1, 8).Font.Color = RGB(0, 0, 0)
                .Cells(1, 8).Font.Bold = True
                
                ' NEW: Highlight Value header (green like total)
                .Cells(1, 9).Interior.Color = RGB(146, 208, 80)
                .Cells(1, 9).Font.Color = RGB(0, 0, 0)
                .Cells(1, 9).Font.Bold = True
                
                Dim col As Long
                For col = 1 To 9
                    .Columns(col).AutoFit
                    .Columns(col).ColumnWidth = .Columns(col).ColumnWidth + 3
                Next col
                 
                .Cells.EntireRow.AutoFit
                .Cells.EntireColumn.AutoFit
                
                .Range("A1").AutoFilter
                .Activate
                .Range("A2").Select
                ActiveWindow.FreezePanes = True
                .Range("A1").Select
            End If
        End With
        End If
    Next ws
    
    ' CREATE GENERAL SUMMARY SHEET
    Dim wsGeneralSummary As Worksheet
    Set wsGeneralSummary = wbOutput.Worksheets.Add(Before:=wbOutput.Worksheets(1))
    wsGeneralSummary.Name = ChrW(&H645) & ChrW(&H644) & ChrW(&H62E) & ChrW(&H635) & " " & ChrW(&H639) & ChrW(&H627) & ChrW(&H645)
    wsGeneralSummary.Tab.Color = RGB(255, 192, 0)
    
    Call CreateGeneralSummarySheet(wsGeneralSummary, rowExpired - 2, row1Month - 2, row2Month - 2, row3Month - 2, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' CREATE TOP 10 STORES SHEET
    Dim wsSummary As Worksheet
    Set wsSummary = wbOutput.Worksheets.Add(Before:=wsExpired)
    wsSummary.Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621)
    wsSummary.Tab.Color = RGB(68, 114, 196)
    
    Call CreateTop10StoresOnly(wsSummary, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' CREATE TOP 10 PRODUCTS ANALYSIS
    Dim wsProductAnalysis As Worksheet
    Set wsProductAnalysis = wbOutput.Worksheets.Add(Before:=wsSummary)
    wsProductAnalysis.Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
    wsProductAnalysis.Tab.Color = RGB(76, 175, 80)
    
    Call CreateProductAnalysis(wsProductAnalysis, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' CREATE RSM DASHBOARD
    Dim wsDash As Worksheet
    Set wsDash = wbOutput.Worksheets.Add(Before:=wsProductAnalysis)
    wsDash.Name = ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H645) & ChrW(&H62F) & ChrW(&H631) & ChrW(&H627) & ChrW(&H621) & " RSM"
    wsDash.Tab.Color = RGB(192, 0, 0)
    
    Call CreateDashboard(wsDash, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' Activate General Summary as the first sheet
    wsGeneralSummary.Activate
    
    Dim baseFileName As String
    Dim fileExtension As String
    
    baseFileName = "FIFO Expiry Report - " & Format(Date, "dd-mmm-yyyy")
    fileExtension = ".xlsx"
    finalFileName = baseFileName & fileExtension
    outputFilePath = sourceFilePath & "\" & finalFileName
    
    ' Archive existing master file if it exists
    If Dir(outputFilePath) <> "" Then
        Dim archivePath As String
        Dim versionNum As Long
        Dim archiveFile As String
        Dim fso As Object
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Create archive subfolders
        archivePath = sourceFilePath & "\M. FIFO Archive"
        If Not fso.FolderExists(archivePath) Then fso.CreateFolder archivePath
        archivePath = archivePath & "\Archive [Outcome of 1st VBA]"
        If Not fso.FolderExists(archivePath) Then fso.CreateFolder archivePath
        
        versionNum = 1
        Do While fso.FileExists(archivePath & "\" & baseFileName & " v" & versionNum & fileExtension)
            versionNum = versionNum + 1
        Loop
        
        archiveFile = archivePath & "\" & baseFileName & " v" & versionNum & fileExtension
        
        On Error Resume Next
        fso.MoveFile outputFilePath, archiveFile
        On Error GoTo ErrorHandler
        
        Set fso = Nothing
    End If
    
    wbOutput.SaveAs Filename:=outputFilePath, FileFormat:=xlOpenXMLWorkbook
    
    On Error Resume Next
    wbOutput.Close SaveChanges:=False
    On Error GoTo ErrorHandler
    
    wbOutput.SaveAs Filename:=outputFilePath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
 
CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    On Error Resume Next
    If Err.Number = 0 Then
        MsgBox "Success!" & vbCrLf & vbCrLf & _
               "Saved as: " & finalFileName & vbCrLf & _
               "Location: " & sourceFilePath & vbCrLf & vbCrLf & _
               "SUMMARY:" & vbCrLf & _
               "• Expired: " & (rowExpired - 2) & vbCrLf & _
               "• Less than 1 month: " & (row1Month - 2) & vbCrLf & _
               "• Less than 2 months: " & (row2Month - 2) & vbCrLf & _
               "• Less than 3 months: " & (row3Month - 2), vbInformation
    End If
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    If Dir(outputFilePath) = "" Then
        MsgBox "Error: " & Err.Description, vbCritical
     Else
        MsgBox "Success!" & vbCrLf & vbCrLf & _
               "Saved as: " & finalFileName & vbCrLf & _
               "Location: " & sourceFilePath & vbCrLf & vbCrLf & _
               "SUMMARY:" & vbCrLf & _
               "• Expired: " & (rowExpired - 2) & vbCrLf & _
               "• Less than 1 month: " & (row1Month - 2) & vbCrLf & _
               "• Less than 2 months: " & (row2Month - 2) & vbCrLf & _
               "• Less than 3 months: " & (row3Month - 2), vbInformation
    End If
    On Error GoTo 0
End Sub

Private Sub SetupReorganizedHeaders(ws As Worksheet, sapDict As Object)
    ws.Cells.Font.Name = "Arial"
    
    ws.Cells(1, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
    ws.Cells(1, 2).Value = "SAP Code"
    ws.Cells(1, 3).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
    ws.Cells(1, 4).Value = ChrW(&H6A9) & ChrW(&H648) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(1, 5).Value = ChrW(&H625) & ChrW(&H633) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(1, 6).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629)
    ws.Cells(1, 7).Value = ChrW(&H62A) & ChrW(&H627) & ChrW(&H631) & ChrW(&H64A) & ChrW(&H62E) & " " & ChrW(&H627) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    ws.Cells(1, 8).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H623) & ChrW(&H64A) & ChrW(&H627) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H62A) & ChrW(&H628) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H629)
    ' NEW: Add ItemValue column (القیمة)
    ws.Cells(1, 9).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
End Sub

Private Sub CopyReorganizedRow(wsSource As Worksheet, wsTarget As Worksheet, sourceRow As Long, targetRow As Long, daysRemaining As Long, sapDict As Object)
    wsTarget.Cells(targetRow, 1).Value = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))
    wsTarget.Cells(targetRow, 1).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 1).VerticalAlignment = xlCenter
    
    ' Add SAP Code (Column 2) - lookup from sapDict by distributor name
    Dim distName As String
    distName = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))
    Dim sapCode As String
    sapCode = ""
    On Error Resume Next
    If sapDict.Exists(distName) Then
        sapCode = sapDict(distName)
    End If
    On Error GoTo 0
    
    wsTarget.Cells(targetRow, 2).Value = sapCode
    wsTarget.Cells(targetRow, 2).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 2).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 3).Value = Trim(CStr(wsSource.Cells(sourceRow, 17).Value))
    wsTarget.Cells(targetRow, 3).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 3).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 4).Value = Trim(CStr(wsSource.Cells(sourceRow, 7).Value))
    wsTarget.Cells(targetRow, 4).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 4).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 5).Value = Trim(CStr(wsSource.Cells(sourceRow, 10).Value))
    wsTarget.Cells(targetRow, 5).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 5).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 6).Value = Trim(CStr(wsSource.Cells(sourceRow, 13).Value))
    wsTarget.Cells(targetRow, 6).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 6).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 7).Value = wsSource.Cells(sourceRow, 14).Value
    wsTarget.Cells(targetRow, 7).NumberFormat = "yyyy-mm-dd"
     wsTarget.Cells(targetRow, 7).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 7).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 8).Value = daysRemaining
    wsTarget.Cells(targetRow, 8).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 8).VerticalAlignment = xlCenter
    
    If daysRemaining < 0 Then
        wsTarget.Cells(targetRow, 8).Interior.Color = RGB(255, 102, 102)
    ElseIf daysRemaining < 30 Then
        wsTarget.Cells(targetRow, 8).Interior.Color = RGB(255, 153, 153)
    ElseIf daysRemaining < 60 Then
        wsTarget.Cells(targetRow, 8).Interior.Color = RGB(255, 192, 128)
    ElseIf daysRemaining < 90 Then
        wsTarget.Cells(targetRow, 8).Interior.Color = RGB(255, 230, 153)
    End If
    
    ' NEW: Add ItemValue (source column 18)
    wsTarget.Cells(targetRow, 9).Value = wsSource.Cells(sourceRow, 18).Value
    wsTarget.Cells(targetRow, 9).NumberFormat = "#,##0"
    wsTarget.Cells(targetRow, 9).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 9).VerticalAlignment = xlCenter
End Sub

Private Sub CreateGeneralSummarySheet(ws As Worksheet, expiredCount As Long, month1Count As Long, month2Count As Long, month3Count As Long, wsExpired As Worksheet, ws1Month As Worksheet, ws2Month As Worksheet, ws3Month As Worksheet)
    Dim totalItems As Long
    totalItems = expiredCount + month1Count + month2Count + month3Count
    
    Dim qtyExpired As Long, qty1Month As Long, qty2Month As Long, qty3Month As Long
    Dim valExpired As Double, val1Month As Double, val2Month As Double, val3Month As Double
    Dim lastRowCalc As Long
    
    qtyExpired = 0: qty1Month = 0: qty2Month = 0: qty3Month = 0
    valExpired = 0: val1Month = 0: val2Month = 0: val3Month = 0
    
    On Error Resume Next
    If expiredCount > 0 Then
        lastRowCalc = wsExpired.Cells(wsExpired.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 1 Then
            qtyExpired = Application.WorksheetFunction.Sum(wsExpired.Range("F2:F" & lastRowCalc))
            valExpired = Application.WorksheetFunction.Sum(wsExpired.Range("I2:I" & lastRowCalc))
        End If
    End If
    
    If month1Count > 0 Then
        lastRowCalc = ws1Month.Cells(ws1Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 1 Then
            qty1Month = Application.WorksheetFunction.Sum(ws1Month.Range("F2:F" & lastRowCalc))
            val1Month = Application.WorksheetFunction.Sum(ws1Month.Range("I2:I" & lastRowCalc))
        End If
    End If
    
    If month2Count > 0 Then
        lastRowCalc = ws2Month.Cells(ws2Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 1 Then
            qty2Month = Application.WorksheetFunction.Sum(ws2Month.Range("F2:F" & lastRowCalc))
            val2Month = Application.WorksheetFunction.Sum(ws2Month.Range("I2:I" & lastRowCalc))
        End If
    End If
    
    If month3Count > 0 Then
        lastRowCalc = ws3Month.Cells(ws3Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 1 Then
            qty3Month = Application.WorksheetFunction.Sum(ws3Month.Range("F2:F" & lastRowCalc))
            val3Month = Application.WorksheetFunction.Sum(ws3Month.Range("I2:I" & lastRowCalc))
        End If
    End If
    On Error GoTo 0
     
    Dim totalQty As Long
    Dim totalVal As Double
    totalQty = qtyExpired + qty1Month + qty2Month + qty3Month
    totalVal = valExpired + val1Month + val2Month + val3Month
    
    ws.Cells.Font.Name = "Arial"
    
    With ws
        .Range("A1:E1").Merge
        .Range("A1").Value = ChrW(&H62A) & ChrW(&H642) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H6A9) & ChrW(&H627) & ChrW(&H645) & ChrW(&H644) & " FIFO"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 255, 255)
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        .Range("A1:E1").Borders.LineStyle = xlContinuous
        .Range("A1:E1").Borders.Weight = xlMedium
        .Range("A1:E1").Borders.Color = RGB(0, 0, 0)
        
        ' Headers: الفوع, عدد المنتجات, الکمیة, القیمة, الحالة
        .Range("A2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H641) & ChrW(&H648) & ChrW(&H639)
        .Range("B2").Value = ChrW(&H639) & ChrW(&H62F) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
        .Range("C2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629)
        .Range("D2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H6CC) & ChrW(&H645) & ChrW(&H629)
        .Range("E2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62D) & ChrW(&H627) & ChrW(&H644) & ChrW(&H629)
        
        .Range("A2:E2").Font.Bold = True
        .Range("A2:E2").Font.Size = 11
        .Range("A2:E2").Font.Color = RGB(255, 255, 255)
        .Range("A2:E2").Interior.Color = RGB(68, 114, 196)
        .Range("A2:E2").HorizontalAlignment = xlCenter
        .Range("A2:E2").VerticalAlignment = xlCenter
        .Range("A2:E2").WrapText = False
        .Range("A2:E2").ShrinkToFit = False
        .Range("A2:E2").Borders.LineStyle = xlContinuous
        .Range("A2:E2").Borders.Weight = xlMedium
        .Range("A2:E2").Borders.Color = RGB(0, 0, 0)
        
        ' Row 3: Expired
        .Cells(3, 1).Value = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(3, 2).Value = expiredCount
        .Cells(3, 3).Value = qtyExpired
        .Cells(3, 4).Value = valExpired
        .Cells(3, 5).Value = ChrW(&H62E) & ChrW(&H637) & ChrW(&H631)
        .Range("A3:E3").Interior.Color = RGB(255, 102, 102)
        .Range("A3:E3").WrapText = False
        .Range("A3:E3").ShrinkToFit = False
        .Range("A3:E3").Borders.LineStyle = xlContinuous
        .Range("A3:E3").Borders.Weight = xlThin
        .Range("A3:E3").Borders.Color = RGB(0, 0, 0)
        
        ' Row 4: 1 Month
        .Cells(4, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & " " & ChrW(&H648) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H62F)
        .Cells(4, 2).Value = month1Count
        .Cells(4, 3).Value = qty1Month
        .Cells(4, 4).Value = val1Month
        .Cells(4, 5).Value = ChrW(&H639) & ChrW(&H627) & ChrW(&H62C) & ChrW(&H644)
        .Range("A4:E4").Interior.Color = RGB(255, 153, 153)
        .Range("A4:E4").WrapText = False
        .Range("A4:E4").ShrinkToFit = False
        .Range("A4:E4").Borders.LineStyle = xlContinuous
        .Range("A4:E4").Borders.Weight = xlThin
        .Range("A4:E4").Borders.Color = RGB(0, 0, 0)
        
        ' Row 5: 2 Months
        .Cells(5, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        .Cells(5, 2).Value = month2Count
        .Cells(5, 3).Value = qty2Month
        .Cells(5, 4).Value = val2Month
        .Cells(5, 5).Value = ChrW(&H62A) & ChrW(&H62D) & ChrW(&H630) & ChrW(&H64A) & ChrW(&H631)
        .Range("A5:E5").Interior.Color = RGB(255, 192, 128)
        .Range("A5:E5").WrapText = False
        .Range("A5:E5").ShrinkToFit = False
        .Range("A5:E5").Borders.LineStyle = xlContinuous
        .Range("A5:E5").Borders.Weight = xlThin
        .Range("A5:E5").Borders.Color = RGB(0, 0, 0)
        
        ' Row 6: 3 Months
        .Cells(6, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(6, 2).Value = month3Count
        .Cells(6, 3).Value = qty3Month
        .Cells(6, 4).Value = val3Month
        .Cells(6, 5).Value = ChrW(&H62A) & ChrW(&H628) & ChrW(&H639) & ChrW(&H64A) & ChrW(&H62F)
        .Range("A6:E6").Interior.Color = RGB(255, 230, 153)
        .Range("A6:E6").WrapText = False
        .Range("A6:E6").ShrinkToFit = False
        .Range("A6:E6").Borders.LineStyle = xlContinuous
        .Range("A6:E6").Borders.Weight = xlThin
        .Range("A6:E6").Borders.Color = RGB(0, 0, 0)
        
        .Range("B3:D6").NumberFormat = "#,##0"
        .Range("A3:E6").HorizontalAlignment = xlCenter
        .Range("A3:E6").VerticalAlignment = xlCenter
        .Range("A3:E6").Font.Size = 11
        
        ' Row 7: Totals
        .Cells(7, 2).Value = totalItems
        .Cells(7, 3).Value = totalQty
        .Cells(7, 4).Value = totalVal
        
        .Range("B7:D7").NumberFormat = "#,##0"
        .Range("B7:D7").HorizontalAlignment = xlCenter
        .Range("B7:D7").VerticalAlignment = xlCenter
        .Range("B7:D7").Font.Size = 11
        .Range("B7:D7").Interior.Color = RGB(255, 255, 0)
        .Range("B7:D7").Font.Bold = True
        .Range("B7:D7").WrapText = False
        .Range("B7:D7").ShrinkToFit = False
        .Range("B7:D7").Borders.LineStyle = xlContinuous
        .Range("B7:D7").Borders.Weight = xlMedium
        .Range("B7:D7").Borders.Color = RGB(0, 0, 0)
        
        .Range("A7").Interior.ColorIndex = xlNone
        .Range("A7").Borders.LineStyle = xlNone
        .Range("E7").Interior.ColorIndex = xlNone
        .Range("E7").Borders.LineStyle = xlNone
        
        .Columns("A:E").AutoFit
        Dim colAdjust As Long
        For colAdjust = 1 To 5
            .Columns(colAdjust).ColumnWidth = .Columns(colAdjust).ColumnWidth + 3
        Next colAdjust
        
        .Range("A1:E7").WrapText = False
        .Range("A1:E7").ShrinkToFit = False
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

Private Sub CreateTop10StoresOnly(ws As Worksheet, wsE As Worksheet, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long, i As Long
    lr = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    
    If lr > 1 Then
        For i = 2 To lr
             Dim store As String, qty As Long
            store = Trim(CStr(wsE.Cells(i, 1).Value))
            On Error Resume Next
            qty = CLng(wsE.Cells(i, 6).Value)
            On Error GoTo 0
            
            If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                If Not dict.Exists(store) Then
                    dict.Add store, Array(wsE.Cells(i, 3).Value, 0, 0, 0)
                End If
                Dim arr As Variant
                Dim sValAmt As Double
                On Error Resume Next
                sValAmt = CDbl(wsE.Cells(i, 9).Value)
                On Error GoTo 0
                arr = dict(store)
                arr(1) = arr(1) + qty
                arr(2) = arr(2) + 1
                arr(3) = arr(3) + sValAmt
                dict(store) = arr
            End If
        Next i
    End If
    
    Dim sortedKeys() As String, sortedVals() As Long
    ReDim sortedKeys(0 To dict.Count - 1)
    ReDim sortedVals(0 To dict.Count - 1)
    
    Dim j As Long, k As Variant
    j = 0
    For Each k In dict.Keys
        sortedKeys(j) = CStr(k)
        sortedVals(j) = dict(k)(1)
        j = j + 1
    Next k
    
    Dim m As Long, n As Long, tempK As String, tempV As Long
    For m = 0 To UBound(sortedVals) - 1
        For n = m + 1 To UBound(sortedVals)
            If sortedVals(m) < sortedVals(n) Then
                tempV = sortedVals(m)
                sortedVals(m) = sortedVals(n)
                 sortedVals(n) = tempV
                
                tempK = sortedKeys(m)
                sortedKeys(m) = sortedKeys(n)
                sortedKeys(n) = tempK
            End If
        Next n
    Next m
    
    ws.Cells.Font.Name = "Arial"
    
    With ws
        .Range("A1:E1").Merge
       .Cells(1, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H639) & ChrW(&H646) & ChrW(&H62F) & ChrW(&H647) & ChrW(&H645) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(68, 114, 196)
         .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).VerticalAlignment = xlCenter
        .Cells(1, 1).WrapText = False
        .Cells(1, 1).ShrinkToFit = False
        .Range("A1:E1").Borders.LineStyle = xlContinuous
        .Range("A1:E1").Borders.Weight = xlMedium
        .Range("A1:E1").Borders.Color = RGB(0, 0, 0)
        
        Dim sr As Long
        sr = 2
        
        .Cells(sr, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H6CC) & ChrW(&H628)
        .Cells(sr, 2).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
        .Cells(sr, 3).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
        .Cells(sr, 4).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H643) & ChrW(&H645) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H629)
        .Cells(sr, 5).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
        .Range("A" & sr & ":E" & sr).Font.Bold = True
        .Range("A" & sr & ":E" & sr).Interior.Color = RGB(217, 217, 217)
        .Range("A" & sr & ":E" & sr).HorizontalAlignment = xlCenter
        .Range("A" & sr & ":E" & sr).VerticalAlignment = xlCenter
        .Range("A" & sr & ":E" & sr).WrapText = False
        .Range("A" & sr & ":E" & sr).ShrinkToFit = False
        .Range("A" & sr & ":E" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":E" & sr).Borders.Weight = xlMedium
        .Range("A" & sr & ":E" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 5).Interior.Color = RGB(146, 208, 80)
        .Cells(sr, 5).Font.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        Dim rank As Long
        For rank = 0 To 9
            If rank > UBound(sortedKeys) Then Exit For
            arr = dict(sortedKeys(rank))
            .Cells(sr, 1).Value = rank + 1
            .Cells(sr, 2).Value = sortedKeys(rank)
             .Cells(sr, 3).Value = arr(0)
            .Cells(sr, 4).Value = arr(1)
            .Cells(sr, 5).Value = arr(3)
            
            .Range("A" & sr & ":E" & sr).Interior.ColorIndex = xlNone
            .Range("A" & sr & ":E" & sr).HorizontalAlignment = xlCenter
            .Range("A" & sr & ":E" & sr).VerticalAlignment = xlCenter
            .Range("A" & sr & ":E" & sr).WrapText = False
            .Range("A" & sr & ":E" & sr).ShrinkToFit = False
            .Range("A" & sr & ":E" & sr).Borders.LineStyle = xlContinuous
            .Range("A" & sr & ":E" & sr).Borders.Weight = xlThin
            .Range("A" & sr & ":E" & sr).Borders.Color = RGB(0, 0, 0)
            .Cells(sr, 4).NumberFormat = "#,##0"
            .Cells(sr, 5).NumberFormat = "#,##0"
            sr = sr + 1
        Next rank
        
        ' Calculate total as value (no formula)
        Dim table1Total As Long
        Dim table1TotalVal As Double
        table1Total = 0
        table1TotalVal = 0
        Dim t1Rank As Long
        For t1Rank = 0 To 9
            If t1Rank > UBound(sortedKeys) Then Exit For
            table1Total = table1Total + dict(sortedKeys(t1Rank))(1)
            table1TotalVal = table1TotalVal + dict(sortedKeys(t1Rank))(3)
        Next t1Rank
        .Cells(sr, 4).Value = table1Total
        .Cells(sr, 5).Value = table1TotalVal
        .Range("A" & sr & ":C" & sr).Interior.ColorIndex = xlNone
        .Range("D" & sr & ":E" & sr).Font.Bold = True
        .Range("D" & sr & ":E" & sr).Interior.Color = RGB(146, 208, 80)
        .Range("D" & sr & ":E" & sr).HorizontalAlignment = xlCenter
        .Range("D" & sr & ":E" & sr).VerticalAlignment = xlCenter
        .Range("D" & sr & ":E" & sr).WrapText = False
        .Range("D" & sr & ":E" & sr).ShrinkToFit = False
        .Range("D" & sr & ":E" & sr).Borders.LineStyle = xlContinuous
        .Range("D" & sr & ":E" & sr).Borders.Weight = xlMedium
        .Range("D" & sr & ":E" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 4).NumberFormat = "#,##0"
        .Cells(sr, 5).NumberFormat = "#,##0"
        
        .Columns("A:E").AutoFit
        Dim colNum As Long
        For colNum = 1 To 5
            .Columns(colNum).ColumnWidth = .Columns(colNum).ColumnWidth + 3
        Next colNum
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        ' ========== ADD SECOND TABLE: TOP 10 BY TOTAL (ALL CATEGORIES) ==========
        ' Leave 1 empty row gap between tables
        If sr <= 14 Then
            ' Table 1 ends around row 13, leave row 14 empty, start at row 15
            sr = 15
        Else
            ' Leave 1 empty row after Table 1 total
            sr = sr + 2
        End If
        
        ' Collect data from ALL sheets (Expired, 1M, 2M, 3M)
        Dim dictAll As Object
        Set dictAll = CreateObject("Scripting.Dictionary")
        
        ' Process Expired sheet
        lr = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
        If lr > 1 Then
            For i = 2 To lr
                store = Trim(CStr(wsE.Cells(i, 1).Value))
                On Error Resume Next
                qty = CLng(wsE.Cells(i, 6).Value)
                On Error GoTo 0
                
                If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                    If Not dictAll.Exists(store) Then
                        dictAll.Add store, Array(wsE.Cells(i, 3).Value, 0, 0, 0, 0, 0)
                    End If
                    On Error Resume Next
                    sValAmt = CDbl(wsE.Cells(i, 9).Value)
                    On Error GoTo 0
                    arr = dictAll(store)
                    arr(1) = arr(1) + qty  ' Expired
                    arr(5) = arr(5) + sValAmt
                    dictAll(store) = arr
                End If
            Next i
        End If
        
        ' Process 1 Month sheet
        lr = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
        If lr > 1 Then
            For i = 2 To lr
                store = Trim(CStr(ws1.Cells(i, 1).Value))
                On Error Resume Next
                qty = CLng(ws1.Cells(i, 6).Value)
                On Error GoTo 0
                
                If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                    If Not dictAll.Exists(store) Then
                        dictAll.Add store, Array(ws1.Cells(i, 3).Value, 0, 0, 0, 0, 0)
                    End If
                    On Error Resume Next
                    sValAmt = CDbl(ws1.Cells(i, 9).Value)
                    On Error GoTo 0
                    arr = dictAll(store)
                    arr(2) = arr(2) + qty  ' 1 Month
                    arr(5) = arr(5) + sValAmt
                    dictAll(store) = arr
                End If
            Next i
        End If
        
        ' Process 2 Months sheet
        lr = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
        If lr > 1 Then
            For i = 2 To lr
                store = Trim(CStr(ws2.Cells(i, 1).Value))
                On Error Resume Next
                qty = CLng(ws2.Cells(i, 6).Value)
                On Error GoTo 0
                
                If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                    If Not dictAll.Exists(store) Then
                        dictAll.Add store, Array(ws2.Cells(i, 3).Value, 0, 0, 0, 0, 0)
                    End If
                    On Error Resume Next
                    sValAmt = CDbl(ws2.Cells(i, 9).Value)
                    On Error GoTo 0
                    arr = dictAll(store)
                    arr(3) = arr(3) + qty  ' 2 Months
                    arr(5) = arr(5) + sValAmt
                    dictAll(store) = arr
                End If
            Next i
        End If
        
        ' Process 3 Months sheet
        lr = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
        If lr > 1 Then
            For i = 2 To lr
                store = Trim(CStr(ws3.Cells(i, 1).Value))
                On Error Resume Next
                qty = CLng(ws3.Cells(i, 6).Value)
                On Error GoTo 0
                
                If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                    If Not dictAll.Exists(store) Then
                        dictAll.Add store, Array(ws3.Cells(i, 3).Value, 0, 0, 0, 0, 0)
                    End If
                    On Error Resume Next
                    sValAmt = CDbl(ws3.Cells(i, 9).Value)
                    On Error GoTo 0
                    arr = dictAll(store)
                    arr(4) = arr(4) + qty  ' 3 Months
                    arr(5) = arr(5) + sValAmt
                    dictAll(store) = arr
                End If
            Next i
        End If
        
        ' Sort by TOTAL (sum of all categories)
        ReDim sortedKeys(0 To dictAll.Count - 1)
        ReDim sortedVals(0 To dictAll.Count - 1)
        
        j = 0
        For Each k In dictAll.Keys
            sortedKeys(j) = CStr(k)
            arr = dictAll(k)
            sortedVals(j) = arr(1) + arr(2) + arr(3) + arr(4)  ' Total
            j = j + 1
        Next k
        
        For m = 0 To UBound(sortedVals) - 1
            For n = m + 1 To UBound(sortedVals)
                If sortedVals(m) < sortedVals(n) Then
                    tempV = sortedVals(m)
                    sortedVals(m) = sortedVals(n)
                    sortedVals(n) = tempV
                    
                    tempK = sortedKeys(m)
                    sortedKeys(m) = sortedKeys(n)
                    sortedKeys(n) = tempK
                End If
            Next n
        Next m
        
        ' Title for second table
        .Range("A" & sr & ":G" & sr).Merge
        .Cells(sr, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H62D) & ChrW(&H633) & ChrW(&H628) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        .Cells(sr, 1).Font.Bold = True
        .Cells(sr, 1).Font.Size = 14
        .Cells(sr, 1).Interior.Color = RGB(68, 114, 196)
        .Cells(sr, 1).Font.Color = RGB(255, 255, 255)
        .Cells(sr, 1).HorizontalAlignment = xlCenter
        .Cells(sr, 1).VerticalAlignment = xlCenter
        .Range("A" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":G" & sr).Borders.Weight = xlMedium
        sr = sr + 1
        
        ' Headers (same as RSM Performance)
        .Cells(sr, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
        .Cells(sr, 2).Value = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(sr, 3).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 4).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        .Cells(sr, 5).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " 3 " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 6).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        .Cells(sr, 7).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
        .Range("A" & sr & ":G" & sr).Font.Bold = True
        .Range("A" & sr & ":G" & sr).Interior.Color = RGB(217, 217, 217)
        .Range("A" & sr & ":G" & sr).HorizontalAlignment = xlCenter
        .Range("A" & sr & ":G" & sr).VerticalAlignment = xlCenter
        .Range("A" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":G" & sr).Borders.Weight = xlMedium
        
        ' Color code headers
        .Cells(sr, 2).Interior.Color = RGB(255, 102, 102)  ' Expired - Red
        .Cells(sr, 3).Interior.Color = RGB(255, 153, 153)  ' 1 Month - Light Red
        .Cells(sr, 4).Interior.Color = RGB(255, 192, 128)  ' 2 Months - Orange
        .Cells(sr, 5).Interior.Color = RGB(255, 230, 153)  ' 3 Months - Yellow
        .Cells(sr, 6).Interior.Color = RGB(255, 255, 0)    ' Total - Bright Yellow
        .Cells(sr, 6).Font.Color = RGB(0, 0, 0)
        .Cells(sr, 7).Interior.Color = RGB(146, 208, 80)   ' Value - Green
        .Cells(sr, 7).Font.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        ' Top 10 data rows
        For rank = 0 To 9
            If rank > UBound(sortedKeys) Then Exit For
            arr = dictAll(sortedKeys(rank))
            Dim distTotal As Long
            distTotal = arr(1) + arr(2) + arr(3) + arr(4)
            
            .Cells(sr, 1).Value = sortedKeys(rank)
            .Cells(sr, 2).Value = arr(1)  ' Expired
            .Cells(sr, 3).Value = arr(2)  ' 1 Month
            .Cells(sr, 4).Value = arr(3)  ' 2 Months
            .Cells(sr, 5).Value = arr(4)  ' 3 Months
            .Cells(sr, 6).Value = distTotal  ' Total
            .Cells(sr, 7).Value = arr(5)  ' Total Value
            
            .Range("A" & sr & ":G" & sr).Interior.ColorIndex = xlNone
            .Range("A" & sr & ":G" & sr).HorizontalAlignment = xlCenter
            .Range("A" & sr & ":G" & sr).VerticalAlignment = xlCenter
            .Range("A" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
            .Range("A" & sr & ":G" & sr).Borders.Weight = xlThin
            .Cells(sr, 2).NumberFormat = "#,##0"
            .Cells(sr, 3).NumberFormat = "#,##0"
            .Cells(sr, 4).NumberFormat = "#,##0"
            .Cells(sr, 5).NumberFormat = "#,##0"
            .Cells(sr, 6).NumberFormat = "#,##0"
            .Cells(sr, 7).NumberFormat = "#,##0"
            sr = sr + 1
        Next rank
        
        ' Total row (calculate as value to avoid any formula issues)
        Dim grandTotal2 As Long
        Dim grandTotal2Val As Double
        grandTotal2 = 0
        grandTotal2Val = 0
        Dim calcRow As Long
        For calcRow = sr - 10 To sr - 1
            grandTotal2 = grandTotal2 + .Cells(calcRow, 6).Value
            grandTotal2Val = grandTotal2Val + .Cells(calcRow, 7).Value
        Next calcRow
        .Cells(sr, 6).Value = grandTotal2
        .Cells(sr, 7).Value = grandTotal2Val
        .Range("F" & sr & ":G" & sr).Font.Bold = True
        .Range("F" & sr & ":G" & sr).Interior.Color = RGB(146, 208, 80)
        .Range("F" & sr & ":G" & sr).HorizontalAlignment = xlCenter
        .Range("F" & sr & ":G" & sr).VerticalAlignment = xlCenter
        .Range("F" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
        .Range("F" & sr & ":G" & sr).Borders.Weight = xlMedium
        .Cells(sr, 6).NumberFormat = "#,##0"
        .Cells(sr, 7).NumberFormat = "#,##0"
        
        .Columns("A:G").AutoFit
        For colNum = 1 To 7
            .Columns(colNum).ColumnWidth = .Columns(colNum).ColumnWidth + 2
        Next colNum
        
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

Private Sub CreateDashboard(ws As Worksheet, wsExp As Worksheet, ws1M As Worksheet, ws2M As Worksheet, ws3M As Worksheet)
    ws.Cells.Font.Name = "Arial"
    
    Call AddRSMPerformance(ws, wsExp, ws1M, ws2M, ws3M, 1)
    
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
    
    ws.Columns("A:G").AutoFit
    Dim colRSM As Long
    For colRSM = 1 To 7
        ws.Columns(colRSM).ColumnWidth = ws.Columns(colRSM).ColumnWidth + 3
    Next colRSM
    
    ws.Rows.AutoFit
    ws.Range("A1").Select
End Sub

Private Sub CreateProductAnalysis(ws As Worksheet, wsExp As Worksheet, ws1M As Worksheet, ws2M As Worksheet, ws3M As Worksheet)
    ws.Cells.Font.Name = "Arial"
    
    ws.Range("A1:I1").Merge
    ws.Cells(1, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A) & " " & ChrW(&H645) & ChrW(&H639) & ChrW(&H631) & ChrW(&H636) & ChrW(&H629) & " " & ChrW(&H644) & ChrW(&H644) & ChrW(&H627) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H627) & ChrW(&H621) & " (" & ChrW(&H623) & ChrW(&H648) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629) & ")"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Interior.Color = RGB(76, 175, 80)
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Cells(1, 1).VerticalAlignment = xlCenter
    ws.Cells(1, 1).WrapText = False
    ws.Cells(1, 1).ShrinkToFit = False
    ws.Range("A1:I1").Borders.LineStyle = xlContinuous
    ws.Range("A1:G1").Borders.Weight = xlMedium
    
    ws.Rows(2).RowHeight = 10
    
    Dim dictProducts As Object
    Set dictProducts = CreateObject("Scripting.Dictionary")
    
    Call CollectProductDataWithID(dictProducts, wsExp, 0)
    Call CollectProductDataWithID(dictProducts, ws1M, 1)
    Call CollectProductDataWithID(dictProducts, ws2M, 2)
    Call CollectProductDataWithID(dictProducts, ws3M, 3)
    
    Dim sortedProducts As Object
    Set sortedProducts = SortProductsByTotal(dictProducts)
    
    ws.Cells(3, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H6CC) & ChrW(&H628)
    ws.Cells(3, 2).Value = ChrW(&H6A9) & ChrW(&H648) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(3, 3).Value = ChrW(&H625) & ChrW(&H633) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(3, 4).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629)
    ws.Cells(3, 5).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws.Cells(3, 6).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    ws.Cells(3, 7).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws.Cells(3, 8).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
    ws.Cells(3, 9).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
    
    ws.Range("A3:I3").Font.Bold = True
    ws.Range("A3:I3").Font.Size = 11
    ws.Range("A3:I3").Font.Color = RGB(255, 255, 255)
    ws.Range("A3:I3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:I3").HorizontalAlignment = xlCenter
    ws.Range("A3:I3").VerticalAlignment = xlCenter
    ws.Range("A3:I3").WrapText = False
    ws.Range("A3:I3").ShrinkToFit = False
    ws.Range("A3:I3").Borders.LineStyle = xlContinuous
    ws.Range("A3:I3").Borders.Weight = xlMedium
    
    ws.Cells(3, 4).Interior.Color = RGB(255, 102, 102)
    ws.Cells(3, 5).Interior.Color = RGB(255, 153, 153)
    ws.Cells(3, 8).Interior.Color = RGB(255, 255, 0)
    ws.Cells(3, 8).Font.Color = RGB(0, 0, 0)
    ws.Cells(3, 9).Interior.Color = RGB(146, 208, 80)
    ws.Cells(3, 9).Font.Color = RGB(0, 0, 0)
    
    Dim rank As Long, rowNum As Long
    Dim productKey As Variant
    
    rank = 1
    rowNum = 4
    
    For Each productKey In sortedProducts.Keys
        If rank > 10 Then Exit For
        
        Dim prodData As Variant
        prodData = sortedProducts(productKey)
        
        Dim totalQty As Long
        totalQty = prodData(1) + prodData(2) + prodData(3) + prodData(4)
        
        ws.Cells(rowNum, 1).Value = rank
        ' FIXED: Column 2 gets the Key (which is now ItemID)
        ws.Cells(rowNum, 2).Value = productKey
        ' FIXED: Column 3 gets prodData(0) (which is now ProductName)
        ws.Cells(rowNum, 3).Value = prodData(0)
        
        ws.Cells(rowNum, 4).Value = prodData(1)
        ws.Cells(rowNum, 5).Value = prodData(2)
        ws.Cells(rowNum, 6).Value = prodData(3)
        ws.Cells(rowNum, 7).Value = prodData(4)
        ws.Cells(rowNum, 8).Value = totalQty
        ws.Cells(rowNum, 9).Value = prodData(5)
        
        ws.Range("A" & rowNum & ":I" & rowNum).Interior.ColorIndex = xlNone
        ws.Range("A" & rowNum & ":I" & rowNum).HorizontalAlignment = xlCenter
        ws.Range("A" & rowNum & ":I" & rowNum).VerticalAlignment = xlCenter
        ws.Range("A" & rowNum & ":I" & rowNum).Font.Size = 11
        ws.Range("A" & rowNum & ":I" & rowNum).WrapText = False
        ws.Range("A" & rowNum & ":I" & rowNum).ShrinkToFit = False
        ws.Range("A" & rowNum & ":I" & rowNum).Borders.LineStyle = xlContinuous
        ws.Range("A" & rowNum & ":I" & rowNum).Borders.Weight = xlThin
        
        ws.Cells(rowNum, 4).NumberFormat = "#,##0"
        ws.Cells(rowNum, 5).NumberFormat = "#,##0"
        ws.Cells(rowNum, 6).NumberFormat = "#,##0"
        ws.Cells(rowNum, 7).NumberFormat = "#,##0"
        ws.Cells(rowNum, 8).NumberFormat = "#,##0"
        ws.Cells(rowNum, 9).NumberFormat = "#,##0"
        
        rank = rank + 1
        rowNum = rowNum + 1
    Next productKey
    
    ' Add TOTAL row
    Dim totalRow As Long
    totalRow = rowNum
    Dim totalSum As Long
    Dim totalValSum As Double
    
    If totalRow > 4 Then
        totalSum = Application.WorksheetFunction.Sum(ws.Range("H4:H" & (totalRow - 1)))
        totalValSum = Application.WorksheetFunction.Sum(ws.Range("I4:I" & (totalRow - 1)))
    Else
        totalSum = 0
        totalValSum = 0
    End If
    
    ws.Cells(totalRow, 8).Value = totalSum
    ws.Cells(totalRow, 9).Value = totalValSum
    ws.Range("H" & totalRow & ":I" & totalRow).NumberFormat = "#,##0"
    ws.Range("H" & totalRow & ":I" & totalRow).Font.Bold = True
    ws.Range("H" & totalRow & ":I" & totalRow).Interior.Color = RGB(146, 208, 80)
    ws.Range("H" & totalRow & ":I" & totalRow).HorizontalAlignment = xlCenter
    ws.Range("H" & totalRow & ":I" & totalRow).VerticalAlignment = xlCenter
    ws.Range("H" & totalRow & ":I" & totalRow).Borders.LineStyle = xlContinuous
    ws.Range("H" & totalRow & ":I" & totalRow).Borders.Weight = xlMedium
    ws.Range("H" & totalRow & ":I" & totalRow).Borders.Color = RGB(0, 0, 0)
    
    ' Keep other cells in TOTAL row clean (no color, no border)
    ws.Range("A" & totalRow & ":G" & totalRow).Interior.ColorIndex = xlNone
    ws.Range("A" & totalRow & ":G" & totalRow).Borders.LineStyle = xlNone
    
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
    
    ws.Columns("A:I").AutoFit
    Dim colNum As Long
    For colNum = 1 To 9
        ws.Columns(colNum).ColumnWidth = ws.Columns(colNum).ColumnWidth + 3
    Next colNum
    
    ws.Rows.AutoFit
    
    ws.Range("A1:I" & totalRow).HorizontalAlignment = xlCenter
    ws.Range("A1:I" & totalRow).VerticalAlignment = xlCenter
    ws.Range("A1:I" & totalRow).WrapText = False
    ws.Range("A1:I" & totalRow).ShrinkToFit = False
    
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Sub CollectProductDataWithID(dict As Object, ws As Worksheet, categoryIndex As Long)
    Dim lastRow As Long
    ' Find last row using column 6 (quantity) - this INCLUDES the TOTAL row
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    ' CRITICAL: Loop to lastRow - 1 to exclude TOTAL row (which is always the last row)
    Dim i As Long
    For i = 2 To lastRow - 1
        Dim distributorName As String, productName As String, itemID As String, qty As Long
        
        distributorName = Trim(CStr(ws.Cells(i, 1).Value))
        productName = Trim(CStr(ws.Cells(i, 5).Value))
        itemID = Trim(CStr(ws.Cells(i, 4).Value))
        
        On Error Resume Next
        qty = CLng(ws.Cells(i, 6).Value)
        On Error GoTo 0
        
        ' Only process rows with valid data
        If Len(distributorName) > 0 And Len(productName) > 0 And Len(itemID) > 0 And qty > 0 Then
            
            ' FIXED: Group by ItemID instead of Name to prevent merging different items
            If Not dict.Exists(itemID) Then
                Dim newArr(0 To 5) As Variant
                newArr(0) = productName  ' Store Name in Index 0
                newArr(1) = 0: newArr(2) = 0: newArr(3) = 0: newArr(4) = 0: newArr(5) = 0
                dict.Add itemID, newArr
            End If
            
            Dim arr As Variant
            Dim pValAmt As Double
            On Error Resume Next
            pValAmt = CDbl(ws.Cells(i, 9).Value)
            On Error GoTo 0
            arr = dict(itemID)
            arr(categoryIndex + 1) = arr(categoryIndex + 1) + qty
            arr(5) = arr(5) + pValAmt
            dict(itemID) = arr
        End If
    Next i
End Sub

Private Function SortProductsByTotal(sourceDict As Object) As Object
    Dim sortedDict As Object
    Set sortedDict = CreateObject("Scripting.Dictionary")
    
    If sourceDict.Count = 0 Then
        Set SortProductsByTotal = sortedDict
        Exit Function
    End If
    
    Dim keys() As Variant, totals() As Long
    ReDim keys(0 To sourceDict.Count - 1)
    ReDim totals(0 To sourceDict.Count - 1)
    
    Dim i As Long, k As Variant
    i = 0
    For Each k In sourceDict.Keys
        Dim arr As Variant
        arr = sourceDict(k)
        keys(i) = k
        totals(i) = arr(1) + arr(2) + arr(3) + arr(4)
        i = i + 1
    Next k
    
    Dim j As Long, tempKey As Variant, tempTotal As Long
    For i = 0 To UBound(totals) - 1
        For j = i + 1 To UBound(totals)
            If totals(i) < totals(j) Then
                tempTotal = totals(i)
                totals(i) = totals(j)
                totals(j) = tempTotal
                 
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
    
    For i = 0 To UBound(keys)
        sortedDict.Add keys(i), sourceDict(keys(i))
    Next i
    
    Set SortProductsByTotal = sortedDict
End Function

Private Sub AddRSMPerformance(ws As Worksheet, wsE As Worksheet, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, sr As Long)
    ws.Cells.Font.Name = "Arial"
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long, i As Long
    
    lr = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            Dim rsm As String, qty As Long
            rsm = Trim(CStr(wsE.Cells(i, 3).Value))
            On Error Resume Next
            qty = CLng(wsE.Cells(i, 6).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(wsE.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                End If
                Dim arr As Variant
                Dim valAmt As Double
                On Error Resume Next
                valAmt = CDbl(wsE.Cells(i, 9).Value)
                On Error GoTo 0
                arr = dict(rsm)
                arr(0) = arr(0) + 1
                arr(1) = arr(1) + qty
                arr(8) = arr(8) + valAmt
                dict(rsm) = arr
            End If
         Next i
    End If
    
    lr = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws1.Cells(i, 3).Value))
            On Error Resume Next
            qty = CLng(ws1.Cells(i, 6).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(ws1.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                valAmt = CDbl(ws1.Cells(i, 9).Value)
                On Error GoTo 0
                arr = dict(rsm)
                arr(2) = arr(2) + 1
                arr(3) = arr(3) + qty
                arr(9) = arr(9) + valAmt
                dict(rsm) = arr
            End If
        Next i
    End If
    
    lr = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws2.Cells(i, 3).Value))
            On Error Resume Next
            qty = CLng(ws2.Cells(i, 6).Value)
            On Error GoTo 0
             
            If Len(rsm) > 0 And UCase(ws2.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                valAmt = CDbl(ws2.Cells(i, 9).Value)
                On Error GoTo 0
                arr = dict(rsm)
                arr(4) = arr(4) + 1
                arr(5) = arr(5) + qty
                arr(10) = arr(10) + valAmt
                dict(rsm) = arr
            End If
        Next i
    End If
    
    lr = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws3.Cells(i, 3).Value))
            On Error Resume Next
            qty = CLng(ws3.Cells(i, 6).Value)
            On Error GoTo 0
             
            If Len(rsm) > 0 And UCase(ws3.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                valAmt = CDbl(ws3.Cells(i, 9).Value)
                On Error GoTo 0
                arr = dict(rsm)
                arr(6) = arr(6) + 1
                arr(7) = arr(7) + qty
                arr(11) = arr(11) + valAmt
                dict(rsm) = arr
            End If
        Next i
    End If
    
    Dim sortedKeys() As String, sortedVals() As Long
    ReDim sortedKeys(0 To dict.Count - 1)
    ReDim sortedVals(0 To dict.Count - 1)
    
    Dim j As Long, k As Variant
    j = 0
    For Each k In dict.Keys
        sortedKeys(j) = CStr(k)
        arr = dict(k)
        sortedVals(j) = arr(1) + arr(3) + arr(5) + arr(7)
        j = j + 1
    Next k
    
    Dim m As Long, n As Long, tempK As String, tempV As Long
    For m = 0 To UBound(sortedVals) - 1
        For n = m + 1 To UBound(sortedVals)
            If sortedVals(m) < sortedVals(n) Then
                tempV = sortedVals(m)
                sortedVals(m) = sortedVals(n)
                sortedVals(n) = tempV
                
                tempK = sortedKeys(m)
                sortedKeys(m) = sortedKeys(n)
                sortedKeys(n) = tempK
            End If
        Next n
    Next m
    
    With ws
        .Range("A" & sr & ":G" & sr).Merge
        .Cells(sr, 1).Value = ChrW(&H62A) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H64A) & ChrW(&H645) & " " & ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H645) & ChrW(&H62F) & ChrW(&H631) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & " (RSMs)"
        .Cells(sr, 1).Font.Bold = True
        .Cells(sr, 1).Interior.Color = RGB(68, 114, 196)
        .Cells(sr, 1).Font.Color = RGB(255, 255, 255)
        .Cells(sr, 1).HorizontalAlignment = xlCenter
        .Cells(sr, 1).VerticalAlignment = xlCenter
        .Cells(sr, 1).WrapText = False
        .Cells(sr, 1).ShrinkToFit = False
        .Cells(sr, 1).Borders.LineStyle = xlContinuous
        .Cells(sr, 1).Borders.Weight = xlMedium
        .Cells(sr, 1).Borders.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        .Cells(sr, 1).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
        .Cells(sr, 2).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(sr, 3).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 4).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        .Cells(sr, 5).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 6).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        .Cells(sr, 7).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
        
        .Range("A" & sr & ":G" & sr).Font.Bold = True
        .Range("A" & sr & ":G" & sr).Interior.Color = RGB(217, 217, 217)
        .Range("A" & sr & ":G" & sr).HorizontalAlignment = xlCenter
        .Range("A" & sr & ":G" & sr).VerticalAlignment = xlCenter
        .Range("A" & sr & ":G" & sr).WrapText = False
        .Range("A" & sr & ":G" & sr).ShrinkToFit = False
        .Range("A" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":G" & sr).Borders.Weight = xlMedium
        .Range("A" & sr & ":G" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 2).Interior.Color = RGB(255, 102, 102)
        .Cells(sr, 3).Interior.Color = RGB(255, 153, 153)
        .Cells(sr, 4).Interior.Color = RGB(255, 192, 128)
        .Cells(sr, 5).Interior.Color = RGB(255, 230, 153)
        .Cells(sr, 6).Interior.Color = RGB(255, 255, 0)
        .Cells(sr, 6).Font.Color = RGB(0, 0, 0)
        .Cells(sr, 7).Interior.Color = RGB(146, 208, 80)
        .Cells(sr, 7).Font.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        For j = 0 To UBound(sortedKeys)
            arr = dict(sortedKeys(j))
            Dim rsmTotal As Long
            Dim rsmTotalVal As Double
            rsmTotal = arr(1) + arr(3) + arr(5) + arr(7)
            rsmTotalVal = arr(8) + arr(9) + arr(10) + arr(11)
            
            .Cells(sr, 1).Value = sortedKeys(j)
            .Cells(sr, 2).Value = arr(1)
            .Cells(sr, 3).Value = arr(3)
            .Cells(sr, 4).Value = arr(5)
            .Cells(sr, 5).Value = arr(7)
            .Cells(sr, 6).Value = rsmTotal
            .Cells(sr, 7).Value = rsmTotalVal
            
             If arr(1) > 10000 Then
                .Range("A" & sr & ":E" & sr).Interior.Color = RGB(255, 102, 102)
            ElseIf arr(1) > 5000 Then
                .Range("A" & sr & ":E" & sr).Interior.Color = RGB(255, 192, 128)
            End If
        
            .Range("A" & sr & ":G" & sr).HorizontalAlignment = xlCenter
            .Range("A" & sr & ":G" & sr).VerticalAlignment = xlCenter
            .Range("A" & sr & ":G" & sr).WrapText = False
            .Range("A" & sr & ":G" & sr).ShrinkToFit = False
            .Range("A" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
            .Range("A" & sr & ":G" & sr).Borders.Weight = xlThin
            .Range("A" & sr & ":G" & sr).Borders.Color = RGB(0, 0, 0)
            .Cells(sr, 2).NumberFormat = "#,##0"
            .Cells(sr, 3).NumberFormat = "#,##0"
            .Cells(sr, 4).NumberFormat = "#,##0"
            .Cells(sr, 5).NumberFormat = "#,##0"
            .Cells(sr, 6).NumberFormat = "#,##0"
            .Cells(sr, 7).NumberFormat = "#,##0"
            sr = sr + 1
        Next j
        
        Dim grandTotal As Long
        Dim grandTotalVal As Double
        grandTotal = 0
        grandTotalVal = 0
        For Each k In dict.Keys
            arr = dict(k)
            grandTotal = grandTotal + arr(1) + arr(3) + arr(5) + arr(7)
            grandTotalVal = grandTotalVal + arr(8) + arr(9) + arr(10) + arr(11)
        Next k
        
        .Cells(sr, 6).Value = grandTotal
        .Cells(sr, 7).Value = grandTotalVal
        .Range("F" & sr & ":G" & sr).Font.Bold = True
        .Range("F" & sr & ":G" & sr).Interior.Color = RGB(146, 208, 80)
        .Range("F" & sr & ":G" & sr).HorizontalAlignment = xlCenter
        .Range("F" & sr & ":G" & sr).VerticalAlignment = xlCenter
        .Range("F" & sr & ":G" & sr).WrapText = False
        .Range("F" & sr & ":G" & sr).ShrinkToFit = False
        .Range("F" & sr & ":G" & sr).Borders.LineStyle = xlContinuous
        .Range("F" & sr & ":G" & sr).Borders.Weight = xlMedium
        .Range("F" & sr & ":G" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 6).NumberFormat = "#,##0"
        .Cells(sr, 7).NumberFormat = "#,##0"
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With
End Sub


' ========================================
' VBA #2: RSM SPLITTER
' ========================================

Public Sub SplitFIFOByRSM_Complete()
    ' Standalone version - uses ThisWorkbook
    Call SplitFIFOByRSM_Complete_Internal(ThisWorkbook)
End Sub

Public Sub SplitFIFOByRSM_Complete_Internal(wbSource As Workbook)
    ' Internal version - accepts workbook parameter

    Dim wsTotal As Worksheet
    Dim dictRSM As Object
    Dim lastRow As Long, i As Long
    Dim rsmVal As String
    Dim TodayStamp As String
    Dim MasterFolder As String
    Dim rsmKey As Variant, idx As Long

    'Set wbSource = ThisWorkbook ' (now passed as parameter)
    Set wsTotal = wbSource.Sheets(wbSource.Sheets.Count)

    If wbSource.Path = "" Then
        MsgBox "Save the workbook first.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    TodayStamp = Format(Date, "dd-mmm-yyyy")
    MasterFolder = wbSource.Path & "\FIFO_Per_RSM_" & TodayStamp
    
    ' Archive existing RSM folder if it exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(MasterFolder) Then
        Dim archivePath As String
        Dim versionNum As Long
        Dim archiveFolder As String
        Dim archiveBase As String
        
        archiveBase = wbSource.Path & "\M. FIFO Archive"
        archivePath = archiveBase & "\Archive [Outcome of 2nd VBA]"
        
        ' Create archive folders if they don't exist
        If Not fso.FolderExists(archiveBase) Then fso.CreateFolder archiveBase
        If Not fso.FolderExists(archivePath) Then fso.CreateFolder archivePath
        
        ' Find next version number
        versionNum = 1
        Do While fso.FolderExists(archivePath & "\FIFO_Per_RSM_" & TodayStamp & " v" & versionNum)
            versionNum = versionNum + 1
        Loop
        
        archiveFolder = archivePath & "\FIFO_Per_RSM_" & TodayStamp & " v" & versionNum
        
        ' Try to move folder first
        On Error Resume Next
        fso.MoveFolder MasterFolder, archiveFolder
        
        ' If move fails, try copy then delete
        If Err.Number <> 0 Then
            Err.Clear
            ' Copy folder to archive
            fso.CopyFolder MasterFolder, archiveFolder, True
            If Err.Number = 0 Then
                ' Delete original after successful copy
                fso.DeleteFolder MasterFolder, True
            End If
            
            If Err.Number <> 0 Then
                Err.Clear
                MsgBox "Warning: Could not archive RSM folder. Close any open files and try again.", vbExclamation
                Set fso = Nothing
                GoTo CleanupRSM
            End If
        End If
        On Error GoTo 0
    End If
    
    SafeMkDir MasterFolder
    Set fso = Nothing

    Set dictRSM = CreateObject("Scripting.Dictionary")

    lastRow = wsTotal.Cells(wsTotal.Rows.Count, 3).End(xlUp).Row
    For i = 2 To lastRow
        On Error Resume Next
        rsmVal = Trim(wsTotal.Cells(i, 3).Value & "")
        On Error GoTo 0
        If rsmVal <> "" And rsmVal <> "TOTAL" And rsmVal <> "المجموع" Then
            If Not dictRSM.Exists(rsmVal) Then dictRSM.Add rsmVal, True
        End If
    Next i

    idx = 0
    For Each rsmKey In dictRSM.Keys
        idx = idx + 1
        CreateRSMWorkbook wbSource, MasterFolder & "\RSM_" & Format(idx, "0000"), CStr(rsmKey), TodayStamp
    Next rsmKey

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done! Created " & idx & " RSM files in:" & vbCrLf & MasterFolder, vbInformation
    Exit Sub
    
CleanupRSM:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub CreateRSMWorkbook(wbSource As Workbook, rsmFolder As String, rsmName As String, stamp As String)

    Dim wbNew As Workbook
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim sheetIdx As Long

    SafeMkDir rsmFolder
    Set wbNew = Workbooks.Add(xlWBATWorksheet)

    For sheetIdx = 1 To wbSource.Sheets.Count
        Set wsSource = wbSource.Sheets(sheetIdx)
        wsSource.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
    Next sheetIdx

    ' STEP 1: Delete sheet 3 (أداء مدراء RSM)
    Application.DisplayAlerts = False
    wbNew.Sheets(3).Delete
    Application.DisplayAlerts = True
    
    ' STEP 2: Filter data sheets (5-8 in new workbook after deletion: منتهية, بقی شهر, بقی شهرین, بقی ٣ أشهر)
    Dim dataSheetIdx As Long
    For dataSheetIdx = 5 To 8
        DeleteNonMatchingRows wbNew.Sheets(dataSheetIdx), rsmName
        SortByDistributor wbNew.Sheets(dataSheetIdx)
        RecalculateTotal wbNew.Sheets(dataSheetIdx)
    Next dataSheetIdx
    
    ' STEP 3: Generate summary sheets FROM the filtered data sheets (5-8)
    GenerateSummaryFromFiltered wbNew.Sheets(2), wbNew.Sheets(5), wbNew.Sheets(6), wbNew.Sheets(7), wbNew.Sheets(8)
    
    GenerateTop10ProductsFromFiltered wbNew.Sheets(3), wbNew.Sheets(5), wbNew.Sheets(6), wbNew.Sheets(7), wbNew.Sheets(8)
    
    ' FIXED: Rename sheet 3 to "أعلى 10 منتجات"
    Application.DisplayAlerts = False
    On Error Resume Next
    wbNew.Sheets(3).Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' STEP 4: Regenerate top distributors sheet FROM filtered expired data
    GenerateTopDistributorsFromFiltered wbNew.Sheets(4), wbNew.Sheets(5), wbNew.Sheets(6), wbNew.Sheets(7), wbNew.Sheets(8), rsmName
    
    ' Rename the sheet from أعلى 10 وکلاء to أعلى وکلاء
    Application.DisplayAlerts = False
    On Error Resume Next
    wbNew.Sheets(4).Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621)
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' STEP 5: Filter کامل sheet (sheet 9) by RSM and SORT IT
    FilterSheetByRSM wbNew.Sheets(9), rsmName
    SortByDistributor wbNew.Sheets(9)
    RecalculateTotal wbNew.Sheets(9)

    ' Delete the blank sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wbNew.Sheets(1).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Save the file
    On Error Resume Next
    wbNew.SaveAs rsmFolder & "\FIFO_" & CleanFileName(rsmName) & "_" & stamp & ".xlsx", xlOpenXMLWorkbook, Local:=True
    wbNew.Close False
    On Error GoTo 0
End Sub

Private Sub GenerateSummaryFromFiltered(wsOut As Worksheet, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet)
    Dim expired As Long, expired_qty As Long
    Dim m1 As Long, m1_qty As Long
    Dim m2 As Long, m2_qty As Long
    Dim m3 As Long, m3_qty As Long
    Dim lastRowCalc As Long

    wsOut.Cells.Clear
    
    ' Title
    Application.DisplayAlerts = False
    wsOut.Range("A1:E1").Merge
    Application.DisplayAlerts = True
    wsOut.Cells(1, 1).Value = ChrW(&H62A) & ChrW(&H642) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H6A9) & ChrW(&H627) & ChrW(&H645) & ChrW(&H644) & " FIFO"
    wsOut.Cells(1, 1).Font.Bold = True
    wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(1, 1).Interior.Color = RGB(68, 114, 196)
    wsOut.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsOut.Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' Headers
    wsOut.Cells(2, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H641) & ChrW(&H648) & ChrW(&H639)
    wsOut.Cells(2, 2).Value = ChrW(&H639) & ChrW(&H62F) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
    wsOut.Cells(2, 3).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629)
    wsOut.Cells(2, 4).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62D) & ChrW(&H627) & ChrW(&H644) & ChrW(&H629)
    
    With wsOut.Range("A2:D2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' FIXED BUG #1: Use column 1 for lastRow, then SUM to lastRow (not lastRow-1)
    On Error Resume Next
    lastRowCalc = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    If lastRowCalc > 1 Then
        expired = lastRowCalc - 1
        expired_qty = Application.WorksheetFunction.Sum(ws1.Range("F2:F" & lastRowCalc))
    End If
    
    lastRowCalc = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    If lastRowCalc > 1 Then
        m1 = lastRowCalc - 1
        m1_qty = Application.WorksheetFunction.Sum(ws2.Range("F2:F" & lastRowCalc))
    End If
    
    lastRowCalc = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
    If lastRowCalc > 1 Then
        m2 = lastRowCalc - 1
        m2_qty = Application.WorksheetFunction.Sum(ws3.Range("F2:F" & lastRowCalc))
    End If
    
    lastRowCalc = ws4.Cells(ws4.Rows.Count, 1).End(xlUp).Row
    If lastRowCalc > 1 Then
        m3 = lastRowCalc - 1
        m3_qty = Application.WorksheetFunction.Sum(ws4.Range("F2:F" & lastRowCalc))
    End If
    On Error GoTo 0
    
    ' Data rows
    wsOut.Cells(3, 1).Value = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    wsOut.Cells(3, 2).Value = expired
    wsOut.Cells(3, 3).Value = expired_qty
    wsOut.Cells(3, 4).Value = ChrW(&H62E) & ChrW(&H637) & ChrW(&H631)
    wsOut.Range("A3:D3").Interior.Color = RGB(255, 102, 102)
    
    wsOut.Cells(4, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & " " & ChrW(&H648) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H62F)
    wsOut.Cells(4, 2).Value = m1
    wsOut.Cells(4, 3).Value = m1_qty
    wsOut.Cells(4, 4).Value = ChrW(&H639) & ChrW(&H627) & ChrW(&H62C) & ChrW(&H644)
    wsOut.Range("A4:D4").Interior.Color = RGB(255, 153, 153)
    
    wsOut.Cells(5, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    wsOut.Cells(5, 2).Value = m2
    wsOut.Cells(5, 3).Value = m2_qty
    wsOut.Cells(5, 4).Value = ChrW(&H62A) & ChrW(&H62D) & ChrW(&H630) & ChrW(&H64A) & ChrW(&H631)
    wsOut.Range("A5:D5").Interior.Color = RGB(255, 192, 128)
    
    wsOut.Cells(6, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    wsOut.Cells(6, 2).Value = m3
    wsOut.Cells(6, 3).Value = m3_qty
    wsOut.Cells(6, 4).Value = ChrW(&H62A) & ChrW(&H628) & ChrW(&H639) & ChrW(&H64A) & ChrW(&H62F)
    wsOut.Range("A6:D6").Interior.Color = RGB(255, 230, 153)
    
    ' Total row
    wsOut.Cells(7, 2).Value = expired + m1 + m2 + m3
    wsOut.Cells(7, 3).Value = expired_qty + m1_qty + m2_qty + m3_qty
    wsOut.Range("B7:C7").Interior.Color = RGB(255, 255, 0)
    wsOut.Range("B7:C7").Font.Bold = True
    
    ' FORMATTING - Grid borders on all data cells (thin borders)
    With wsOut.Range("A1:D7")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
    End With
    
    ' Thicker borders for title row
    With wsOut.Range("A1:D1")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(0, 0, 0)
    End With
    
    ' Thicker borders for header row
    With wsOut.Range("A2:D2")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(0, 0, 0)
    End With
    
    ' Remove borders from empty cells in TOTAL row (A7 and D7)
    wsOut.Range("A7").Borders.LineStyle = xlNone
    wsOut.Range("A7").Interior.ColorIndex = xlNone
    wsOut.Range("D7").Borders.LineStyle = xlNone
    wsOut.Range("D7").Interior.ColorIndex = xlNone
    
    ' Thicker borders for TOTAL cells that have data
    With wsOut.Range("B7:C7")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(0, 0, 0)
    End With
    
    ' Number format for quantities
    wsOut.Range("B3:C6").NumberFormat = "#,##0"
    wsOut.Range("B7:C7").NumberFormat = "#,##0"
    
    ' Row heights for professional look
    wsOut.Rows(1).RowHeight = 30
    wsOut.Rows("2:7").RowHeight = 25
    
    ' Column widths
    wsOut.Columns("A:D").AutoFit
    Dim colAdjust As Long
    For colAdjust = 1 To 4
        wsOut.Columns(colAdjust).ColumnWidth = wsOut.Columns(colAdjust).ColumnWidth + 3
    Next colAdjust
End Sub

Private Sub GenerateTop10ProductsFromFiltered(wsOut As Worksheet, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet)
    Dim dictProducts As Object
    Dim i As Long, lastRow As Long
    Dim prodCode As String, prodName As String, qty As Long
    Dim k As Variant, arr As Variant
    Dim keys() As String, totals() As Long
    Dim j As Long, m As Long, n As Long, tempK As String, tempT As Long
    Dim rank As Long, row As Long
    
    Set dictProducts = CreateObject("Scripting.Dictionary")
    
    wsOut.Cells.Clear
    
    ' Title
    Application.DisplayAlerts = False
    wsOut.Range("A1:I1").Merge
    Application.DisplayAlerts = True
    wsOut.Cells(1, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A) & " " & ChrW(&H645) & ChrW(&H639) & ChrW(&H631) & ChrW(&H636) & ChrW(&H629) & " " & ChrW(&H644) & ChrW(&H644) & ChrW(&H627) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H627) & ChrW(&H621) & " (" & ChrW(&H623) & ChrW(&H648) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629) & ")"
    wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(1, 1).Font.Bold = True
    wsOut.Cells(1, 1).Interior.Color = RGB(76, 175, 80)
    wsOut.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsOut.Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' Headers in row 3
    wsOut.Cells(3, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H6CC) & ChrW(&H628)
    wsOut.Cells(3, 2).Value = ChrW(&H6A9) & ChrW(&H648) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    wsOut.Cells(3, 3).Value = ChrW(&H625) & ChrW(&H633) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    wsOut.Cells(3, 4).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629)
    wsOut.Cells(3, 5).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    wsOut.Cells(3, 6).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    wsOut.Cells(3, 7).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    wsOut.Cells(3, 8).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
    wsOut.Cells(3, 9).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
    
    With wsOut.Range("A3:I3")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
    
    wsOut.Cells(3, 4).Interior.Color = RGB(255, 102, 102)
    wsOut.Cells(3, 5).Interior.Color = RGB(255, 153, 153)
    wsOut.Cells(3, 8).Interior.Color = RGB(255, 255, 0)
    wsOut.Cells(3, 8).Font.Color = RGB(0, 0, 0)
    wsOut.Cells(3, 9).Interior.Color = RGB(146, 208, 80)
    wsOut.Cells(3, 9).Font.Color = RGB(0, 0, 0)
    
    ' Collect from filtered sheets
    Call CollectFromFilteredSheet(dictProducts, ws1, 1)
    Call CollectFromFilteredSheet(dictProducts, ws2, 2)
    Call CollectFromFilteredSheet(dictProducts, ws3, 3)
    Call CollectFromFilteredSheet(dictProducts, ws4, 4)
    
    ' Sort by total
    If dictProducts.Count > 0 Then
        ReDim keys(0 To dictProducts.Count - 1)
        ReDim totals(0 To dictProducts.Count - 1)
        j = 0
        For Each k In dictProducts.Keys
            keys(j) = CStr(k)
            arr = dictProducts(k)
            totals(j) = arr(1) + arr(2) + arr(3) + arr(4)
            j = j + 1
        Next k
        
        For m = 0 To UBound(totals) - 1
            For n = m + 1 To UBound(totals)
                If totals(m) < totals(n) Then
                    tempT = totals(m)
                    totals(m) = totals(n)
                    totals(n) = tempT
                    tempK = keys(m)
                    keys(m) = keys(n)
                    keys(n) = tempK
                End If
            Next n
        Next m
        
        ' Output top 10
        rank = 1
        row = 4
        For j = 0 To UBound(keys)
            If rank > 10 Then Exit For
            arr = dictProducts(keys(j))
            
            wsOut.Cells(row, 1).Value = rank
            wsOut.Cells(row, 2).Value = keys(j)
            wsOut.Cells(row, 3).Value = arr(0)
            wsOut.Cells(row, 4).Value = arr(1)
            wsOut.Cells(row, 5).Value = arr(2)
            wsOut.Cells(row, 6).Value = arr(3)
            wsOut.Cells(row, 7).Value = arr(4)
            wsOut.Cells(row, 8).Value = arr(1) + arr(2) + arr(3) + arr(4)
            wsOut.Cells(row, 9).Value = arr(5)
            
            With wsOut.Range("A" & row & ":I" & row)
                .HorizontalAlignment = xlCenter
            End With
            
            For m = 4 To 9
                wsOut.Cells(row, m).NumberFormat = "#,##0"
            Next m
            
            row = row + 1
            rank = rank + 1
        Next j
    End If
    
    wsOut.Columns("A:I").AutoFit
End Sub

Private Sub CollectFromFilteredSheet(dict As Object, ws As Worksheet, catIndex As Long)
    Dim lastRow As Long, i As Long
    Dim prodCode As String, prodName As String, distName As String, qty As Long
    Dim arr As Variant
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' FIXED BUG #3: Added distributor check to exclude TOTAL rows
    For i = 2 To lastRow
        distName = Trim(ws.Cells(i, 1).Value & "")
        prodCode = Trim(ws.Cells(i, 4).Value & "")
        prodName = Trim(ws.Cells(i, 5).Value & "")
        On Error Resume Next
        qty = CLng(ws.Cells(i, 6).Value)
        On Error GoTo 0
        
        If distName <> "" And prodCode <> "" And prodName <> "" And qty > 0 Then
            If Not dict.Exists(prodCode) Then
                dict.Add prodCode, Array(prodName, 0, 0, 0, 0, 0)
            End If
            Dim fValAmt As Double
            On Error Resume Next
            fValAmt = CDbl(ws.Cells(i, 9).Value)
            On Error GoTo 0
            arr = dict(prodCode)
            arr(catIndex) = arr(catIndex) + qty
            arr(5) = arr(5) + fValAmt
            dict(prodCode) = arr
        End If
    Next i
End Sub

Private Sub GenerateTopDistributorsFromFiltered(wsOut As Worksheet, wsExpired As Worksheet, ws1M As Worksheet, ws2M As Worksheet, ws3M As Worksheet, rsmName As String)
    ' Regenerate top distributors sheet from filtered expired data
    Dim dictDist As Object
    Set dictDist = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, i As Long
    Dim distName As String, qty As Long
    
    ' Collect distributors from expired sheet (excluding TOTAL row)
    lastRow = wsExpired.Cells(wsExpired.Rows.Count, 1).End(xlUp).Row
    
    Dim dValAmt As Double
    For i = 2 To lastRow
        distName = Trim(wsExpired.Cells(i, 1).Value & "")
        On Error Resume Next
        qty = CLng(wsExpired.Cells(i, 6).Value)
        dValAmt = CDbl(wsExpired.Cells(i, 9).Value)
        On Error GoTo 0
        
        ' Skip empty, TOTAL, or zero quantity
        If distName <> "" And UCase(distName) <> "TOTAL" And qty > 0 Then
            If Not dictDist.Exists(distName) Then
                dictDist.Add distName, Array(qty, dValAmt)
            Else
                Dim tmpArr As Variant
                tmpArr = dictDist(distName)
                tmpArr(0) = tmpArr(0) + qty
                tmpArr(1) = tmpArr(1) + dValAmt
                dictDist(distName) = tmpArr
            End If
        End If
    Next i
    
    ' Clear and rebuild the sheet
    wsOut.Cells.Clear
    
    ' Title
    Application.DisplayAlerts = False
    wsOut.Range("A1:F1").Merge
    Application.DisplayAlerts = True
    wsOut.Cells(1, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H639) & ChrW(&H646) & ChrW(&H62F) & ChrW(&H647) & ChrW(&H645) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    wsOut.Cells(1, 1).Font.Bold = True
    wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(1, 1).Interior.Color = RGB(68, 114, 196)
    wsOut.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsOut.Cells(1, 1).HorizontalAlignment = xlCenter
    wsOut.Cells(1, 1).VerticalAlignment = xlCenter
    wsOut.Range("A1:F1").Borders.LineStyle = xlContinuous
    wsOut.Range("A1:F1").Borders.Weight = xlMedium
    
    ' Headers
    wsOut.Cells(2, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H6CC) & ChrW(&H628)
    wsOut.Cells(2, 2).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
    wsOut.Cells(2, 3).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
    wsOut.Cells(2, 4).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H643) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H629)
    wsOut.Cells(2, 5).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
    
    With wsOut.Range("A2:E2")
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' Sort distributors by quantity (descending)
    If dictDist.Count > 0 Then
        Dim keys() As String, vals() As Long
        ReDim keys(0 To dictDist.Count - 1)
        ReDim vals(0 To dictDist.Count - 1)
        
        Dim j As Long, k As Variant
        j = 0
        For Each k In dictDist.Keys
            keys(j) = CStr(k)
            vals(j) = dictDist(k)(0)
            j = j + 1
        Next k
        
        ' Bubble sort descending
        Dim m As Long, n As Long, tempK As String, tempV As Long
        For m = 0 To UBound(vals) - 1
            For n = m + 1 To UBound(vals)
                If vals(m) < vals(n) Then
                    tempV = vals(m)
                    vals(m) = vals(n)
                    vals(n) = tempV
                    tempK = keys(m)
                    keys(m) = keys(n)
                    keys(n) = tempK
                End If
            Next n
        Next m
        
        ' Output data rows
        Dim rowNum As Long, totalQty As Long
        rowNum = 3
        totalQty = 0
        
        Dim totalVal As Double
        totalVal = 0
        For j = 0 To UBound(keys)
            Dim distArr As Variant
            distArr = dictDist(keys(j))
            wsOut.Cells(rowNum, 1).Value = j + 1  ' Rank
            wsOut.Cells(rowNum, 2).Value = keys(j)  ' Distributor name
            wsOut.Cells(rowNum, 3).Value = rsmName  ' RSM name
            wsOut.Cells(rowNum, 4).Value = distArr(0)  ' Quantity
            wsOut.Cells(rowNum, 5).Value = distArr(1)  ' Value
            
            wsOut.Range("A" & rowNum & ":E" & rowNum).HorizontalAlignment = xlCenter
            wsOut.Range("A" & rowNum & ":E" & rowNum).VerticalAlignment = xlCenter
            wsOut.Cells(rowNum, 4).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 5).NumberFormat = "#,##0"
            
            totalQty = totalQty + distArr(0)
            totalVal = totalVal + distArr(1)
            rowNum = rowNum + 1
        Next j
        
        ' Add TOTAL row
        wsOut.Cells(rowNum, 4).Value = totalQty
        wsOut.Cells(rowNum, 5).Value = totalVal
        wsOut.Range("D" & rowNum & ":E" & rowNum).NumberFormat = "#,##0"
        wsOut.Range("D" & rowNum & ":E" & rowNum).Font.Bold = True
        wsOut.Range("D" & rowNum & ":E" & rowNum).Interior.Color = RGB(146, 208, 80)
        wsOut.Range("D" & rowNum & ":E" & rowNum).HorizontalAlignment = xlCenter
        wsOut.Range("D" & rowNum & ":E" & rowNum).VerticalAlignment = xlCenter
        wsOut.Range("D" & rowNum & ":E" & rowNum).Borders.LineStyle = xlContinuous
        wsOut.Range("D" & rowNum & ":E" & rowNum).Borders.Weight = xlMedium
    End If
    
    ' Auto-fit columns for Table 1
    wsOut.Columns("A:E").AutoFit
    
    ' ========== ADD TABLE 2: TOP 10 BY TOTAL (ALL CATEGORIES) ==========
    
    ' Leave 1 empty row gap between tables
    If rowNum <= 3 Then
        ' No data in Table 1 (headers at row 2), leave row 3 empty, start at row 4
        rowNum = 4
    Else
        ' Has data in Table 1 (total at rowNum), leave next row empty, start at rowNum + 2
        rowNum = rowNum + 2
    End If
    
    ' Collect data from ALL 4 category sheets
    Dim dictAll As Object
    Set dictAll = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long
    
    ' Process Expired sheet
    lr = wsExpired.Cells(wsExpired.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            distName = Trim(wsExpired.Cells(i, 1).Value & "")
            On Error Resume Next
            qty = CLng(wsExpired.Cells(i, 6).Value)
            On Error GoTo 0
            
            If distName <> "" And UCase(distName) <> "TOTAL" And qty > 0 Then
                If Not dictAll.Exists(distName) Then
                    dictAll.Add distName, Array(0, 0, 0, 0, 0)
                End If
                Dim arrAll As Variant
                On Error Resume Next
                dValAmt = CDbl(wsExpired.Cells(i, 9).Value)
                On Error GoTo 0
                arrAll = dictAll(distName)
                arrAll(0) = arrAll(0) + qty  ' Expired
                arrAll(4) = arrAll(4) + dValAmt  ' Value
                dictAll(distName) = arrAll
            End If
        Next i
    End If
    
    ' Process 1 Month sheet
    lr = ws1M.Cells(ws1M.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            distName = Trim(ws1M.Cells(i, 1).Value & "")
            On Error Resume Next
            qty = CLng(ws1M.Cells(i, 6).Value)
            On Error GoTo 0
            
            If distName <> "" And UCase(distName) <> "TOTAL" And qty > 0 Then
                If Not dictAll.Exists(distName) Then
                    dictAll.Add distName, Array(0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                dValAmt = CDbl(ws1M.Cells(i, 9).Value)
                On Error GoTo 0
                arrAll = dictAll(distName)
                arrAll(1) = arrAll(1) + qty  ' 1 Month
                arrAll(4) = arrAll(4) + dValAmt  ' Value
                dictAll(distName) = arrAll
            End If
        Next i
    End If
    
    ' Process 2 Months sheet
    lr = ws2M.Cells(ws2M.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            distName = Trim(ws2M.Cells(i, 1).Value & "")
            On Error Resume Next
            qty = CLng(ws2M.Cells(i, 6).Value)
            On Error GoTo 0
            
            If distName <> "" And UCase(distName) <> "TOTAL" And qty > 0 Then
                If Not dictAll.Exists(distName) Then
                    dictAll.Add distName, Array(0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                dValAmt = CDbl(ws2M.Cells(i, 9).Value)
                On Error GoTo 0
                arrAll = dictAll(distName)
                arrAll(2) = arrAll(2) + qty  ' 2 Months
                arrAll(4) = arrAll(4) + dValAmt  ' Value
                dictAll(distName) = arrAll
            End If
        Next i
    End If
    
    ' Process 3 Months sheet
    lr = ws3M.Cells(ws3M.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            distName = Trim(ws3M.Cells(i, 1).Value & "")
            On Error Resume Next
            qty = CLng(ws3M.Cells(i, 6).Value)
            On Error GoTo 0
            
            If distName <> "" And UCase(distName) <> "TOTAL" And qty > 0 Then
                If Not dictAll.Exists(distName) Then
                    dictAll.Add distName, Array(0, 0, 0, 0, 0)
                End If
                On Error Resume Next
                dValAmt = CDbl(ws3M.Cells(i, 9).Value)
                On Error GoTo 0
                arrAll = dictAll(distName)
                arrAll(3) = arrAll(3) + qty  ' 3 Months
                arrAll(4) = arrAll(4) + dValAmt  ' Value
                dictAll(distName) = arrAll
            End If
        Next i
    End If
    
    ' Sort by TOTAL
    If dictAll.Count > 0 Then
        ReDim keys(0 To dictAll.Count - 1)
        ReDim vals(0 To dictAll.Count - 1)
        
        j = 0
        For Each k In dictAll.Keys
            keys(j) = CStr(k)
            arrAll = dictAll(k)
            vals(j) = arrAll(0) + arrAll(1) + arrAll(2) + arrAll(3)
            j = j + 1
        Next k
        
        ' Bubble sort descending
        For m = 0 To UBound(vals) - 1
            For n = m + 1 To UBound(vals)
                If vals(m) < vals(n) Then
                    tempV = vals(m)
                    vals(m) = vals(n)
                    vals(n) = tempV
                    tempK = keys(m)
                    keys(m) = keys(n)
                    keys(n) = tempK
                End If
            Next n
        Next m
        
        ' Title for Table 2
        Application.DisplayAlerts = False
        wsOut.Range("A" & rowNum & ":G" & rowNum).Merge
        Application.DisplayAlerts = True
        wsOut.Cells(rowNum, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H62D) & ChrW(&H633) & ChrW(&H628) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        wsOut.Cells(rowNum, 1).Font.Bold = True
        wsOut.Cells(rowNum, 1).Font.Size = 14
        wsOut.Cells(rowNum, 1).Interior.Color = RGB(68, 114, 196)
        wsOut.Cells(rowNum, 1).Font.Color = RGB(255, 255, 255)
        wsOut.Cells(rowNum, 1).HorizontalAlignment = xlCenter
        wsOut.Cells(rowNum, 1).VerticalAlignment = xlCenter
        wsOut.Range("A" & rowNum & ":G" & rowNum).Borders.LineStyle = xlContinuous
        wsOut.Range("A" & rowNum & ":G" & rowNum).Borders.Weight = xlMedium
        rowNum = rowNum + 1
        
        ' Headers
        wsOut.Cells(rowNum, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
        wsOut.Cells(rowNum, 2).Value = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        wsOut.Cells(rowNum, 3).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        wsOut.Cells(rowNum, 4).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        wsOut.Cells(rowNum, 5).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " 3 " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        wsOut.Cells(rowNum, 6).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        wsOut.Cells(rowNum, 7).Value = ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H629)
        
        With wsOut.Range("A" & rowNum & ":G" & rowNum)
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        
        ' Color code headers
        wsOut.Cells(rowNum, 2).Interior.Color = RGB(255, 102, 102)
        wsOut.Cells(rowNum, 3).Interior.Color = RGB(255, 153, 153)
        wsOut.Cells(rowNum, 4).Interior.Color = RGB(255, 192, 128)
        wsOut.Cells(rowNum, 5).Interior.Color = RGB(255, 230, 153)
        wsOut.Cells(rowNum, 6).Interior.Color = RGB(255, 255, 0)
        wsOut.Cells(rowNum, 6).Font.Color = RGB(0, 0, 0)
        wsOut.Cells(rowNum, 7).Interior.Color = RGB(146, 208, 80)
        wsOut.Cells(rowNum, 7).Font.Color = RGB(0, 0, 0)
        rowNum = rowNum + 1
        
        ' Data rows (top 10)
        Dim distTotal As Long, grandTotalAll As Long, grandTotalAllVal As Double
        grandTotalAll = 0
        grandTotalAllVal = 0
        
        For j = 0 To UBound(keys)
            If j >= 10 Then Exit For  ' Only top 10
            
            arrAll = dictAll(keys(j))
            distTotal = arrAll(0) + arrAll(1) + arrAll(2) + arrAll(3)
            
            wsOut.Cells(rowNum, 1).Value = keys(j)  ' Distributor
            wsOut.Cells(rowNum, 2).Value = arrAll(0)  ' Expired
            wsOut.Cells(rowNum, 3).Value = arrAll(1)  ' 1 Month
            wsOut.Cells(rowNum, 4).Value = arrAll(2)  ' 2 Months
            wsOut.Cells(rowNum, 5).Value = arrAll(3)  ' 3 Months
            wsOut.Cells(rowNum, 6).Value = distTotal  ' Total
            wsOut.Cells(rowNum, 7).Value = arrAll(4)  ' Total Value
            
            wsOut.Range("A" & rowNum & ":G" & rowNum).HorizontalAlignment = xlCenter
            wsOut.Range("A" & rowNum & ":G" & rowNum).VerticalAlignment = xlCenter
            wsOut.Range("A" & rowNum & ":G" & rowNum).Borders.LineStyle = xlContinuous
            wsOut.Range("A" & rowNum & ":G" & rowNum).Borders.Weight = xlThin
            
            wsOut.Cells(rowNum, 2).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 3).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 4).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 5).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 6).NumberFormat = "#,##0"
            wsOut.Cells(rowNum, 7).NumberFormat = "#,##0"
            
            grandTotalAll = grandTotalAll + distTotal
            grandTotalAllVal = grandTotalAllVal + arrAll(4)
            rowNum = rowNum + 1
        Next j
        
        ' Total row (calculated value, not formula)
        wsOut.Cells(rowNum, 6).Value = grandTotalAll
        wsOut.Cells(rowNum, 7).Value = grandTotalAllVal
        wsOut.Range("F" & rowNum & ":G" & rowNum).NumberFormat = "#,##0"
        wsOut.Range("F" & rowNum & ":G" & rowNum).Font.Bold = True
        wsOut.Range("F" & rowNum & ":G" & rowNum).Interior.Color = RGB(146, 208, 80)
        wsOut.Range("F" & rowNum & ":G" & rowNum).HorizontalAlignment = xlCenter
        wsOut.Range("F" & rowNum & ":G" & rowNum).VerticalAlignment = xlCenter
        wsOut.Range("F" & rowNum & ":G" & rowNum).Borders.LineStyle = xlContinuous
        wsOut.Range("F" & rowNum & ":G" & rowNum).Borders.Weight = xlMedium
    End If
    
    ' Auto-fit all columns
    wsOut.Columns("A:G").AutoFit
End Sub

Private Sub DeleteNonMatchingRows(ws As Worksheet, rsmName As String)
    Dim lastRow As Long, i As Long, j As Long
    Dim rsmVal As String

    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    For i = lastRow To 2 Step -1
        On Error Resume Next
        rsmVal = Trim(ws.Cells(i, 3).Value & "")
        On Error GoTo 0

        If rsmVal <> "" And rsmVal <> rsmName And rsmVal <> "TOTAL" And rsmVal <> "المجموع" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        For j = 1 To 9
            If ws.Cells(i, j).Value = "#REF!" Then
                ws.Cells(i, j).Value = ""
            End If
        Next j
    Next i
End Sub

Private Sub SortByDistributor(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 3 Then Exit Sub
    
    On Error Resume Next
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add ws.Columns(1), xlSortOnValues, xlAscending
    ws.Sort.SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    ws.Sort.Header = xlYes
    ws.Sort.Apply
    On Error GoTo 0
End Sub

Private Sub RecalculateTotal(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim totalQty As Long
    Dim j As Long
    Dim col1Val As String, col3Val As String
    
    ' FIXED: Use column 6 (quantity) to find last row, not column 1
    ' TOTAL rows have empty column 1, so using column 1 misses them!
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    ' FIXED BUG #2: Delete TOTAL rows properly (check for empty column 1)
    For i = lastRow To 2 Step -1
        col1Val = Trim(ws.Cells(i, 1).Value & "")
        col3Val = Trim(ws.Cells(i, 3).Value & "")
        
        If col1Val = "" Or col1Val = "TOTAL" Or col3Val = "TOTAL" Or col3Val = "المجموع" Then
            ws.Rows(i).Delete
            lastRow = lastRow - 1
        End If
    Next i
    
    For i = 2 To lastRow
        For j = 1 To 9
            If ws.Cells(i, j).Value = "#REF!" Then
                ws.Cells(i, j).Value = ""
            End If
        Next j
    Next i
    
    ' Recalculate lastRow using column 1 (for data rows only)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' If no data rows (only header), set lastRow to 1
    If lastRow < 2 Then lastRow = 1
    
    totalQty = 0
    Dim totalValue As Double
    totalValue = 0
    For i = 2 To lastRow
        On Error Resume Next
        totalQty = totalQty + CLng(ws.Cells(i, 6).Value)
        totalValue = totalValue + CDbl(ws.Cells(i, 9).Value)
        On Error GoTo 0
    Next i
    
    lastRow = lastRow + 1
    ws.Range("A" & lastRow & ":I" & lastRow).ClearContents
    ws.Cells(lastRow, 6).Value = totalQty
    ws.Cells(lastRow, 6).NumberFormat = "#,##0"
    ws.Cells(lastRow, 6).Font.Bold = True
    ws.Cells(lastRow, 6).Interior.Color = RGB(146, 208, 80)
    ws.Cells(lastRow, 6).HorizontalAlignment = xlCenter
    ws.Cells(lastRow, 6).VerticalAlignment = xlCenter
    ws.Cells(lastRow, 6).Borders.LineStyle = xlContinuous
    ws.Cells(lastRow, 6).Borders.Weight = xlMedium
    ws.Cells(lastRow, 6).Borders.Color = RGB(0, 0, 0)
    
    ' NEW: Add ItemValue total with same formatting as Qty
    ws.Cells(lastRow, 9).Value = totalValue
    ws.Cells(lastRow, 9).NumberFormat = "#,##0"
    ws.Cells(lastRow, 9).Font.Bold = True
    ws.Cells(lastRow, 9).Interior.Color = RGB(146, 208, 80)
    ws.Cells(lastRow, 9).HorizontalAlignment = xlCenter
    ws.Cells(lastRow, 9).VerticalAlignment = xlCenter
    ws.Cells(lastRow, 9).Borders.LineStyle = xlContinuous
    ws.Cells(lastRow, 9).Borders.Weight = xlMedium
    ws.Cells(lastRow, 9).Borders.Color = RGB(0, 0, 0)
End Sub

Private Sub FilterSheetByRSM(ws As Worksheet, rsmName As String)
    Dim lastRow As Long, i As Long, j As Long
    Dim rsmVal As String, col1Val As String
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    ' FIXED BUG #4: Delete non-matching rows properly
    For i = lastRow To 2 Step -1
        On Error Resume Next
        rsmVal = Trim(ws.Cells(i, 3).Value & "")
        col1Val = Trim(ws.Cells(i, 1).Value & "")
        On Error GoTo 0
        
        If rsmVal <> rsmName And rsmVal <> "" And col1Val <> "" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        For j = 1 To 9
            If ws.Cells(i, j).Value = "#REF!" Then
                ws.Cells(i, j).Value = ""
            End If
        Next j
    Next i
End Sub

Private Sub SafeMkDir(ByVal p As String)
    On Error Resume Next
    If Dir(p, vbDirectory) = "" Then MkDir p
    On Error GoTo 0
End Sub

Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant
    s = Trim(s)
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, bad, "-")
    Next
    CleanFileName = s
End Function