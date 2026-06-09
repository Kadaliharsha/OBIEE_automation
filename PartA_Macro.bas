Option Explicit

' =========================================================================
' 1. MAIN ENTRY POINTS (Macros tied to dashboard buttons)
' =========================================================================

Sub BrowseOBIEE()
    SelectFileAndWriteToCell "Select OBIEE Export File", "B5"
End Sub

Sub BrowseServiceMax()
    SelectFileAndWriteToCell "Select ServiceMax Export File", "B7"
End Sub

Sub RunDataCleansing()
    Dim pathOBIEE As String, pathSMax As String
    Dim wsDash As Worksheet, wsOutput As Worksheet, wsConfig As Worksheet
    Dim dictSMax As Object, dictCountries As Object
    
    On Error GoTo ErrorHandler
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    pathOBIEE = wsDash.Range("B5").Value
    pathSMax = wsDash.Range("B7").Value
    
    If pathOBIEE = "" Or pathSMax = "" Then
        MsgBox "Please select both files first.", vbExclamation
        Exit Sub
    End If
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Orchestration ---
    
    ' Step 1: Load Config rules (Valid Countries)
    Set dictCountries = GetValidCountries(wsConfig)
    If dictCountries.Count = 0 Then
        MsgBox "No valid countries found in Config sheet Column A.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Step 2: Load ServiceMax existing keys
    Set dictSMax = BuildServiceMaxDictionary(pathSMax)
    If dictSMax Is Nothing Then GoTo Cleanup ' Halt if error occurred during load
    
    ' Step 3: Clear Output sheet
    wsOutput.Cells.Clear
    
    ' Step 4: Process OBIEE data and output new IPs
    ProcessOBIEEAndOutput pathOBIEE, dictSMax, wsOutput, dictCountries
    
    MsgBox "Process Complete! Data exported to Output sheet.", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred in Main Routine: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub


' =========================================================================
' 2. UI / CONFIG / FILE HELPERS
' =========================================================================

Private Sub SelectFileAndWriteToCell(dialogTitle As String, targetCell As String)
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv"
        If .Show = -1 Then
            ThisWorkbook.Sheets("Dashboard").Range(targetCell).Value = .SelectedItems(1)
        End If
    End With
End Sub

Private Function GetValidCountries(wsConfig As Worksheet) As Object
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim countryName As String, countryABB As String
    Dim colCountry As Integer, colABB As Integer
    Set dict = CreateObject("Scripting.Dictionary")
    
    colCountry = FindColumnHeader(wsConfig, "Country", 1)
    colABB = FindColumnHeader(wsConfig, "ABB", 1)
    
    ' Default to column E and D if headers aren't found
    If colCountry = 0 Then colCountry = 5
    If colABB = 0 Then colABB = 4
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        For i = 2 To lastRow
            countryName = UCase(Trim(CStr(wsConfig.Cells(i, colCountry).Value)))
            countryABB = UCase(Trim(CStr(wsConfig.Cells(i, colABB).Value)))
            
            If countryName <> "" And Not dict.Exists(countryName) Then dict.Add countryName, True
            If countryABB <> "" And Not dict.Exists(countryABB) Then dict.Add countryABB, True
        Next i
    End If
    Set GetValidCountries = dict
End Function


' =========================================================================
' 3. SERVICEMAX DATA EXTRACTION
' =========================================================================

Private Function BuildServiceMaxDictionary(filePath As String) As Object
    Dim wbSMax As Workbook
    Dim wsSMax As Worksheet
    Dim dict As Object
    Dim colSO As Integer, colESET As Integer, colSerial As Integer
    Dim i As Long, lastRow As Long
    Dim key As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error GoTo SMaxError
    Set wbSMax = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSMax = wbSMax.Sheets(1)
    
    ' Assuming ServiceMax has headers on row 1
    colSO = FindColumnHeader(wsSMax, "Sales order number", 1)
    colESET = FindColumnHeader(wsSMax, "ESET", 1)
    colSerial = FindColumnHeader(wsSMax, "Serial Number", 1)
    
    If colSO = 0 Or colESET = 0 Or colSerial = 0 Then
        MsgBox "Required columns missing in ServiceMax file.", vbCritical
        wbSMax.Close False
        Set BuildServiceMaxDictionary = Nothing
        Exit Function
    End If
    
    lastRow = wsSMax.Cells(wsSMax.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        key = CStr(wsSMax.Cells(i, colSO).Value) & "_" & _
              CStr(wsSMax.Cells(i, colESET).Value) & "_" & _
              CStr(wsSMax.Cells(i, colSerial).Value)
              
        If Not dict.Exists(key) Then dict.Add key, True
    Next i
    
    wbSMax.Close False
    Set BuildServiceMaxDictionary = dict
    Exit Function

SMaxError:
    MsgBox "Error reading ServiceMax file: " & Err.Description, vbCritical
    If Not wbSMax Is Nothing Then wbSMax.Close False
    Set BuildServiceMaxDictionary = Nothing
End Function


' =========================================================================
' 4. OBIEE PROCESSING & COMPARISON
' =========================================================================

Private Sub ProcessOBIEEAndOutput(filePath As String, dictSMax As Object, wsOutput As Worksheet, dictCountries As Object)
    Dim wbOBIEE As Workbook
    Dim wsOBIEE As Worksheet
    Dim colCountry As Integer, colSO As Integer, colESET As Integer, colSerial As Integer, colModality As Integer
    Dim i As Long, lastRow As Long, lastCol As Long, outRow As Long
    Dim key As String
    Dim headerRow As Long, dataStartRow As Long
    
    ' The requirement is to skip the first 2 rows of the OBIEE file.
    ' This means headers are on row 3, and data starts on row 4.
    headerRow = 3
    dataStartRow = 4
    
    On Error GoTo OBIEEError
    Set wbOBIEE = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsOBIEE = wbOBIEE.Sheets(1)
    
    colCountry = FindColumnHeader(wsOBIEE, "Country", headerRow)
    colSO = FindColumnHeader(wsOBIEE, "Sales order number", headerRow)
    colESET = FindColumnHeader(wsOBIEE, "ESET", headerRow)
    colSerial = FindColumnHeader(wsOBIEE, "Serial Number", headerRow)
    colModality = FindColumnHeader(wsOBIEE, "Modality", headerRow)
    
    If colCountry = 0 Or colSO = 0 Then
        MsgBox "Required columns (Country, Sales order number) missing in OBIEE file on row 3.", vbCritical
        wbOBIEE.Close False
        Exit Sub
    End If
    
    lastRow = wsOBIEE.Cells(wsOBIEE.Rows.Count, "A").End(xlUp).Row
    lastCol = wsOBIEE.Cells(headerRow, wsOBIEE.Columns.Count).End(xlToLeft).Column
    
    ' Copy headers to Output sheet
    wsOBIEE.Range(wsOBIEE.Cells(headerRow, 1), wsOBIEE.Cells(headerRow, lastCol)).Copy Destination:=wsOutput.Range("A1")
    outRow = 2
    
    For i = dataStartRow To lastRow
        ' Extract row data and validate against business rules
        If IsValidRecord(wsOBIEE, i, colCountry, colModality, dictCountries) Then
            
            ' Generate unique matching key
            key = CStr(wsOBIEE.Cells(i, colSO).Value) & "_" & _
                  CStr(wsOBIEE.Cells(i, colESET).Value) & "_" & _
                  CStr(wsOBIEE.Cells(i, colSerial).Value)
                  
            ' If NOT in ServiceMax dict, it's a new IP. Copy row to output.
            If Not dictSMax.Exists(key) Then
                wsOBIEE.Range(wsOBIEE.Cells(i, 1), wsOBIEE.Cells(i, lastCol)).Copy Destination:=wsOutput.Cells(outRow, 1)
                outRow = outRow + 1
            End If
            
        End If
    Next i
    
    wbOBIEE.Close False
    wsOutput.Columns.AutoFit
    Exit Sub

OBIEEError:
    MsgBox "Error processing OBIEE data: " & Err.Description, vbCritical
    If Not wbOBIEE Is Nothing Then wbOBIEE.Close False
End Sub


' =========================================================================
' 5. BUSINESS RULES ENGINE
' =========================================================================

Private Function IsValidRecord(ws As Worksheet, rowIdx As Long, colCountry As Integer, colMod As Integer, dictCountries As Object) As Boolean
    Dim countryVal As String, modVal As String
    
    countryVal = UCase(Trim(ws.Cells(rowIdx, colCountry).Value))
    If colMod > 0 Then modVal = UCase(Trim(ws.Cells(rowIdx, colMod).Value)) Else modVal = ""
    
    ' Default to False
    IsValidRecord = False
    
    ' Rule 1: Country must exist in Config Sheet
    If Not dictCountries.Exists(countryVal) Then Exit Function
    
    ' Rule 2: Modality must start with 5
    If Left(modVal, 1) <> "5" Then Exit Function
    
    ' If it passes all rules, it is valid
    IsValidRecord = True
End Function


' =========================================================================
' 6. UTILITY FUNCTIONS
' =========================================================================

Private Function FindColumnHeader(ws As Worksheet, headerName As String, headerRow As Long) As Integer
    Dim headerCell As Range
    Set headerCell = ws.Rows(headerRow).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not headerCell Is Nothing Then
        FindColumnHeader = headerCell.Column
    Else
        FindColumnHeader = 0
    End If
End Function


' =========================================================================
' 7. DASHBOARD GENERATOR (Run this once to set up the workbook!)
' =========================================================================

Sub SetupProfessionalDashboard()
    Dim wsDash As Worksheet, wsTemp As Worksheet
    Dim btn As Shape
    
    ' Setup Config and Output sheets
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets("Config")
    If wsTemp Is Nothing Then Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): wsTemp.Name = "Config"
    
    ' Pre-fill config with the exact table format provided
    wsTemp.Cells.Clear
    
    Dim headers As Variant
    headers = Array("Name", "Service Area", "GE Team Lead", "ABB", "Country")
    wsTemp.Range("A1:E1").Value = headers
    wsTemp.Range("A1:E1").Font.Bold = True
    wsTemp.Range("A1:E1").Interior.Color = RGB(220, 230, 241)
    
    Dim rawData As Variant
    rawData = Array( _
        Array("Australia", "AKA", "Rashmi", "AU", "Australia"), _
        Array("Brunei", "AKA", "Rashmi", "BN", "Brunei"), _
        Array("Cambodia", "AKA", "Rashmi", "KH", "Cambodia"), _
        Array("Fiji", "AKA", "Rashmi", "FJ", "Fiji"), _
        Array("French Polynesia", "AKA", "Rashmi", "PF", "French Polynesia"), _
        Array("Indonesia", "AKA", "Rashmi", "ID", "Indonesia"), _
        Array("Lao People's Democratic Republic", "AKA", "Rashmi", "LA", "Lao People's Democratic Republic"), _
        Array("Malaysia", "AKA", "Rashmi", "MY", "Malaysia"), _
        Array("Myanmar", "AKA", "Rashmi", "MM", "Myanmar"), _
        Array("New Zealand", "AKA", "Rashmi", "NZ", "New Zealand"), _
        Array("Norfolk Island", "AKA", "Rashmi", "NF", "Norfolk Island"), _
        Array("Papua New Guinea", "AKA", "Rashmi", "PG", "Papua New Guinea"), _
        Array("Philippines", "AKA", "Rashmi", "PH", "Philippines"), _
        Array("Korea, Republic Of", "AKA", "Rashmi", "KR", "Korea, Republic Of"), _
        Array("Samoa", "AKA", "Rashmi", "WS", "Samoa"), _
        Array("Singapore", "AKA", "Rashmi", "SG", "Singapore"), _
        Array("Thailand", "AKA", "Rashmi", "TH", "Thailand"), _
        Array("Tonga", "AKA", "Rashmi", "TO", "Tonga"), _
        Array("Tuvalu", "AKA", "Rashmi", "TV", "Tuvalu"), _
        Array("Vanuatu", "AKA", "Rashmi", "VU", "Vanuatu"), _
        Array("Vietnam", "AKA", "Rashmi", "VN", "Vietnam"), _
        Array("Korea", "AKA", "Rashmi", "KR", "Korea"), _
        Array("Viet Nam", "AKA", "Rashmi", "VN", "Viet Nam") _
    )
    
    Dim i As Integer, j As Integer
    For i = LBound(rawData) To UBound(rawData)
        For j = LBound(rawData(i)) To UBound(rawData(i))
            wsTemp.Cells(i + 2, j + 1).Value = rawData(i)(j)
        Next j
    Next i
    
    wsTemp.Columns("A:E").AutoFit
    
    Set wsTemp = ThisWorkbook.Sheets("Output")
    If wsTemp Is Nothing Then Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): wsTemp.Name = "Output"
    On Error GoTo 0
    
    ' Create or set Dashboard sheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsDash.Name = "Dashboard"
    End If
    
    ' Clean slate
    wsDash.Cells.Clear
    For Each btn In wsDash.Shapes
        btn.Delete
    Next btn
    
    ' Setup UI Design
    ActiveWindow.DisplayGridlines = False
    
    ' Simple Header
    With wsDash.Range("B2")
        .Value = "OBIEE Automation Dashboard"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Name = "Segoe UI"
    End With
    
    ' File Path Inputs Labels
    wsDash.Range("B4").Value = "OBIEE Export File:"
    wsDash.Range("B6").Value = "ServiceMax Export File:"
    wsDash.Range("B4:B6").Font.Name = "Segoe UI"
    wsDash.Range("B4:B6").Font.Size = 11
    wsDash.Range("B4:B6").Font.Bold = True
    
    ' Input boxes (B5 and B7)
    With wsDash.Range("B5:F5, B7:F7")
        .Merge
        .Interior.Color = RGB(245, 245, 245) ' Light gray
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .RowHeight = 25
        .Font.Name = "Consolas"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    
    ' Add Browse OBIEE Button
    Set btn = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, wsDash.Range("G5").Left + 10, wsDash.Range("G5").Top, 100, 25)
    With btn
        .Fill.ForeColor.RGB = RGB(0, 120, 212) ' Windows blue
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Browse..."
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "BrowseOBIEE"
    End With
    
    ' Add Browse ServiceMax Button
    Set btn = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, wsDash.Range("G7").Left + 10, wsDash.Range("G7").Top, 100, 25)
    With btn
        .Fill.ForeColor.RGB = RGB(0, 120, 212)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Browse..."
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "BrowseServiceMax"
    End With
    
    ' Add Run Button
    Set btn = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, wsDash.Range("B10").Left, wsDash.Range("B10").Top, 200, 40)
    With btn
        .Fill.ForeColor.RGB = RGB(40, 167, 69) ' Success green
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "RUN DATA CLEANSING"
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "RunDataCleansing"
    End With
    
    ' Adjust column widths
    wsDash.Columns("A").ColumnWidth = 3
    wsDash.Columns("B").ColumnWidth = 20
    
    wsDash.Activate
    wsDash.Range("A1").Select
    
    MsgBox "Dashboard and Config sheet updated! Add your valid countries to the Config sheet Column A.", vbInformation
End Sub
