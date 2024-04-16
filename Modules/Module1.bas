Function DebugPrint(log As String, Optional timeStamp As Boolean = True, Optional newLine As Boolean = True)
    Dim filePath As String, dateFormat As String, logFileName As String
    Dim fs, f
    Const IOMODE = 8 ' ForAppending
    Const CREATE = True ' Create new file if no file exists
    Const TEXTFORMAT = 0 ' Write file as ASCII
    filePath = ThisWorkbook.Path & "\"
    currentDate = FORMAT(Date, "YYYY-MM-DD")
    currentTime = FORMAT(Now(), "hh:nn:ss")
    logFileName = "debuglog_" & currentDate & ".txt"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(filePath & logFileName, IOMODE, CREATE, TEXTFORMAT)
        
    Dim logMessage As String
    
    If timeStamp Then logMessage = currentTime & " - "
    logMessage = logMessage & log
    Debug.Print (logMessage);
    If newLine Then
        Debug.Print ("") ' create new line in Immediate window
        logMessage = logMessage & vbNewLine
    Else
        Debug.Print (""); ' prevent new line in Immediate window
    End If
        
    f.Write logMessage
    f.Close
End Function

Function GetSheet(sheetName As String, Optional wb As Workbook) As Worksheet
    DebugPrint ("GetSheet(" & sheetName & ")")
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim ws As Worksheet, sheet As Worksheet
    Set sheet = Nothing
    For Each ws In wb.Sheets
        If sheetName = ws.Name Then
            DebugPrint ("Found: " & ws.Name & " = " & sheetName)
            Set sheet = ws
            Set GetSheet = ws
            Exit Function
        End If
    Next ws
    
    If sheet Is Nothing Then
        DebugPrint ("Sheet not found. Creating sheet.")
        Set sheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        sheet.Name = sheetName
        Set GetSheet = sheet
        Exit Function
    End If
    
    Set GetSheet = sheet
End Function

Function GetSheetLike(sheetName As String, Optional wb As Workbook) As Worksheet
    DebugPrint ("GetSheetLike(" & sheetName & ")")
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim ws As Worksheet, sheet As Worksheet
    Set sheet = Nothing
    For Each ws In wb.Sheets
        If ws.Name Like sheetName Then
            DebugPrint ("Found: " & ws.Name & " Like " & sheetName)
            Set sheet = ws
            Set GetSheetLike = ws
            Exit Function
        End If
    Next ws
    
    If sheet Is Nothing Then
        DebugPrint ("Sheet not found.")
    End If
    
    Set GetSheetLike = sheet
End Function

Function SetHeadersEJC(ws)
    ' Set headers for EJC List
    ws.Range("A1").Value = "Empl ID"
    ws.Range("B1").Value = "Name (LN,FN)"
    ws.Range("C1").Value = "Job Code"
End Function
Function SetHeadersAppointed(ws)
    ' Add Headers to Appointed
    ws.Range("A1").Value = "Dept."
    ws.Range("B1").Value = "Class #"
    ws.Range("C1").Value = "Subject"
    ws.Range("D1").Value = "Catalog"
    ws.Range("E1").Value = "Description"
    ws.Range("F1").Value = "DEPT"
    ws.Range("G1").Value = "Empl ID"
    ws.Range("H1").Value = "Rcd#"
    ws.Range("I1").Value = "Name (LN,FN)"
    ws.Range("J1").Value = "Job Code"
    ws.Range("K1").Value = "Units"
    ws.Range("L1").Value = "FTE %"
    ws.Range("M1").Value = "Cntct hrs"
    ws.Range("N1").Value = "LAB/LEC"
    ws.Range("O1").Value = "Rate"
    ws.Range("P1").Value = "Total Pay"
    ws.Range("Q1").Value = "Combo Code"
    ws.Range("R1").Value = "Begin"
    ws.Range("S1").Value = "End"
    ws.Range("T1").Value = "Days"
    ws.Range("U1").Value = "Start Tm"
    ws.Range("V1").Value = "End Tm"
    ws.Range("W1").Value = "01A"
    ws.Range("X1").Value = "01B"
    ws.Range("Y1").Value = "02A"
    ws.Range("Z1").Value = "02B"
    ws.Range("AA1").Value = "03A"
    ws.Range("AB1").Value = "03B"
    ws.Range("AC1").Value = "04A"
    ws.Range("AD1").Value = "04B"
    ws.Range("AE1").Value = "05A"
    ws.Range("AF1").Value = "05B"
    ws.Range("AG1").Value = "06A"
    ws.Range("AH1").Value = "06B"
    ws.Range("AI1").Value = "07A"
    ws.Range("AJ1").Value = "07B"
    ws.Range("AK1").Value = "08A"
    ws.Range("AL1").Value = "08B"
    ws.Range("AM1").Value = "09A"
    ws.Range("AN1").Value = "09B"
    ws.Range("AO1").Value = "10A"
    ws.Range("AP1").Value = "10B"
    ws.Range("AQ1").Value = "11A"
    ws.Range("AR1").Value = "11B"
    ws.Range("AS1").Value = "12A"
    ws.Range("AT1").Value = "12B"
    ws.Range("AU1").Value = "Canceled Class"
End Function

Function SetHeadersHourly(ws)
    ' Add Headers to Hourly
    ws.Range("A1").Value = "Item"
    ws.Range("B1").Value = "Course"
    ws.Range("C1").Value = "Description"
    ws.Range("D1").Value = "DEPT"
    ws.Range("E1").Value = "Empl ID"
    ws.Range("F1").Value = "Rcd#"
    ws.Range("G1").Value = "Name (LN,FN)"
    ws.Range("H1").Value = "Job Code"
    ws.Range("I1").Value = "FTE %"
    ws.Range("J1").Value = "Cntct hrs"
    ws.Range("K1").Value = "LAB/LEC"
    ws.Range("L1").Value = "Rate"
    ws.Range("M1").Value = "Est Hrs"
    ws.Range("N1").Value = "Total Pay"
    ws.Range("O1").Value = "Combo Code"
    ws.Range("P1").Value = "Begin"
    ws.Range("Q1").Value = "End"
    ws.Range("R1").Value = "Days"
    ws.Range("S1").Value = "Start Tm"
    ws.Range("T1").Value = "End Tm"
    ws.Range("U1").Value = "Notes:"
    
    ws.Range("V1").Value = "01A Hours"
    ws.Range("W1").Value = "01A Pay"
    ws.Range("X1").Value = "01B Hours"
    ws.Range("Y1").Value = "01B Pay"
    
    ws.Range("Z1").Value = "02A Hours"
    ws.Range("AA1").Value = "02A Pay"
    ws.Range("AB1").Value = "02B Hours"
    ws.Range("AC1").Value = "02B Pay"
    
    ws.Range("AD1").Value = "03A Hours"
    ws.Range("AE1").Value = "03A Pay"
    ws.Range("AF1").Value = "03B Hours"
    ws.Range("AG1").Value = "03B Pay"
    
    ws.Range("AH1").Value = "04A Hours"
    ws.Range("AI1").Value = "04A Pay"
    ws.Range("AJ1").Value = "04B Hours"
    ws.Range("AK1").Value = "04B Pay"
    
    ws.Range("AL1").Value = "05A Hours"
    ws.Range("AM1").Value = "05A Pay"
    ws.Range("AN1").Value = "05B Hours"
    ws.Range("AO1").Value = "05B Pay"
    
    ws.Range("AP1").Value = "06A Hours"
    ws.Range("AQ1").Value = "06A Pay"
    ws.Range("AR1").Value = "06B Hours"
    ws.Range("AS1").Value = "06B Pay"
    
    ws.Range("AT1").Value = "07A Hours"
    ws.Range("AU1").Value = "07A Pay"
    ws.Range("AV1").Value = "07B Hours"
    ws.Range("AW1").Value = "07B Pay"
    
    ws.Range("AX1").Value = "08A Hours"
    ws.Range("AY1").Value = "08A Pay"
    ws.Range("AZ1").Value = "08B Hours"
    ws.Range("BA1").Value = "08B Pay"
    
    ws.Range("BB1").Value = "09A Hours"
    ws.Range("BC1").Value = "09A Pay"
    ws.Range("BD1").Value = "09B Hours"
    ws.Range("BE1").Value = "09B Pay"
    
    ws.Range("BF1").Value = "10A Hours"
    ws.Range("BG1").Value = "10A Pay"
    ws.Range("BH1").Value = "10B Hours"
    ws.Range("BI1").Value = "10B Pay"
    
    ws.Range("BJ1").Value = "11A Hours"
    ws.Range("BK1").Value = "11A Pay"
    ws.Range("BL1").Value = "11B Hours"
    ws.Range("BM1").Value = "11B Pay"
    
    ws.Range("BN1").Value = "12A Hours"
    ws.Range("BO1").Value = "12A Pay"
    ws.Range("BP1").Value = "12B Hours"
    ws.Range("BQ1").Value = "12B Pay"
    
    ws.Range("BR1").Value = "Canceled Class"
End Function

Function GetColumnLetterByNumber(columnNumber) As String
    ' Define array of columns
    Dim colArr(0 To 26) As String
    colArr(0) = ""
    colArr(1) = "A"
    colArr(2) = "B"
    colArr(3) = "C"
    colArr(4) = "D"
    colArr(5) = "E"
    colArr(6) = "F"
    colArr(7) = "G"
    colArr(8) = "H"
    colArr(9) = "I"
    colArr(10) = "J"
    colArr(11) = "K"
    colArr(12) = "L"
    colArr(13) = "M"
    colArr(14) = "N"
    colArr(15) = "O"
    colArr(16) = "P"
    colArr(17) = "Q"
    colArr(18) = "R"
    colArr(19) = "S"
    colArr(20) = "T"
    colArr(21) = "U"
    colArr(22) = "V"
    colArr(23) = "W"
    colArr(24) = "X"
    colArr(25) = "Y"
    colArr(26) = "Z"
    
    ' if the tensLetter is A and the onesLetter is B then: AB
    ' this will break if the column is bigger than ZZ
    tensLetter = colArr(Int((columnNumber - 1) / 26))
    onesLetter = colArr(((columnNumber - 1) Mod 26) + 1)
    
    GetColumnLetterByNumber = tensLetter & onesLetter
End Function
Function CopyRange(wsCopy, startColNumCopy, endColNumCopy, startRowNumCopy, endRowNumCopy, wsDest, startColNumDest, startRowNumDest)
    Dim rg As Range
    
    Set rg = wsCopy.Range(GetColumnLetterByNumber(startColNumCopy) & startRowNumCopy & ":" & GetColumnLetterByNumber(endColNumCopy) & endRowNumCopy)
    wsDest.Range(GetColumnLetterByNumber(startColNumDest) & startRowNumDest).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
    
End Function

Function FindLastRowInSheet(ws) As Long
    ' based on https://stackoverflow.com/a/11169920
    Dim lastRow As Long
    
    DebugPrint ("FindLastRowInSheet(" & ws.Name & ")")
    
    With ws
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            lastRow = .Cells.Find(What:="*", _
                After:=.Range("A1"), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
            DebugPrint ("Last row in " & ws.Name & " = " & lastRow)
        Else
            DebugPrint ("Worksheet " & ws.Name & " has no rows! Setting lastRow = 1")
            lastRow = 1
        End If
    End With
    
    FindLastRowInSheet = lastRow
End Function

' Query output typically has rows at the top before the header with the number of rows returned
' Sometimes there are also rows with query input parameters
' These rows typically only take up two columns
' By default, the following function can expect a row with at least 3 columns to be the header.
' By SBCTC standards, queries will not reduce the number of columns, but may increase them.
' Therefore, checking for a minimum number of columns, rather than a total number, is better.
' If no header row is detected, return -1
Function QueryOutputHeaderRow(ws As Worksheet, Optional minimumColumnsExpected As Integer = 3) As Long
    DebugPrint ("QueryOutputHeaderRow(" & ws.Name & ", " & minimumColumnsExpected & ")")
    Dim lastRow As Long, headerRow As Long, r As Long
    lastRow = FindLastRowInSheet(ws)
    Debug.Print ("Last Row: " & lastRow)
    
    If lastRow = 1 Then
        QueryOutputHeaderRow = 1
        DebugPrint ("QueryOutputHeaderRow = 1")
        Exit Function
    End If
    
    rangeString = GetColumnLetterByNumber(minimumColumnsExpected) & "1:" & GetColumnLetterByNumber(minimumColumnsExpected) & lastRow
    DebugPrint ("Checking Range: " & rangeString)
    For Each c In ws.Range(rangeString)
        If c.Value <> "" Then
            DebugPrint ("QueryOutputHeaderRow = " & c.Row)
            QueryOutputHeaderRow = c.Row
            Exit Function
        End If
    Next c
    QueryOutputHeaderRow = -1
    
    DebugPrint ("QueryOutputHeaderRow(): No Header Row Detected")
End Function

Function GetColumnNumberByName(ws, columnName, Optional headerRow As Long = 1) As Integer
    For Each c In ws.Range("A" & headerRow & ":ZZ" & headerRow)
        If c.Value = columnName Then
            GetColumnNumberByName = c.Column
            Exit For
        Else
            GetColumnNumberByName = -1
        End If
    Next c
End Function

Sub RefreshData()
    DebugPrint ("VBA Subroutine Main(): Start.")
    Dim rg As Range
    Dim ThatWorkbook As Workbook
    Dim destAppointed As Worksheet, destHourly As Worksheet, destOther As Worksheet, destEJC As Worksheet
    Dim copyAppointed As Worksheet, copyHourly As Worksheet, copyOther As Worksheet
    Dim copyLastRow As Long, destLastRow As Long
    Dim msgString As String
        
    ' Ask if clearing Existing Data is acceptable
        ' If No: Exit Macro
    continue = MsgBox("Continuing will delete some of the data in this workbook before it begins, and will open and close other workbooks while it runs. Please save and close all other Excel Workbooks before continuing." & vbNewLine & vbNewLine & "Do you want to continue?", vbExclamation + vbYesNo + vbDefaultButton2, "Continue?")
    If continue <> 6 Then ' 6 is MsgBox "Yes"
        DebugPrint ("User declined to continue with script. Terminating Script.")
        End
    End If
        
    ' Get Path to Workbook
    Path = ThisWorkbook.Path & "\"
    DebugPrint ("Workbook Path: " & Path)
    
    Set destAppointed = GetSheet("Appointed")
    Set destHourly = GetSheet("Hourly")
    Set destOther = GetSheet("QHC_PY_PAY_CHECK_OTH_EARNS")
    Set destEJC = GetSheet("EJC List")
    
    ' Clear Existing Data
    DebugPrint ("Deleting Data from 'Appointed'")
    destAppointed.UsedRange.Delete
    DebugPrint ("Deleting Data from 'Hourly'")
    destHourly.UsedRange.Delete
    DebugPrint ("Deleting Data from 'QHC_PY_PAY_CHECK_TH_EARNS'")
    destOther.UsedRange.Delete
    DebugPrint ("Deleting Data from 'EJC List'")
    destEJC.UsedRange.Delete

    
    x = SetHeadersAppointed(destAppointed)
    x = SetHeadersHourly(destHourly)
    x = SetHeadersEJC(destEJC)
    
    
    ' Get List of Workbooks in Current Dir
    Filename = Dir(Path & "*.xlsx")
    ' For each workbook:
    Do While Filename <> ""
    DebugPrint (Filename)
    If Filename Like "*QHC_PY_PAY_CHECK_OTH_EARNS.xlsx" Then
        Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
        Set copySheet = Workbooks(Filename).Worksheets("Sheet1")
        Set rg = copySheet.UsedRange
        destOther.Range("A1").Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
    Else
        Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
            Set ThatWorkbook = Workbooks(Filename)
            Set copyAppointed = GetSheetLike("*Appointed*", ThatWorkbook)
            Set copyHourly = GetSheetLike("*Hourly*", ThatWorkbook)
            
            promptArg = " sheet not found for workbook " _
                & ThatWorkbook.Name & vbNewLine & vbNewLine _
                & "Click OK to continue, or Cancel to exit the macro."
            buttonsArg = vbExclamation _
                + vbOKCancel _
                + vbApplicationModal _
                + vbMsgBoxSetForeground
                        
            If copyAppointed Is Nothing Then
                response = MsgBox("Appointed" & promptArg _
                    , buttonsArg _
                    , "Appointed Sheet Not Found")
                If response <> 1 Then ' 1 = vbOK
                    DebugPrint ("User declined to continue with script. Terminating Script.")
                    End
                End If
            End If
            
            If copyHourly Is Nothing Then
                response = MsgBox("Hourly" & promptArg _
                     , buttonsArg _
                    , "Hourly Sheet Not Found")
                If response <> 1 Then ' 1 = vbOK
                    DebugPrint ("User declined to continue with script. Terminating Script.")
                    End
                End If
            End If
            
            ' Find last non-empty row in copy and Find first empty row in This Workbook
            
            If Not copyAppointed Is Nothing Then
                copyLastRow = FindLastRowInSheet(copyAppointed)
                destLastRow = FindLastRowInSheet(destAppointed) + 1 ' TODO: fix references so that the +1 is not in the definition
            
                ' Get header for Column A and Match header in This Workbook
                For i = 1 To 70
                    a = GetColumnLetterByNumber(i)
                    x = DebugPrint("Scanning Column " & a & "..." & vbTab, True, False)
                    copyVal = copyAppointed.Range(a & "1").Value
                    If copyVal = "" Then
                        x = DebugPrint("Blank Column Detected, Moving to Copy Step..." & vbTab, False, True)
                        Exit For
                    End If
                    dest_head = -1 ' -1 is an impossible column, indicating failure
                    msgString = "Column " & a & " (" & copyVal & "): No Match Detected!!!" ' msgString will be updated if match is detected
                    
                    For Each c In destAppointed.Range("A1:CZ1")
                        If copyVal = c.Value Then
                            msgString = "Column " & a & "(" & copyVal & ") matched with Column " & c.Column & "."
                            dest_head = c.Column
                        ElseIf Left(copyVal, 3) = c.Value Then
                            msgString = "Column " & a & "(" & copyVal & ") matched with Column " & c.Column & "."
                            dest_head = c.Column
                        End If
                    Next c
                    
                    If Len(msgString) < 40 Then
                        x = DebugPrint(msgString & vbTab & vbTab & vbTab, False, False)
                    ElseIf Len(msgString) < 44 Then
                        x = DebugPrint(msgString & vbTab & vbTab, False, False)
                    Else
                        x = DebugPrint(msgString & vbTab, False, False)
                    End If
                    
                    If dest_head = -1 Then
                        mb = MsgBox(msgString, vbCritical)
                    Else
                        ' Starting at first (fully) blank row in This Workbook:
                        ' Copy Column A to This Workbook in correct column
                        x = DebugPrint("Copying Column " & a & "... ", False, False)
                        Set rg = copyAppointed.Range(a & "2:" & a & copyLastRow)
                        destAppointed.Range(GetColumnLetterByNumber(dest_head) & destLastRow).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
                        ' below commented out - does not copy values only, copies formulas, which break due to rearranging data
                        ' copyAppointed.Range(a & "2:" & a & copyLastRow).Copy _
                        '     destAppointed.Range(GetColumnLetterByNumber(dest_head) & destLastRow)
                        x = DebugPrint("Column " & a & " Complete!", False, True)
                    End If
                Next i
            End If ' End of work on Appointed sheet
            
            ' HOURLY
            If Not copyHourly Is Nothing Then
                DebugPrint ("Working on worksheet: Hourly")
                destHourly.Activate
                
                 ' Find last non-empty row in copy and Find first empty row in This Workbook
                copyLastRow = FindLastRowInSheet(copyHourly)
                destLastRow = FindLastRowInSheet(destHourly) + 1
                
                ' Get header for Column A and Match header in This Workbook
                For i = 1 To 70
                    a = GetColumnLetterByNumber(i)
                    x = DebugPrint("Scanning Column " & a & "..." & vbTab, True, False)
                    copyVal = copyHourly.Range(a & "1").Value
                    If copyVal = "" Then
                        x = DebugPrint("Blank Column Detected, Moving to Copy Step..." & vbTab, False, True)
                        Exit For
                    End If
                    dest_head = -1 ' -1 is an impossible column, indicating failure
                    msgString = "Column " & a & " (" & copyVal & "): No Match Detected!!!" ' msgString will be updated if match is detected
                    
                    For Each c In destHourly.Range("A1:CZ1")
                        If copyVal = c.Value Then
                            msgString = "Column " & a & "(" & copyVal & ") matched with Column " & c.Column & "."
                            dest_head = c.Column
                        ElseIf Right(c.Value, 5) = "Hours" Then
                            If Left(copyVal, 3) = Left(c.Value, 3) Then
                                msgString = "Column " & a & "(" & copyVal & ") matched with Column " & c.Column & "."
                                dest_head = c.Column
                            End If
                        ElseIf Right(c.Value, 3) = "Pay" Then
                            If copyVal = "$ " & Left(c.Value, 3) & " $" Then
                                msgString = "Column " & a & "(" & copyVal & ") matched with Column " & c.Column & "."
                                dest_head = c.Column
                            End If
                        End If
                    Next c
                    
                    If Len(msgString) < 40 Then
                        x = DebugPrint(msgString & vbTab & vbTab & vbTab, False, False)
                    ElseIf Len(msgString) < 44 Then
                        x = DebugPrint(msgString & vbTab & vbTab, False, False)
                    Else
                        x = DebugPrint(msgString & vbTab, False, False)
                    End If
                    
                    If dest_head = -1 Then
                        mb = MsgBox(msgString, vbCritical)
                    Else
                        ' Starting at first (fully) blank row in This Workbook:
                        ' Copy Column A to This Workbook in correct column
                        copy_range_string = a & "2:" & a & copyLastRow
                        dest_range_start_string = GetColumnLetterByNumber(dest_head) & destLastRow
                        x = DebugPrint("Copying " & copy_range_string & " to " & dest_range_start_string & "...", False, False)
                        Set rg = copyHourly.Range(copy_range_string)
                        destHourly.Range(dest_range_start_string).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
                        'copyHourly.Range(a & "2:" & a & copyLastRow).Copy _
                        '    destHourly.Range(GetColumnLetterByNumber(dest_head) & destLastRow)
                        x = DebugPrint("Column " & a & " Complete!", False, True)
                    End If
                Next i
            End If ' End of work on Hourly sheet
        End If
        Workbooks(Filename).Close SaveChanges:=False
        DebugPrint (Filename + " is closed." & vbNewLine)
        Filename = Dir()
    Loop
    
    
    
    DebugPrint ("VBA Subroutine Main(): End.")
    a = MsgBox("Workbook Refresh Complete!")
End Sub

Sub GenerateEmployeeList()
    DebugPrint ("GenerateEmployeeList(): Start.")
    Dim wsAppointed As Worksheet, wsHourly As Worksheet, wsEJC As Worksheet
    Dim lastRowAppointed As Long, lastRowHourly As Long, lastRowEJC As Long
    Dim emplColAppointed As Integer, emplColHourly As Integer, emplColEJC As Integer
    Dim nameColAppointed As Integer, nameColHourly As Integer, nameColEJC As Integer
    Dim jobcodeAppointed As Integer, jobcodeHourly As Integer, jobcodeEJC As Integer
           
    Set wsAppointed = GetSheet("Appointed")
    Set wsHourly = GetSheet("Hourly")
    Set wsEJC = GetSheet("EJC List")
    
    lastRowAppointed = FindLastRowInSheet(wsAppointed)
    lastRowHourly = FindLastRowInSheet(wsHourly)
    lastRowEJC = FindLastRowInSheet(wsEJC)
    
    emplColAppointed = GetColumnNumberByName(wsAppointed, "Empl ID")
    emplColHourly = GetColumnNumberByName(wsHourly, "Empl ID")
    emplColEJC = GetColumnNumberByName(wsEJC, "Empl ID")
    nameColAppointed = GetColumnNumberByName(wsAppointed, "Name (LN,FN)")
    nameColHourly = GetColumnNumberByName(wsHourly, "Name (LN,FN)")
    nameColEJC = GetColumnNumberByName(wsEJC, "Name (LN,FN)")
    jobcodeAppointed = GetColumnNumberByName(wsAppointed, "Job Code")
    jobcodeHourly = GetColumnNumberByName(wsHourly, "Job Code")
    jobcodeEJC = GetColumnNumberByName(wsEJC, "Job Code")
    
    ' Copy wsAppointed
    DebugPrint ("Copying wsAppointed to " & wsEJC.Name)
    x = CopyRange(wsAppointed, emplColAppointed, emplColAppointed, 2, lastRowAppointed, wsEJC, 1, lastRowEJC + 1)
    x = CopyRange(wsAppointed, nameColAppointed, nameColAppointed, 2, lastRowAppointed, wsEJC, 2, lastRowEJC + 1)
    x = CopyRange(wsAppointed, jobcodeAppointed, jobcodeAppointed, 2, lastRowAppointed, wsEJC, 3, lastRowEJC + 1)
    lastRowEJC = FindLastRowInSheet(wsEJC)
    
    ' Copy wsHourly
    DebugPrint ("Copying wsHourly to " & wsEJC.Name)
    x = CopyRange(wsHourly, emplColHourly, emplColHourly, 2, lastRowHourly, wsEJC, 1, lastRowEJC + 1)
    x = CopyRange(wsHourly, nameColHourly, nameColHourly, 2, lastRowHourly, wsEJC, 2, lastRowEJC + 1)
    x = CopyRange(wsHourly, jobcodeHourly, jobcodeHourly, 2, lastRowHourly, wsEJC, 3, lastRowEJC + 1)
    lastRowEJC = FindLastRowInSheet(wsEJC)
    
    
    ' Remove Duplicates and Blanks
    With wsEJC
        DebugPrint ("Deleting Duplicates in " & .Name)
        With .Range("A1", "C" & lastRowEJC)
            .RemoveDuplicates Columns:=Array(1, 3), Header:=xlYes
        End With
        lastRowEJC = FindLastRowInSheet(wsEJC)
        DebugPrint ("Deleting Blank Rows in " & .Name)
        With .Range("A1", "C" & lastRowEJC)
            For r = .Rows.Count To 1 Step -1 ' https://spreadsheetplanet.com/excel-vba/delete-blank-rows/
                If Application.WorksheetFunction.CountA(.Rows(r)) = 0 Then
                    .Rows(r).Delete
                End If
            Next r
        End With
    End With
    
    x = MsgBox("Employee/Job Code list has been generated.")
End Sub

Function GeneratePayrollSummarySheet(payPeriod As String)
    Dim wsEJC As Worksheet, wsPeriod As Worksheet, wsAppointed As Worksheet, wsHourly As Worksheet, wsOther As Worksheet
    Dim lastRowEJC As Long
    
    Set wsEJC = GetSheet("EJC List")
    Set wsPeriod = GetSheet("Summary")
    Set wsAppointed = GetSheet("Appointed")
    Set wsHourly = GetSheet("Hourly")
    Set wsOther = GetSheet("QHC_PY_PAY_CHECK_OTH_EARNS")
    lastRow = FindLastRowInSheet(wsEJC)
    
    wsPeriod.UsedRange.Delete
    
    x = CopyRange(wsEJC, 1, 3, 1, lastRow, wsPeriod, 1, 1)
    wsPeriod.Range("D1").Value = "Sum Appointed"
    wsPeriod.Range("E1").Value = "Sum Hourly"
    wsPeriod.Range("F1").Value = "Sum Total"
    wsPeriod.Range("G1").Value = "Other Earns"
    wsPeriod.Range("H1").Value = "DIFF"
    wsPeriod.Range("I1").Value = "Has Difference?"
    
    cPayAppointed = GetColumnNumberByName(wsAppointed, payPeriod)
    cEmplAppointed = GetColumnNumberByName(wsAppointed, "Empl ID")
    cJobCodeAppointed = GetColumnNumberByName(wsAppointed, "Job Code")
    
    cPayHourly = GetColumnNumberByName(wsHourly, payPeriod & " Pay")
    cEmplHourly = GetColumnNumberByName(wsHourly, "Empl ID")
    cJobCodeHourly = GetColumnNumberByName(wsHourly, "Job Code")
    
    Dim headerRowOther As Long
    headerRowOther = QueryOutputHeaderRow(wsOther)
    cPayOther = GetColumnNumberByName(wsOther, "Oth Earns", headerRowOther)
    cEmplOther = GetColumnNumberByName(wsOther, "ID", headerRowOther)
    cJobCodeOther = GetColumnNumberByName(wsOther, "Earn Code", headerRowOther)
    
    For Each c In wsPeriod.Range("D2:D" & lastRow)
        c.Value = "=ROUND(SUMIFS(" & wsAppointed.Name & "!" & GetColumnLetterByNumber(cPayAppointed) & ":" & GetColumnLetterByNumber(cPayAppointed) & ", " _
            & wsAppointed.Name & "!" & GetColumnLetterByNumber(cEmplAppointed) & ":" & GetColumnLetterByNumber(cEmplAppointed) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsAppointed.Name & "!" & GetColumnLetterByNumber(cJobCodeAppointed) & ":" & GetColumnLetterByNumber(cJobCodeAppointed) & ", " _
            & "C" & c.Row & "),4)"
    Next c
    For Each c In wsPeriod.Range("E2:E" & lastRow)
        c.Value = "=ROUND(SUMIFS(" & wsHourly.Name & "!" & GetColumnLetterByNumber(cPayHourly) & ":" & GetColumnLetterByNumber(cPayHourly) & ", " _
            & wsHourly.Name & "!" & GetColumnLetterByNumber(cEmplHourly) & ":" & GetColumnLetterByNumber(cEmplHourly) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsHourly.Name & "!" & GetColumnLetterByNumber(cJobCodeHourly) & ":" & GetColumnLetterByNumber(cJobCodeHourly) & ", " _
            & "C" & c.Row & "),4)"
    Next c
    For Each c In wsPeriod.Range("F2:F" & lastRow)
        c.Value = "=SUM(D" & c.Row & ":E" & c.Row & ")"
    Next c
    For Each c In wsPeriod.Range("G2:G" & lastRow)
        c.Value = "=ROUND(SUMIFS(" & wsOther.Name & "!" & GetColumnLetterByNumber(cPayOther) & ":" & GetColumnLetterByNumber(cPayOther) & ", " _
            & wsOther.Name & "!" & GetColumnLetterByNumber(cEmplOther) & ":" & GetColumnLetterByNumber(cEmplOther) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsOther.Name & "!" & GetColumnLetterByNumber(cJobCodeOther) & ":" & GetColumnLetterByNumber(cJobCodeOther) & ", " _
            & "TRIM(C" & c.Row & ")),4)"
    Next c
    For Each c In wsPeriod.Range("H2:H" & lastRow)
        c.Value = "=ROUND(G" & c.Row & " - F" & c.Row & ",2)"
    Next c
    For Each c In wsPeriod.Range("I2:I" & lastRow)
        c.Value = "=H" & c.Row & "<> 0"
    Next c
    
End Function

Sub GeneratePayrollSummary()
    Dim inputString As String
    inputString = InputBox("Please enter the 3-digit Payroll Period to be summarized.", "Payroll Summary Generator: Enter Payroll Period")
    x = GeneratePayrollSummarySheet(inputString)
    x = MsgBox("Payroll Summary has been generated.")
End Sub

Sub RemoveCanceledClasses()
    DebugPrint ("RemoveCanceledClasses(): Start.")
    Dim wsAppointed As Worksheet, wsHourly As Worksheet
    Dim lastRowAppoinetd As Long, lastRowHourly As Long
    Set wsAppointed = GetSheet("Appointed")
    Set wsHourly = GetSheet("Hourly")

    lastRowAppointed = FindLastRowInSheet(wsAppointed)
    lastRowHourly = FindLastRowInSheet(wsHourly)

    colCanceledAppointed = GetColumnNumberByName(wsAppointed, "Canceled Class")
    colCanceledHourly = GetColumnNumberByName(wsHourly, "Canceled Class")
    
    DebugPrint ("Appointed: Deleting Rows with Canceled Classes...")
    Dim r As Long ' r will store the row in the upcoming loop
    With wsAppointed.Range("A1:" & GetColumnLetterByNumber(colCanceledAppointed) & lastRowAppointed)
        For r = .Rows.Count To 1 Step -1
            If .Cells(r, colCanceledAppointed) = "Y" Then
                DebugPrint ("Deleting row " & r)
                .Rows(r).EntireRow.Delete
            End If
        Next r
    End With

    With wsHourly.Range("A1:" & GetColumnLetterByNumber(colCanceledHourly) & lastRowHourly)
        For r = .Rows.Count To 1 Step -1
            If .Cells(r, colCanceledHourly) = "Y" Then
                DebugPrint ("Deleting row " & r)
                .Rows(r).EntireRow.Delete
            End If
        Next r
    End With

    x = MsgBox("Canceled Classes removed.")
    DebugPrint ("RemoveCanceledClasses(): End.")
End Sub

Function CleanWorkbook()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Start Here" Then ws.Delete
    Next
    Application.DisplayAlerts = True
End Function

Function TestGetColumnLetterByNumber() As Boolean
    TestGetColumnLetterByNumber = False
    Debug.Print (Date & " " & Time & " - TestGetColumnLetterByNumber()")
    Debug.Assert (GetColumnLetterByNumber(1) = "A")
    Debug.Assert (GetColumnLetterByNumber(26 * 1) = "Z")
    Debug.Assert (GetColumnLetterByNumber(27) = "AA")
    Debug.Assert (GetColumnLetterByNumber(26 * 2) = "AZ")
    Debug.Assert (GetColumnLetterByNumber(26 * 3) = "BZ")
    Debug.Assert (GetColumnLetterByNumber(26 * 4) = "CZ")
    Debug.Assert (GetColumnLetterByNumber(26 * 5) = "DZ")
    Debug.Assert (GetColumnLetterByNumber(26 * 26) = "YZ")
    Debug.Assert (GetColumnLetterByNumber(26 * 27) = "ZZ")
    Debug.Print (Date & "" & Time & " - TestColumnLetterByNumber() Completed Successfully.")
    TestGetColumnLetterByNumber = True
End Function

Function TestCleanWorkbook() As Boolean
    Debug.Print Date & " " & Time & " - TestCleanWorkbook()"
    TestCleanWorkbook = False
    
    x = CleanWorkbook()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Debug.Assert ws.Name = "Start Here"
    Next ws
    
    TestCleanWorkbook = True
    Debug.Print Date & " " & Time & " - TestCleanWorkbook() Completed Successfully."
End Function

Function TestGetSheet() As Boolean
    Debug.Print (Date & " " & Time & " - TestGetSheet()")
    TestGetSheet = False
    x = CleanWorkbook()
    
    Dim wb As Workbook, ws As Worksheet, tws1 As Worksheet, tws2 As Worksheet
    Set wb = ThisWorkbook
        
    Set tws1 = GetSheet("Test WS 1")
    Debug.Assert tws1.Name = "Test WS 1"
    Debug.Assert wb.Sheets.Count = 2
    Set tws2 = GetSheet("Test WS 2")
    Debug.Assert tws2.Name = "Test WS 2"
    Debug.Assert wb.Sheets.Count = 3
    
    Debug.Print (Date & " " & Time & " - TestGetSheet() Completed Successfully.")
    TestGetSheet = True
End Function

Function TestGetSheetLike_SheetFound() As Boolean
    Debug.Print Date & " " & Time & " - TestGetSheetLike_SheetFound()"
    TestGetSheetLike_SheetFound = False
    
    x = CleanWorkbook()
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook
    
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Test WS 1"
    Set ws = Nothing
    Debug.Assert ws Is Nothing
    
    Set ws = GetSheetLike("Test WS 1")
    Debug.Assert ws.Name = "Test WS 1"
    Set ws = Nothing
    Debug.Assert ws Is Nothing
    
    Set ws = GetSheetLike("Test*")
    Debug.Assert ws.Name = "Test WS 1"
    Set ws = Nothing
    Debug.Assert ws Is Nothing

    Set ws = GetSheetLike("*Test*")
    Debug.Assert ws.Name = "Test WS 1"
    Set ws = Nothing
    Debug.Assert ws Is Nothing
   
    x = CleanWorkbook()
    
    TestGetSheetLike_SheetFound = True
    Debug.Print Date & " " & Time & " - TestGetSheetLike_SheetFound() Completed Successfully."
End Function

Function TestGetSheetLike_SheetNotFound() As Boolean
    TestGetSheetLike_SheetNotFound = False
    
    x = CleanWorkbook()
    Dim ws As Worksheet
    Set ws = GetSheetLike("This Will Fail")
    Debug.Assert ws Is Nothing
    
    TestGetSheetLike_SheetNotFound = True
End Function

Function TestGetSheetLike_CapitalizationMismatch() As Boolean
    Debug.Print Date & " " & Time & " - TestGetSheetLike_CapitalizationMismatch()"
    TestGetSheetLike_CapitalizationMismatch = False
    
    x = CleanWorkbook()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook
    
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Test WS 1"
    Set ws = Nothing
    Debug.Assert ws Is Nothing
    
    Set ws = GetSheetLike("TEST WS 1")
    Debug.Assert Not ws Is Nothing
    Debug.Assert ws.Name = "Test WS 1"
    Set ws = Nothing
    
    x = CleanWorkbook()
    
    TestGetSheetLike_CapitalizationMismatch = True
    Debug.Print Date & " " & Time & " - TestGetSheetLike_CapitalizationMismatch() Completed Successfully."
End Function

Sub TestModule()
    Dim counter As Integer
    counter = 0
    Debug.Print "Deleting Worksheets."
    x = CleanWorkbook()
    Debug.Print ("Running Tests... (" & Date & " " & Time & ")")
    
    Debug.Assert TestCleanWorkbook()
    counter = counter + 1
    Debug.Assert TestGetSheet()
    counter = counter + 1
    Debug.Assert TestGetSheetLike_SheetFound()
    counter = counter + 1
    Debug.Assert TestGetSheetLike_SheetNotFound()
    counter = counter + 1
    Debug.Assert TestGetSheetLike_CapitalizationMismatch()
    counter = counter + 1
    Debug.Assert TestGetColumnLetterByNumber()
    counter = counter + 1
    
    Debug.Print (Date & " " & Time & " - Completed " & counter & " tests successfully.")
End Sub
