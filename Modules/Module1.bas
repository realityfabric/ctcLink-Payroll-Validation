Public Function GetSheet(sheetName As String, Optional wb As Workbook) As Worksheet
    Debug.Print ("GetSheet(" & sheetName & ")")
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim ws As Worksheet, sheet As Worksheet
    Set sheet = Nothing
    For Each ws In wb.Sheets
        If sheetName = ws.Name Then
            Debug.Print ("Found: " & ws.Name & " = " & sheetName)
            Set sheet = ws
            Set GetSheet = ws
            Exit Function
        End If
    Next ws
    
    If sheet Is Nothing Then
        Debug.Print ("Sheet not found. Creating sheet.")
        Set sheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        sheet.Name = sheetName
        Set GetSheet = sheet
        Exit Function
    End If
    
    GetSheet = sheet
End Function

Public Function GetSheetLike(sheetName As String, Optional wb As Workbook) As Worksheet
    Debug.Print ("GetSheetLike(" & sheetName & ")")
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim ws As Worksheet, sheet As Worksheet
    Set sheet = Nothing
    For Each ws In wb.Sheets
        If sheetName Like ws.Name Then
            Debug.Print ("Found: " & ws.Name & " Like " & sheetName)
            Set sheet = ws
            Set GetSheetLike = ws
            Exit Function
        End If
    Next ws
    
    If sheet Is Nothing Then
        Debug.Print ("Sheet not found.")
    End If
    
    Set GetSheetLike = sheet
End Function

Public Function SetHeadersEJC(ws)
    ' Set headers for EJC List
    ws.Range("A1").Value = "Empl ID"
    ws.Range("B1").Value = "Name (LN,FN)"
    ws.Range("C1").Value = "Job Code"
End Function
Public Function SetHeadersAppointed(ws)
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

Public Function SetHeadersHourly(ws)
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

Public Function GetColumnLetterByNumber(columnNumber) As String
    ' Define array of columns
    Dim colArr(1 To 70) As String
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
    colArr(27) = "AA"
    colArr(28) = "AB"
    colArr(29) = "AC"
    colArr(30) = "AD"
    colArr(31) = "AE"
    colArr(32) = "AF"
    colArr(33) = "AG"
    colArr(34) = "AH"
    colArr(35) = "AI"
    colArr(36) = "AJ"
    colArr(37) = "AK"
    colArr(38) = "AL"
    colArr(39) = "AM"
    colArr(40) = "AN"
    colArr(41) = "AO"
    colArr(42) = "AP"
    colArr(43) = "AQ"
    colArr(44) = "AR"
    colArr(45) = "AS"
    colArr(46) = "AT"
    colArr(47) = "AU"
    colArr(48) = "AV"
    colArr(49) = "AW"
    colArr(50) = "AX"
    colArr(51) = "AY"
    colArr(52) = "AZ"
    colArr(53) = "BA"
    colArr(54) = "BB"
    colArr(55) = "BC"
    colArr(56) = "BD"
    colArr(57) = "BE"
    colArr(58) = "BF"
    colArr(59) = "BG"
    colArr(60) = "BH"
    colArr(61) = "BI"
    colArr(62) = "BJ"
    colArr(63) = "BK"
    colArr(64) = "BL"
    colArr(65) = "BM"
    colArr(66) = "BN"
    colArr(67) = "BO"
    colArr(68) = "BP"
    colArr(69) = "BQ"
    colArr(70) = "BR"
    
    GetColumnLetterByNumber = colArr(columnNumber)
End Function
Public Function CopyRange(wsCopy, startColNumCopy, endColNumCopy, startRowNumCopy, endRowNumCopy, wsDest, startColNumDest, startRowNumDest)
    Dim rg As Range
    
    Set rg = wsCopy.Range(GetColumnLetterByNumber(startColNumCopy) & startRowNumCopy & ":" & GetColumnLetterByNumber(endColNumCopy) & endRowNumCopy)
    wsDest.Range(GetColumnLetterByNumber(startColNumDest) & startRowNumDest).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
    
End Function


Public Function FindLastRowInSheet(ws) As Long
    ' based on https://stackoverflow.com/a/11169920
    Dim lastRow As Long
    
    With ws
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            Debug.Print ("CountA(.Cells) <> 0")
            lastRow = .Cells.Find(What:="*", _
                After:=.Range("A1"), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
        Else
            Debug.Print ("CountA(.Cells) = 0")
            lastRow = 1
        End If
    End With
    
    FindLastRowInSheet = lastRow

End Function

Public Function FindColumnByName(ws, columnName, Optional headerRow As Integer = 1) As Integer
    For Each c In ws.Range("A" & headerRow & ":ZZ" & headerRow)
        If c.Value = columnName Then
            FindColumnByName = c.Column
            Exit For
        Else
            FindColumnByName = -1
        End If
    Next c
End Function

Sub RefreshData()
    Debug.Print ("VBA Subroutine Main(): Start.")
    Dim rg As Range
    Dim ThatWorkbook As Workbook
    Dim destAppointed As Worksheet, destHourly As Worksheet, destOther As Worksheet, destEJC As Worksheet
    Dim copyAppointed As Worksheet, copyHourly As Worksheet, copyOther As Worksheet
        
    ' Ask if clearing Existing Data is acceptable
        ' If No: Exit Macro
    continue = MsgBox("Continuing will delete some of the data in this workbook before it begins, and will open and close other workbooks while it runs. Please save and close all other Excel Workbooks before continuing." & vbNewLine & vbNewLine & "Do you want to continue?", vbExclamation + vbYesNo + vbDefaultButton2, "Continue?")
    If continue <> 6 Then ' 6 is MsgBox "Yes"
        Debug.Print ("User declined to continue with script. Terminating Script.")
        End
    End If
        
    ' Get Path to Workbook
    Path = ThisWorkbook.Path & "\"
    Debug.Print ("Workbook Path: " & Path)
    
    Set destAppointed = GetSheet("Appointed")
    Set destHourly = GetSheet("Hourly")
    Set destOther = GetSheet("QHC_PY_PAY_CHECK_OTH_EARNS")
    Set destEJC = GetSheet("EJC List")
    
    ' Clear Existing Data
    Debug.Print ("Deleting Data from 'Appointed'")
    destAppointed.UsedRange.Delete
    Debug.Print ("Deleting Data from 'Hourly'")
    destHourly.UsedRange.Delete
    Debug.Print ("Deleting Data from 'QHC_PY_PAY_CHECK_TH_EARNS'")
    destOther.UsedRange.Delete
    Debug.Print ("Deleting Data from 'EJC List'")
    destEJC.UsedRange.Delete

    
    x = SetHeadersAppointed(destAppointed)
    x = SetHeadersHourly(destHourly)
    x = SetHeadersEJC(destEJC)
    
    
    ' Get List of Workbooks in Current Dir
    Filename = Dir(Path & "*.xlsx")
    ' For each workbook:
    Do While Filename <> ""
    Debug.Print (Filename)
    If Filename Like "*QHC_PY_PAY_CHECK_OTH_EARNS.xlsx" Then
        Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
        Set copySheet = Workbooks(Filename).Worksheets("Sheet1")
        Set rg = copySheet.UsedRange
        destOther.Range("A1").Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
    Else
        Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
            Set ThatWorkbook = Workbooks(Filename)
            Set copyAppointed = GetSheetLike("*Appointed*", ThatWorkbook)
            For Each ws In Workbooks(Filename).Worksheets
                Debug.Print ("Worksheet: " & ws.Name);
                If ws.Name Like "*Appointed*" Then
                    Debug.Print (vbTab & "copyAppointed")
                    Set copyAppointed = ws
                ElseIf ws.Name Like "*Hourly*" Then
                    Debug.Print (vbTab & "copyHourly")
                    Set copyHourly = ws
                Else
                    Debug.Print (vbTab & "Not Used")
                End If
            Next ws
            'Set copyAppointed = Workbooks(Filename).Worksheets("BDV&ITC Appointed")
            'Set copyHourly = Workbooks(Filename).Worksheets("BDV&ITC Hourly ")
            
            ' Find last non-empty row in copy and Find first empty row in This Workbook
            lCopyLastRow = FindLastRowInSheet(copyAppointed)
            lDestLastRow = FindLastRowInSheet(destAppointed) + 1 ' TODO: fix references so that the +1 is not in the definition
            
            ' Get header for Column A and Match header in This Workbook
            For i = 1 To 70
                a = GetColumnLetterByNumber(i)
                Debug.Print ("Scanning Column " & a & "..." & vbTab);
                copy_val = copyAppointed.Range(a & "1").Value
                If copy_val = "" Then
                    Debug.Print ("Blank Column Detected, Moving to Copy Step..." & vbTab)
                    Exit For
                End If
                dest_head = -1 ' -1 is an impossible column, indicating failure
                msg_string = "Column " & a & " (" & copy_val & "): No Match Detected!!!" ' msg_string will be updated if match is detected
                
                For Each c In destAppointed.Range("A1:CZ1")
                    If copy_val = c.Value Then
                        msg_string = "Column " & a & "(" & copy_val & ") matched with Column " & c.Column & "."
                        dest_head = c.Column
                    ElseIf Left(copy_val, 3) = c.Value Then
                        msg_string = "Column " & a & "(" & copy_val & ") matched with Column " & c.Column & "."
                        dest_head = c.Column
                    End If
                Next c
                
                If Len(msg_string) < 40 Then
                    Debug.Print (msg_string & vbTab & vbTab & vbTab);
                ElseIf Len(msg_string) < 44 Then
                    Debug.Print (msg_string & vbTab & vbTab);
                Else
                    Debug.Print (msg_string & vbTab);
                End If
                
                If dest_head = -1 Then
                    mb = MsgBox(msg_string, vbCritical)
                Else
                    ' Starting at first (fully) blank row in This Workbook:
                    ' Copy Column A to This Workbook in correct column
                    Debug.Print ("Copying Column " & a & "... ");
                    Set rg = copyAppointed.Range(a & "2:" & a & lCopyLastRow)
                    destAppointed.Range(GetColumnLetterByNumber(dest_head) & lDestLastRow).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
                    ' below commented out - does not copy values only, copies formulas, which break due to rearranging data
                    ' copyAppointed.Range(a & "2:" & a & lCopyLastRow).Copy _
                    '     destAppointed.Range(GetColumnLetterByNumber(dest_head) & lDestLastRow)
                    Debug.Print ("Column " & a & " Complete!")
                End If
            Next i
            
            ' HOURLY
            Debug.Print ("Working on worksheet: Hourly")
            destHourly.Activate
            
             ' Find last non-empty row in copy and Find first empty row in This Workbook
            lCopyLastRow = FindLastRowInSheet(copyHourly)
            lDestLastRow = FindLastRowInSheet(destHourly) + 1
            
            ' Get header for Column A and Match header in This Workbook
            For i = 1 To 70
                a = GetColumnLetterByNumber(i)
                Debug.Print ("Scanning Column " & a & "..." & vbTab);
                copy_val = copyHourly.Range(a & "1").Value
                If copy_val = "" Then
                    Debug.Print ("Blank Column Detected, Moving to Copy Step..." & vbTab)
                    Exit For
                End If
                dest_head = -1 ' -1 is an impossible column, indicating failure
                msg_string = "Column " & a & " (" & copy_val & "): No Match Detected!!!" ' msg_string will be updated if match is detected
                
                For Each c In destHourly.Range("A1:CZ1")
                    If copy_val = c.Value Then
                        msg_string = "Column " & a & "(" & copy_val & ") matched with Column " & c.Column & "."
                        dest_head = c.Column
                    ElseIf Right(c.Value, 5) = "Hours" Then
                        If Left(copy_val, 3) = Left(c.Value, 3) Then
                            msg_string = "Column " & a & "(" & copy_val & ") matched with Column " & c.Column & "."
                            dest_head = c.Column
                        End If
                    ElseIf Right(c.Value, 3) = "Pay" Then
                        If copy_val = "$ " & Left(c.Value, 3) & " $" Then
                            msg_string = "Column " & a & "(" & copy_val & ") matched with Column " & c.Column & "."
                            dest_head = c.Column
                        End If
                    End If
                Next c
                
                If Len(msg_string) < 40 Then
                    Debug.Print (msg_string & vbTab & vbTab & vbTab);
                ElseIf Len(msg_string) < 44 Then
                    Debug.Print (msg_string & vbTab & vbTab);
                Else
                    Debug.Print (msg_string & vbTab);
                End If
                
                If dest_head = -1 Then
                    mb = MsgBox(msg_string, vbCritical)
                Else
                    ' Starting at first (fully) blank row in This Workbook:
                    ' Copy Column A to This Workbook in correct column
                    copy_range_string = a & "2:" & a & lCopyLastRow
                    dest_range_start_string = GetColumnLetterByNumber(dest_head) & lDestLastRow
                    Debug.Print ("Copying " & copy_range_string & " to " & dest_range_start_string & "...");
                    Set rg = copyHourly.Range(copy_range_string)
                    destHourly.Range(dest_range_start_string).Resize(rg.Rows.Count, rg.Columns.Count).Cells.Value = rg.Cells.Value
                    'copyHourly.Range(a & "2:" & a & lCopyLastRow).Copy _
                    '    destHourly.Range(GetColumnLetterByNumber(dest_head) & lDestLastRow)
                    Debug.Print ("Column " & a & " Complete!")
                End If
            Next i
  
        End If
        Workbooks(Filename).Close SaveChanges:=False
        Debug.Print (Filename + " is closed." & vbNewLine)
        Filename = Dir()
    Loop
    
    
    
    Debug.Print ("VBA Subroutine Main(): End.")
    a = MsgBox("Workbook Refresh Complete!")
End Sub

Sub GenerateEmployeeList()
    Debug.Print ("GenerateEmployeeList(): Start.")
    Debug.Print ("Initializing Variables.")
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
    
    emplColAppointed = FindColumnByName(wsAppointed, "Empl ID")
    emplColHourly = FindColumnByName(wsHourly, "Empl ID")
    emplColEJC = FindColumnByName(wsEJC, "Empl ID")
    nameColAppointed = FindColumnByName(wsAppointed, "Name (LN,FN)")
    nameColHourly = FindColumnByName(wsHourly, "Name (LN,FN)")
    nameColEJC = FindColumnByName(wsEJC, "Name (LN,FN)")
    jobcodeAppointed = FindColumnByName(wsAppointed, "Job Code")
    jobcodeHourly = FindColumnByName(wsHourly, "Job Code")
    jobcodeEJC = FindColumnByName(wsEJC, "Job Code")
    
    ' Copy wsAppointed
    Debug.Print ("Copying wsAppointed to wsEJC")
    x = CopyRange(wsAppointed, emplColAppointed, emplColAppointed, 2, lastRowAppointed, wsEJC, 1, lastRowEJC + 1)
    x = CopyRange(wsAppointed, nameColAppointed, nameColAppointed, 2, lastRowAppointed, wsEJC, 2, lastRowEJC + 1)
    x = CopyRange(wsAppointed, jobcodeAppointed, jobcodeAppointed, 2, lastRowAppointed, wsEJC, 3, lastRowEJC + 1)
    Debug.Print ("Finding new last row for EJC")
    lastRowEJC = FindLastRowInSheet(wsEJC)
    
    ' Copy wsHourly
    Debug.Print ("Copying wsHourly to wsEJC")
    x = CopyRange(wsHourly, emplColHourly, emplColHourly, 2, lastRowHourly, wsEJC, 1, lastRowEJC + 1)
    x = CopyRange(wsHourly, nameColHourly, nameColHourly, 2, lastRowHourly, wsEJC, 2, lastRowEJC + 1)
    x = CopyRange(wsHourly, jobcodeHourly, jobcodeHourly, 2, lastRowHourly, wsEJC, 3, lastRowEJC + 1)
    Debug.Print ("Finding new last row for EJC")
    lastRowEJC = FindLastRowInSheet(wsEJC)
    
    
    ' Remove Duplicates
    With wsEJC
        With .Cells(1, 1).CurrentRegion
            .RemoveDuplicates Columns:=Array(1, 3), Header:=xlYes
        End With
    End With
    x = MsgBox("Employee/Job Code list has been generated.")
End Sub

Public Function GeneratePayrollSummarySheet(payPeriod As String)
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
    
    cPayAppointed = FindColumnByName(wsAppointed, payPeriod)
    cEmplAppointed = FindColumnByName(wsAppointed, "Empl ID")
    cJobCodeAppointed = FindColumnByName(wsAppointed, "Job Code")
    
    cPayHourly = FindColumnByName(wsHourly, payPeriod & " Pay")
    cEmplHourly = FindColumnByName(wsHourly, "Empl ID")
    cJobCodeHourly = FindColumnByName(wsHourly, "Job Code")
    
    cPayOther = FindColumnByName(wsOther, "Oth Earns", 2)
    cEmplOther = FindColumnByName(wsOther, "ID", 2)
    cJobCodeOther = FindColumnByName(wsOther, "Earn Code", 2)
    
    For Each c In wsPeriod.Range("D2:D" & lastRow)
        c.Value = "=SUMIFS(" & wsAppointed.Name & "!" & GetColumnLetterByNumber(cPayAppointed) & ":" & GetColumnLetterByNumber(cPayAppointed) & ", " _
            & wsAppointed.Name & "!" & GetColumnLetterByNumber(cEmplAppointed) & ":" & GetColumnLetterByNumber(cEmplAppointed) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsAppointed.Name & "!" & GetColumnLetterByNumber(cJobCodeAppointed) & ":" & GetColumnLetterByNumber(cJobCodeAppointed) & ", " _
            & "C" & c.Row & ")"
    Next c
    For Each c In wsPeriod.Range("E2:E" & lastRow)
        c.Value = "=SUMIFS(" & wsHourly.Name & "!" & GetColumnLetterByNumber(cPayHourly) & ":" & GetColumnLetterByNumber(cPayHourly) & ", " _
            & wsHourly.Name & "!" & GetColumnLetterByNumber(cEmplHourly) & ":" & GetColumnLetterByNumber(cEmplHourly) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsHourly.Name & "!" & GetColumnLetterByNumber(cJobCodeHourly) & ":" & GetColumnLetterByNumber(cJobCodeHourly) & ", " _
            & "C" & c.Row & ")"
    Next c
    For Each c In wsPeriod.Range("F2:F" & lastRow)
        c.Value = "=SUM(D" & c.Row & ":E" & c.Row & ")"
    Next c
    For Each c In wsPeriod.Range("G2:G" & lastRow)
        c.Value = "=SUMIFS(" & wsOther.Name & "!" & GetColumnLetterByNumber(cPayOther) & ":" & GetColumnLetterByNumber(cPayOther) & ", " _
            & wsOther.Name & "!" & GetColumnLetterByNumber(cEmplOther) & ":" & GetColumnLetterByNumber(cEmplOther) & ", " _
            & "TEXT(A" & c.Row & "," & Chr(34) & "0" & Chr(34) & "), " _
            & wsOther.Name & "!" & GetColumnLetterByNumber(cJobCodeOther) & ":" & GetColumnLetterByNumber(cJobCodeOther) & ", " _
            & "C" & c.Row & ")"
    Next c
    For Each c In wsPeriod.Range("H2:H" & lastRow)
        c.Value = "=G" & c.Row & " - F" & c.Row
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
    Debug.Print ("RemoveCanceledClasses(): Start.")
    Dim wsAppointed As Worksheet, wsHourly As Worksheet
    Dim lastRowAppoinetd As Long, lastRowHourly As Long
    Set wsAppointed = GetSheet("Appointed")
    Set wsHourly = GetSheet("Hourly")

    lastRowAppointed = FindLastRowInSheet(wsAppointed)
    lastRowHourly = FindLastRowInSheet(wsHourly)

    colCanceledAppointed = FindColumnByName(wsAppointed, "Canceled Class")
    colCanceledHourly = FindColumnByName(wsHourly, "Canceled Class")
    
    Debug.Print ("Appointed: Deleting Rows with Canceled Classes...")
    Dim r As Long ' r will store the row in the upcoming loop
    With wsAppointed.Range("A1:" & GetColumnLetterByNumber(colCanceledAppointed) & lastRowAppointed)
        For r = .Rows.Count To 1 Step -1
            If .Cells(r, colCanceledAppointed) = "Y" Then
                Debug.Print ("Deleting row " & r)
                .Rows(r).EntireRow.Delete
            End If
        Next r
    End With
    
    With wsHourly.Range("A1:" & GetColumnLetterByNumber(colCanceledHourly) & lastRowHourly)
        For r = .Rows.Count To 1 Step -1
            If .Cells(r, colCanceledHourly) = "Y" Then
                Debug.Print ("Deleting row " & r)
                .Rows(r).EntireRow.Delete
            End If
        Next r
    End With

    x = MsgBox("Canceled Classes removed.")
    Debug.Print ("RemoveCanceledClasses(): End.")
End Sub
