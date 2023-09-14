Attribute VB_Name = "Phase2_New"
Option Explicit
'Public ShName As String
Sub Generate_Output_Phase2()

    Dim TR_Wrkbook As Workbook: Dim GOS_Wrkbook As Workbook: Dim Mapping_WrkBook As Workbook: Dim Summary_Workbook As Workbook
    Dim Path As String: Dim LastRow As Long: Dim Rng As Range: Dim Cell As Range: Dim k
    Dim Fnd As Range: Dim Response As Integer: Dim WrkSht As Worksheet
    Dim LastCol As Long: Dim LastRow_Summary As Long: Dim Month_Num As Integer:
    Dim Temp: Dim Wrksht_Map As Worksheet: Dim Col_Num As Integer
    Dim ADC_Wrkbk As Workbook: Dim Followup_Phase_A1 As Workbook: Dim Followup_Phase_A2 As Workbook
    Dim Phase1_Wrksht As Worksheet: Dim FollowUp_Wrksht As Worksheet
    
    Response = MsgBox("You have clicked Phase 2 button, if agree click Yes else click No", vbYesNo, "Confirmation")
    
    If Response = vbNo Then
        MsgBox "Since response is No: Exiting Macro"
        GoTo ExitHere
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Calculation = xlCalculationManual
    End With
    
    'Set Mapping sheet
    Set Wrksht_Map = ThisWorkbook.Sheets("Mapping")
    
    'Open TR Status report file
    MsgBox "Please choose TR Status report file"
    Call File_Picker_Fun(Path, "Please choose TR Status report file")
    If Path <> "" Then
        Set TR_Wrkbook = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose TR Status report file :Exiting Macro"
        GoTo ExitHere
    End If
    
    'Open Global Organizer status report file
    Path = ""
    MsgBox "Please choose Global Organizer status report file"
    Call File_Picker_Fun(Path, "Please choose Global Organizer status report file")
    If Path <> "" Then
        Set GOS_Wrkbook = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Global Organizer status report file :Exiting Macro"
        TR_Wrkbook.Close
        GoTo ExitHere
    End If
           
    'Open Summary report file
    Path = ""
    MsgBox "Please choose Summary report file"
    Call File_Picker_Fun(Path, "Please choose Summary report file")
    If Path <> "" Then
        Set Summary_Workbook = Workbooks.Open(Path)
    Else
        MsgBox "You did not Summary report file :Exiting Macro"
        TR_Wrkbook.Close
        GOS_Wrkbook.Close
        GoTo ExitHere
    End If
    
    With ThisWorkbook
    
        'Check whether All columns are available or not with global report
        With .Sheets("Mapping")
            LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
            Set Rng = .Range("B3:B" & LastRow)
            For Each Cell In Rng
                .Activate
                Set Fnd = GOS_Wrkbook.ActiveSheet.Range("A1:ZZ1").Find(what:=Trim(Cell.Value), LookAt:=xlWhole)
                If Fnd Is Nothing Then
                    MsgBox (Cell.Value & " column is not available in Global report: Exiting Macro!")
                    TR_Wrkbook.Close
                    Summary_Workbook.Close
                    GOS_Wrkbook.Close
                    GoTo ExitHere
                End If
            Next Cell
        End With
        
        'Check whether All columns are available or not with TR Status report
        With .Sheets("Mapping")
            LastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
            Set Rng = .Range("E3:E" & LastRow)
            For Each Cell In Rng
                .Activate
                Set Fnd = TR_Wrkbook.ActiveSheet.Range("A1:ZZ1").Find(what:=Trim(Cell.Value), LookAt:=xlWhole)
                If Fnd Is Nothing Then
                    MsgBox (Cell.Value & " column is not available in TR Status report: Exiting Macro!")
                    TR_Wrkbook.Close
                    Summary_Workbook.Close
                    GOS_Wrkbook.Close
                    GoTo ExitHere
                End If
            Next Cell
        End With
    End With
    
   'With Global Organizer status
   With GOS_Wrkbook.ActiveSheet
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("B3").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="<>United Kingdom"
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("B4").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("B5").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:=Array("_Test_GlobalOne Test, Inc", "_Test_GlobalOne Test, Inc 1", "_TEST_2 CO", "_Test_RPA", "_QA_Engagement1"), Operator:=xlFilterValues
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
   
   End With
   
   
   'With TR Status report file
   With TR_Wrkbook.ActiveSheet
        For Each WrkSht In Summary_Workbook.Sheets
            If WrkSht.Name = "Phase-2_ADC" Then
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                If WrkSht.Range("A2").Value <> "" Then
                    .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Range("B1").Value = "ID's from GOS"
                    .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]" & WrkSht.Name & "'!C1:C2,1,FALSE),1)"
                    .Activate
                    If LastRow <> 2 Then
                        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
                    End If
                    
                    Application.Calculation = xlCalculationAutomatic
                    Application.Calculation = xlCalculationManual
                    
                    .Range("$A$1:$Y$" & LastRow).AutoFilter Field:=2, Criteria1:="<>1"
                    On Error Resume Next
                    .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                    On Error GoTo 0
                    .Columns("B:B").Delete Shift:=xlToLeft
                End If
                Exit For
            End If
        Next WrkSht
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("B1").Value = "ID's from GOS"
        .Activate
        .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'" & GOS_Wrkbook.Name & "'!C1:C2,1,FALSE),1)"
        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
        .Columns("B:B").Copy
        .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range("$A$1:$Y$" & LastRow).AutoFilter Field:=2, Criteria1:=1
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E3").Value, .Range("A1:ZZ1"), 0)
        .Range("A2:A" & LastRow).AutoFilter Field:=Col_Num, Criteria1:=Array("No Filing Requirement", "No Longer Authorized", "Decliner", "Prepared By Other", "Return Filed", "Return Sent", "Missing Info Requested", "MI Reminder 1", "MI Reminder 2"), Operator:=xlFilterValues
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
    End With
    
    GOS_Wrkbook.Close
    
    'Add 3 Workbooks for Followup_Phase-A1, Followup_Phase-A2 & ADC

    Set Followup_Phase_A1 = Workbooks.Add
    Followup_Phase_A1.ActiveSheet.Name = "Phase-2_Followup-A1"
    
    Set Followup_Phase_A2 = Workbooks.Add
    Followup_Phase_A2.ActiveSheet.Name = "Phase-2_Followup-A2"
    
    Set ADC_Wrkbk = Workbooks.Add
    ADC_Wrkbk.ActiveSheet.Name = "Phase-2_ADC"
    
    'Copy Data from TR report to Individual File
    With TR_Wrkbook.ActiveSheet
        .Activate
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        .Columns("B:B").Delete Shift:=xlToLeft
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=Followup_Phase_A1.Sheets(1).Range("A1")
        .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=Followup_Phase_A2.Sheets(1).Range("A1")
        .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=ADC_Wrkbk.Sheets(1).Range("A1")
    End With

    'Combain Phase-A1 & Phase-A2 in Phase-1 Tab, FollowUp-A1 & FollowUp-A2 in FollowUp Tab..... Delete Later
    With Summary_Workbook
        Set Phase1_Wrksht = .Sheets.Add
            Phase1_Wrksht.Name = "Phase1"
        Set FollowUp_Wrksht = .Sheets.Add
            FollowUp_Wrksht.Name = "Followup"
        
        'For Phase 1
        With .Sheets("Phase-A1")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=Phase1_Wrksht.Range("A1")
        End With
        
        Temp = Phase1_Wrksht.Cells(Phase1_Wrksht.Rows.Count, "A").End(xlUp).Row
        
        With .Sheets("Phase-A2")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            If LastRow = 1 Then LastRow = 2
            LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Phase1_Wrksht.Range("A" & Temp + 1)
        End With
        
    End With
    
    
    'Phase-2_Followup-A1
    With Followup_Phase_A1.ActiveSheet
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        'Complete & Return Organizer Date Complete
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E4").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        'All Data Complete Date Complete
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E5").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="<>"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0

            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
        'Keep data which is only existing in Phase1
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B1").Value = "ID's from GOS"
        .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]Phase1'!C1:C2,1,FALSE),1)"
        .Activate
        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
        
        'Exclude current month data
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        'LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B1").Value = "Month"
        .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]Phase1'!C1:C9,9,FALSE),1)"
        '.Activate
        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
        
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
        
        .Columns("B:C").Copy
        .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=3, Criteria1:="1"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0

            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
        Month_Num = Month(Date)
        .Range("$A$1:$XX$" & LastRow).AutoFilter Field:=2, Criteria1:=Month_Num
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
        With Summary_Workbook.Sheets("Phase-2_Followup-A1")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            LastRow_Summary = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
        'Change Month number
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range("B2:B" & LastRow).Value = Month(Date)
        
        'Copy to Summary sheet
        If IsEmpty(.Range("A2")) Then
            GoTo Step1
        Else
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Summary_Workbook.Sheets("Phase-2_Followup-A1").Range("A" & LastRow_Summary + 1)
        End If
    End With
    
Step1:
    Followup_Phase_A1.SaveAs Filename:=ThisWorkbook.Path & "\FollowUp_A1_Output_file.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Followup_Phase_A1.Close
    
    'Phase-2_Followup-A2
    With Followup_Phase_A2.ActiveSheet
        .Activate
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
            
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        'Complete & Return Organizer Date Complete
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E4").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
            
        'Tax Return Status
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E3").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="<>Information due from Employer"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E5").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        'Keep data which is only existing in Phase1
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B1").Value = "ID's from GOS"
        .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]Phase1'!C1:C2,1,FALSE),1)"
        .Activate
        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
        
        'Exclude current month data
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B1").Value = "Month"
        .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]Phase1'!C1:C9,9,FALSE),1)"
        .Activate
        .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
        
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
        
        .Columns("B:C").Copy
        .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=3, Criteria1:="1"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0

            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
        Month_Num = Month(Date)
        .Range("$A$1:$XX$" & LastRow).AutoFilter Field:=2, Criteria1:=Month_Num
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        With Summary_Workbook.Sheets("Phase-2_Followup-A2")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            LastRow_Summary = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If

        'Change Month number
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range("B2:B" & LastRow).Value = Month(Date)
        
        'Copy to Summary sheet
        If IsEmpty(.Range("A2")) Then
            GoTo Step2
        Else
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Summary_Workbook.Sheets("Phase-2_Followup-A2").Range("A" & LastRow_Summary + 1)
        End If
    End With
    
Step2:
    Followup_Phase_A2.SaveAs Filename:=ThisWorkbook.Path & "\FollowUp_A2_Output_file.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Followup_Phase_A2.Close
    
    'ADC
    With ADC_Wrkbk.ActiveSheet
        .Activate
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        'Complete & Return Organizer Date Complete
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E4").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        'Tax Return Status
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E3").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="Information due from Employer" ', Operator:=xlAnd, Criteria2:="Information due from Assignee"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
         
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E3").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="Information due from Assignee" ', Operator:=xlAnd, Criteria2:="Information due from Assignee"
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E5").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("B1").Value = "ID'S"
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("B1").Value = "Month"
        .Range("B2:B" & LastRow).Value = Month(Date)
        
        .Columns("B:B").Copy
        .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        With Summary_Workbook.Sheets("Phase-2_ADC")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            
            LastRow_Summary = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        'Copy to Summary sheet
        If IsEmpty(.Range("A2")) Then
            GoTo Step3
        Else
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Summary_Workbook.Sheets("Phase-2_ADC").Range("A" & LastRow_Summary + 1)
        End If
    End With
    
Step3:
    ADC_Wrkbk.SaveAs Filename:=ThisWorkbook.Path & "\ADC_Output_file.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ADC_Wrkbk.Close
    
    k = 1
    TR_Wrkbook.Close
    With Summary_Workbook
        Phase1_Wrksht.Delete
        FollowUp_Wrksht.Delete
        .Save
        .Close
    End With
    MsgBox "Output files are Generated"
    
ExitHere:
    On Error Resume Next
    If k = 0 Then
        GOS_Wrkbook.Close
        TR_Wrkbook.Close
        Summary_Workbook.Close
    End If
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlCalculationAutomatic
    End With
    On Error GoTo 0
End Sub
Sub Phase_2_Mail()

    Dim Mapping_WrkBook As Workbook
    Dim Path As String: Dim Phase_Type As String: Dim Phase_Sub_Type As String
    Dim LastRow As Long: Dim Rng As Range: 'Dim k
    
    Dim Response As Integer: Dim WrkSht As Worksheet
    Dim Summary_Workbook As Workbook: Dim Followup_A1_Wrkbk As Workbook: Dim Followup_A2_Wrkbk As Workbook: Dim ADC_Wrkbk As Workbook
    Dim FP1 As Integer, FP2 As Integer, AD As Integer
    'Sign related
    Dim sPath As String: Dim Sign As String: Dim signImageFolderName As String
    Dim completeFolderPath As String: Dim StrSignature As String:
    
    Response = MsgBox("Have you checked sample mails?, if agree click Yes else click No", vbYesNo, "Confirmation")
    If Response = vbNo Then
        MsgBox "Since response is No: Exiting Macro"
        GoTo ExitHere
    End If
    
    Response = MsgBox("You have clicked Phase 2 button, if agree click Yes else click No", vbYesNo, "Confirmation")
    If Response = vbNo Then
        MsgBox "Since response is No: Exiting Macro"
        GoTo ExitHere
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Calculation = xlCalculationManual
    End With
    
    'Open Followup-A1 Output file
    MsgBox "Please choose Followup-A1 Output file"
    Call File_Picker_Fun(Path, "Please choose Followup-A1 Output file")
    If Path <> "" Then
        Set Followup_A1_Wrkbk = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Followup-A1 Output file :Exiting Macro"
        GoTo ExitHere
    End If

    If Followup_A1_Wrkbk.ActiveSheet.Range("A2") = "" Then
        MsgBox "Followup-A1 Output file is Empty: No Emails from Followup-A1 Output file, Click Ok to continue"
        Followup_A1_Wrkbk.Close
        FP1 = 1
    End If
    
    'Open Followup-A2 Output file
    MsgBox "Please choose Followup-A2 Output file"
    Call File_Picker_Fun(Path, "Please choose Followup-A2 Output file")
    If Path <> "" Then
        Set Followup_A2_Wrkbk = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Followup-A2 Output file :Exiting Macro"
        GoTo ExitHere
    End If

    If Followup_A2_Wrkbk.ActiveSheet.Range("A2") = "" Then
        MsgBox "Followup-A2 Output file is Empty: No Emails from Followup-A2 Output file, Click Ok to continue"
        Followup_A2_Wrkbk.Close
        FP2 = 1
    End If
        
    'Open ADC Output file
    MsgBox "Please choose ADC Output file"
    Call File_Picker_Fun(Path, "Please choose ADC Output file")
    If Path <> "" Then
        Set ADC_Wrkbk = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose ADC Output file :Exiting Macro"
        GoTo ExitHere
    End If

    If ADC_Wrkbk.ActiveSheet.Range("A2") = "" Then
        MsgBox "ADC Output file is Empty: No Emails from ADC Output file, Click Ok to continue"
        ADC_Wrkbk.Close
        AD = 1
    End If
    
    'Open Summary report file
    MsgBox "Please choose Summary file"
    Call File_Picker_Fun(Path, "Please choose Summary file")
    If Path <> "" Then
        Set Summary_Workbook = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Summary file :Exiting Macro"
        GoTo ExitHere
    End If
    
    'Open Mapping file
    Path = ""
    MsgBox "Please choose Mapping file"
    Call File_Picker_Fun(Path, "Please choose Mapping file")
    If Path <> "" Then
        Set Mapping_WrkBook = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Mapping file :Exiting Macro"
        GoTo ExitHere
    End If

    'Add Signature
    Sign = "jj"
    sPath = Environ("appdata") & "\Microsoft\Signatures\" & Sign & ".htm"
    signImageFolderName = Sign & "_files"
    completeFolderPath = Environ("appdata") & "\Microsoft\Signatures\" & signImageFolderName

    StrSignature = GetSignature(sPath)
    StrSignature = VBA.Replace(StrSignature, """" & signImageFolderName, """" & completeFolderPath)

   '******FollowUp-A1_Wrkbk file
   If FP1 = 0 Then
        Phase_Type = "Template-Phase2"
        Phase_Sub_Type = "Phase-2_Followup-A1"
        With Followup_A1_Wrkbk.Sheets(Phase_Sub_Type)
        
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
                
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        End With
        Call SendMail_Phase_2(Followup_A1_Wrkbk, Mapping_WrkBook, Summary_Workbook, Rng, Phase_Type, Phase_Sub_Type, StrSignature)
        Set Rng = Nothing
        Followup_A1_Wrkbk.Save
        Followup_A1_Wrkbk.Close
    End If
    
    
   '******FollowUp-A2_Wrkbk file
   If FP2 = 0 Then
        Phase_Type = "Template-Phase2"
        Phase_Sub_Type = "Phase-2_Followup-A2"
        With Followup_A2_Wrkbk.Sheets(Phase_Sub_Type)
        
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
                
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        End With
        Call SendMail_Phase_2(Followup_A2_Wrkbk, Mapping_WrkBook, Summary_Workbook, Rng, Phase_Type, Phase_Sub_Type, StrSignature)
        Set Rng = Nothing
        Followup_A2_Wrkbk.Save
        Followup_A2_Wrkbk.Close
    End If
    
    
   '******ADC_Wrkbk file
   If AD = 0 Then
        Phase_Type = "Template-Phase2"
        Phase_Sub_Type = "Phase-2_ADC"
        With ADC_Wrkbk.Sheets(Phase_Sub_Type)
        
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
                
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        End With
        Call SendMail_Phase_2(ADC_Wrkbk, Mapping_WrkBook, Summary_Workbook, Rng, Phase_Type, Phase_Sub_Type, StrSignature)
        Set Rng = Nothing
        ADC_Wrkbk.Save
        ADC_Wrkbk.Close
    End If
    
    
    Summary_Workbook.Save
    Summary_Workbook.Close
    Mapping_WrkBook.Close
    
    MsgBox "Email has been sent!!", vbInformation
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlCalculationAutomatic
    End With
    On Error GoTo 0
    'Application.Quit
    
ExitHere:
    On Error Resume Next
        Followup_A1_Wrkbk.Close
        Followup_A2_Wrkbk.Close
        ADC_Wrkbk.Close
        Summary_Workbook.Close
        Mapping_WrkBook.Close
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlCalculationAutomatic
    End With
    On Error GoTo 0
End Sub

'The Functionality of this Fun is to Select a file object
Function File_Picker_Fun(StrFile As String, Title_Str As String)

Dim FD As Office.FileDialog
Set FD = Application.FileDialog(msoFileDialogFilePicker)

    With FD
        .Filters.Clear
        .Filters.Add "Excel Files", "*.*", 1
        .Title = Title_Str
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show = True Then
            StrFile = .SelectedItems(1)
        End If
    End With

End Function

'Signature
Function GetSignature(fPath As String) As String
    Dim fso As Object
    Dim TSet As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set TSet = fso.GetFile(fPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.ReadAll
    TSet.Close
End Function




