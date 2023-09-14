Attribute VB_Name = "Phase1_New"
Option Explicit
Sub Generate_Output_Phase1()

    Dim TR_Wrkbook As Workbook: Dim GOS_Wrkbook As Workbook: Dim Mapping_WrkBook As Workbook
    Dim Path As String: Dim LastRow As Long: Dim Rng As Range: Dim Cell As Range: Dim k
    Dim Fnd As Range: Dim Response As Integer: Dim WrkSht As Worksheet
    Dim Summary_Workbook As Workbook
    Dim LastCol As Long: Dim LastRow_Summary As Long
    Dim Temp: Dim Wrksht_Map As Worksheet: Dim Col_Num As Integer
    Dim Phase_A1_Wrkbk As Workbook: Dim Phase_A2_Wrkbk As Workbook
    
    Response = MsgBox("You have clicked Phase 1 button, if agree click Yes else click No", vbYesNo, "Confirmation")
    
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
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        'Remove repeatative data by comparing with Summary file
        For Each WrkSht In Summary_Workbook.Sheets
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            If WrkSht.Range("A2").Value <> "" Then
                .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                .Range("B1").Value = "ID's from GOS"
                .Range("B2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[" & Summary_Workbook.Name & "]" & WrkSht.Name & "'!C1:C2,1,FALSE),1)"
                .Activate
                If LastRow <> 2 Then
                    .Range("B2").AutoFill Destination:=Range("B2:B" & LastRow)
                    Application.Calculation = xlCalculationAutomatic
                    Application.Calculation = xlCalculationManual
                    .Range("$A$1:$Y$" & LastRow).AutoFilter Field:=2, Criteria1:="<>1"
                    On Error Resume Next
                    .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                    On Error GoTo 0
                    .Columns("B:B").Delete Shift:=xlToLeft
                End If
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
        .Range("$A$1:$Y$" & LastRow).AutoFilter Field:=2, Criteria1:=1
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        
        .Columns("B:B").Copy
        .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
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
    'Add two workbooks, one for Phase_A1 & another for Phase_A2
    Set Phase_A1_Wrkbk = Workbooks.Add
    Phase_A1_Wrkbk.ActiveSheet.Name = "Phase-A1"
    
    Set Phase_A2_Wrkbk = Workbooks.Add
    Phase_A2_Wrkbk.ActiveSheet.Name = "Phase-A2"
    
    'Copy Data from TR report to Individual File
    With TR_Wrkbook.ActiveSheet
        .Activate
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=Phase_A1_Wrkbk.Sheets(1).Range("A1")
        .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Copy Destination:=Phase_A2_Wrkbk.Sheets(1).Range("A1")
    End With
    
    'Phase_A1
    With Phase_A1_Wrkbk.ActiveSheet
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
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
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        .Range("CX1").Value = "Status"
        .Range("CY1").Value = "Time"
        
        With Summary_Workbook.Sheets("Phase-A1")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        
            LastRow_Summary = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
        .Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        LastCol = LastCol - 2
        
        'Copy to Summary sheet
        If IsEmpty(.Range("A2")) Then
            GoTo Step1
        Else
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Summary_Workbook.Sheets("Phase-A1").Range("A" & LastRow_Summary + 1)
        End If

        
        With Summary_Workbook.Sheets("Phase-A1")
            .Activate
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
            Temp = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range("H" & LastRow_Summary + 1).FormulaR1C1 = "=Today()"
            .Range("I" & LastRow_Summary + 1).FormulaR1C1 = "=Month(Today())"
            .Range("H" & LastRow_Summary + 1).AutoFill Destination:=Range("H" & LastRow_Summary + 1 & ":H" & Temp)
            .Range("I" & LastRow_Summary + 1).AutoFill Destination:=Range("I" & LastRow_Summary + 1 & ":I" & Temp)
            
            .Columns("H:H").Copy
            .Range("H1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Columns("I:I").Copy
            .Range("I1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        
        .Columns("H:I").Delete Shift:=xlToLeft
        
    End With
    
Step1:
    Phase_A1_Wrkbk.SaveAs Filename:=ThisWorkbook.Path & "\Phase_A1_Output_file.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Phase_A1_Wrkbk.Close
    
    'Phase_A2
    With Phase_A2_Wrkbk.ActiveSheet
        .Activate
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
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
        
        If .FilterMode = True Then .Cells.AutoFilter
        
        Col_Num = WorksheetFunction.Match(Wrksht_Map.Range("E5").Value, .Range("A1:ZZ1"), 0)
        .Range("$A$1:$ZZ$" & LastRow).AutoFilter Field:=Col_Num, Criteria1:="="
        
        On Error Resume Next
        .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If

        Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        .Range("CX1").Value = "Status"
        .Range("CY1").Value = "Time"
        
        With Summary_Workbook.Sheets("Phase-A2")
            If .FilterMode = True Then .Cells.AutoFilter
            LastRow_Summary = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
        .Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        LastCol = LastCol - 2
        
        'Copy to Summary sheet
        If IsEmpty(.Range("A2")) Then
            GoTo Step2
        Else
            .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Copy Destination:=Summary_Workbook.Sheets("Phase-A2").Range("A" & LastRow_Summary + 1)
        End If
        
        With Summary_Workbook.Sheets("Phase-A2")
            If .FilterMode = True Then
                .Cells.AutoFilter
                .Cells.AutoFilter
            End If

            Temp = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range("H" & LastRow_Summary + 1).FormulaR1C1 = "=Today()"
            .Range("I" & LastRow_Summary + 1).FormulaR1C1 = "=Month(Today())"
            .Activate
            .Range("H" & LastRow_Summary + 1).AutoFill Destination:=Range("H" & LastRow_Summary + 1 & ":H" & Temp)
            .Range("I" & LastRow_Summary + 1).AutoFill Destination:=Range("I" & LastRow_Summary + 1 & ":I" & Temp)
            
            .Columns("H:H").Copy
            .Range("H1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Columns("I:I").Copy
            .Range("I1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        
        .Columns("H:I").Delete Shift:=xlToLeft
        
    End With
    
Step2:
    Phase_A2_Wrkbk.SaveAs Filename:=ThisWorkbook.Path & "\Phase_A2_Output_file.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Phase_A2_Wrkbk.Close
    
    'TR_Wrkbook.Close
    Summary_Workbook.Save
    Summary_Workbook.Close
    TR_Wrkbook.Close
    MsgBox "Output is Generated!!!"
    k = 1

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
 
Sub Phase_1_Mail()

    Dim Path As String: Dim Phase_Type As String: Dim Phase_Sub_Type As String
    Dim LastRow As Long: Dim Rng As Range: 'Dim k
    
    Dim Response As Integer: Dim WrkSht As Worksheet: Dim Mapping_WrkBook As Workbook
    Dim Summary_Workbook As Workbook: Dim Phase_A1_Wrkbk As Workbook: Dim Phase_A2_Wrkbk As Workbook '
    Dim PA1 As Integer: Dim PA2 As Integer
    
    'Sign related
    Dim sPath As String: Dim Sign As String: Dim signImageFolderName As String
    Dim completeFolderPath As String: Dim StrSignature As String:


    Response = MsgBox("Have you checked sample mails?, if agree click Yes else click No", vbYesNo, "Confirmation")
    If Response = vbNo Then
        MsgBox "Since response is No: Exiting Macro"
        GoTo ExitHere
    End If
    
    Response = MsgBox("You have clicked Phase 1 button, if agree click Yes else click No", vbYesNo, "Confirmation")
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
    
    'Open Phase-A1 Output file
    MsgBox "Please choose Phase-A1 Output file"
    Call File_Picker_Fun(Path, "Please choose Phase-A1 Output file")
    If Path <> "" Then
        Set Phase_A1_Wrkbk = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Phase-A1 output file :Exiting Macro"
        GoTo ExitHere
    End If

    If Phase_A1_Wrkbk.Sheets("Phase-A1").Range("A2") = "" Then
        MsgBox "Phase-A1 output file is Empty: No Emails from Phase1-A1 Output file, Click Ok to continue"
        Phase_A1_Wrkbk.Close
        PA1 = 1
    End If
    
    'Open Phase-A2 Output file
    MsgBox "Please choose Phase-A2 Output file"
    Call File_Picker_Fun(Path, "Please choose Phase-A2 Output file")
    If Path <> "" Then
        Set Phase_A2_Wrkbk = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Phase-A2 output file :Exiting Macro"
        GoTo ExitHere
    End If

    If Phase_A2_Wrkbk.Sheets("Phase-A2").Range("A2") = "" Then
        MsgBox "Phase-A2 output file is Empty: No Emails from Phase1-A2 Output file, Click Ok to continue"
        Phase_A2_Wrkbk.Close
        PA2 = 1
    End If
        
    'Open Summary report file
    MsgBox "Please choose Summary file"
    Call File_Picker_Fun(Path, "Please choose Summary file")
    If Path <> "" Then
        Set Summary_Workbook = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose Summary file :Exiting Macro"
        Phase_A1_Wrkbk.Close
        Phase_A2_Wrkbk.Close
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
        Phase_A1_Wrkbk.Close
        Phase_A2_Wrkbk.Close
        Summary_Workbook.Close
        GoTo ExitHere
    End If

    'Add Signature
    Sign = "jj"
    sPath = Environ("appdata") & "\Microsoft\Signatures\" & Sign & ".htm"
    signImageFolderName = Sign & "_files"
    completeFolderPath = Environ("appdata") & "\Microsoft\Signatures\" & signImageFolderName

    StrSignature = GetSignature(sPath)
    StrSignature = VBA.Replace(StrSignature, """" & signImageFolderName, """" & completeFolderPath)
    
    
   '******Phase-A1_Wrkbk file
   If PA1 = 0 Then
        Phase_Type = "Template-Phase1"
        Phase_Sub_Type = "Phase-A1"
        With Phase_A1_Wrkbk.Sheets(Phase_Sub_Type)
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Set Rng = .Range("A2:A" & LastRow) '.SpecialCells(xlCellTypeVisible)
        End With
        Call SendMail_Phase_1(Phase_A1_Wrkbk, Mapping_WrkBook, Summary_Workbook, Rng, Phase_Type, Phase_Sub_Type, StrSignature)
        Set Rng = Nothing
        Phase_A1_Wrkbk.Save
        Phase_A1_Wrkbk.Close
    End If
    
   '******Phase-A2_Wrkbk file
   If PA2 = 0 Then
        Phase_Type = "Template-Phase1"
        Phase_Sub_Type = "Phase-A2"
        With Phase_A2_Wrkbk.Sheets(Phase_Sub_Type)
                If .FilterMode = True Then
                    .Cells.AutoFilter
                    .Cells.AutoFilter
                End If
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Set Rng = .Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible)
        End With
        Call SendMail_Phase_1(Phase_A2_Wrkbk, Mapping_WrkBook, Summary_Workbook, Rng, Phase_Type, Phase_Sub_Type, StrSignature)
        Set Rng = Nothing
        Phase_A2_Wrkbk.Save
        Phase_A2_Wrkbk.Close
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
    
        Phase_A1_Wrkbk.Close
        Phase_A2_Wrkbk.Close
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



