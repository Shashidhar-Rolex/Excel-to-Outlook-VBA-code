Attribute VB_Name = "Final_Mail"
Option Explicit

Sub SendMail_Phase_1(Output_Wrkbk As Workbook, Mapping_WrkBook As Workbook, Summary_Workbook As Workbook, Rng As Range, Phase_Type As String, Phase_Sub_Type As String, StrSignature As String)
    
    Dim OutApp As Object: Dim OutMail As Object
    
    Dim TempNum As Integer: Dim Subj As String: Dim SendBehalf As String: Dim BccEmail As String
    Dim Mail_Body As String: Dim Rng_Del As Range: Dim Cell As Range: Dim Fnd As Range: Dim Cell_Del As Range
    Dim LastRow_Del As Long: Dim Month_Index As Integer: Dim ToEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    
    'With Mapping sheet
    With Mapping_WrkBook.ActiveSheet
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
    End With
        'Output_Wrkbk.ActiveSheet.Range("CY1") = "Status"
        'Output_Wrkbk.ActiveSheet.Range("CZ1") = "Time"

        For Each Cell In Rng
        Set Fnd = Nothing
            Set Fnd = Mapping_WrkBook.ActiveSheet.Range("A:A").Find(what:=Trim(Cell.Offset(0, 5).Value), LookAt:=xlWhole)
            ToEmail = WorksheetFunction.Clean(Trim(Cell.Offset(0, 4).Value))
            Month_Index = Month(Date) 'WorksheetFunction.Clean(Trim(Cell.Offset(0, 4).Value))
            
            If Not Fnd Is Nothing Then
                 If Fnd.Offset(0, 1).Value = "" Then
                    If Fnd.Offset(0, 2).Value = "Yes" Then
                        SendBehalf = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 4).Value))
                        BccEmail = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 5).Value))
                        
                        With ThisWorkbook.Sheets(Phase_Type)
                            If Fnd.Offset(0, 3) = "Yes" Then
                                TempNum = 2
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            Else
                                TempNum = 4
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            End If
                        End With
                        
                    Else
                        SendBehalf = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 4).Value))
                        BccEmail = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 5).Value))
                    
                        With ThisWorkbook.Sheets(Phase_Type)
                            If Fnd.Offset(0, 3) = "Yes" Then
                                TempNum = 3
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            Else
                                TempNum = 5
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            End If
                        End With
                    End If
                    
                    Mail_Body = Replace(Mail_Body, "{FirstName}", WorksheetFunction.Proper(Trim(Cell.Offset(0, 2).Value)))

                    Set OutMail = OutApp.CreateItem(0)
                    'On Error Resume Next comment
                    With OutMail

                        .SentOnBehalfOfName = SendBehalf
                        .To = ToEmail
                        .CC = ""
                        .BCC = BccEmail
                        .Subject = Subj
                        .HTMLBody = Mail_Body & StrSignature
                        .Display
                        Application.Wait (Now + TimeValue("0:00:03"))
                        DoEvents
                        SendKeys "%(s)"

                    End With
                    'On Error GoTo 0
                 Output_Wrkbk.ActiveSheet.Range("CX" & Cell.Row) = ThisWorkbook.Sheets(Phase_Type).Cells(1, TempNum).Value
                 Output_Wrkbk.ActiveSheet.Range("CY" & Cell.Row) = Now
                 'Z = Z + 1
                 'If Z = 3 Then Stop 'GoTo ExitHere
            Else
                Set Fnd = Nothing
                Output_Wrkbk.ActiveSheet.Range("CX" & Cell.Row) = "No Email"
                
                With Summary_Workbook.Sheets(Phase_Sub_Type)
                
                    If .FilterMode = True Then
                        .Cells.AutoFilter
                        .Cells.AutoFilter
                    End If
                    
                    LastRow_Del = .Cells(.Rows.Count, "A").End(xlUp).Row
                    
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=1, Criteria1:=Cell.Value
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=9, Criteria1:=Month_Index
                    
                    On Error Resume Next
                        Set Rng_Del = .Range("A2:A" & LastRow_Del).SpecialCells(xlCellTypeVisible)
                        If Err.Number = 0 Then
                            For Each Cell_Del In Rng_Del
                                .Rows(Cell_Del.Row).EntireRow.Delete
                            Next Cell_Del
                        End If
                    On Error GoTo 0
                    
                    If .FilterMode = True Then
                        .Cells.AutoFilter
                        .Cells.AutoFilter
                    End If
                    
                    Set Rng_Del = Nothing
                End With
                
                    'Set Fnd = Summary_Workbook.Sheets("Phase-1").Range("A:A").Find(what:=Trim(Cell.Value))
                    'Fnd.EntireRow.Delete
            End If
        Else
            Set Fnd = Nothing
            Output_Wrkbk.ActiveSheet.Range("CX" & Cell.Row) = "No Email"
            
                With Summary_Workbook.Sheets(Phase_Sub_Type)
                    If .FilterMode = True Then
                        .Cells.AutoFilter
                        .Cells.AutoFilter
                    End If
                    LastRow_Del = .Cells(.Rows.Count, "A").End(xlUp).Row
                    
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=1, Criteria1:=Cell.Value
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=9, Criteria1:=Month_Index
                    
                    On Error Resume Next
                        Set Rng_Del = .Range("A2:A" & LastRow_Del).SpecialCells(xlCellTypeVisible)
                        If Err.Number = 0 Then
                            For Each Cell_Del In Rng_Del
                                .Rows(Cell_Del.Row).EntireRow.Delete
                            Next Cell_Del
                        End If
                    On Error GoTo 0

                    
                    If .FilterMode = True Then
                        .Cells.AutoFilter
                        .Cells.AutoFilter
                    End If
                    
                    Set Rng_Del = Nothing
                End With
                
        End If
        
            Set Fnd = Nothing
            Set OutMail = Nothing
        Next Cell

End Sub

Sub SendMail_Phase_2(Output_Wrkbk As Workbook, Mapping_WrkBook As Workbook, Summary_Workbook As Workbook, Rng As Range, Phase_Type As String, Phase_Sub_Type As String, StrSignature As String)
    
    Dim OutApp As Object: Dim OutMail As Object
    
    Dim TempNum As Integer: Dim Subj As String: Dim SendBehalf As String: Dim BccEmail As String
    Dim Mail_Body As String: Dim Rng_Del As Range: Dim Cell As Range: Dim Fnd As Range: Dim Cell_Del As Range
    Dim LastRow_Del As Long: Dim Month_Index As Integer: Dim ToEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    
    'With Mapping sheet
    With Mapping_WrkBook.ActiveSheet
    
        If .FilterMode = True Then
            .Cells.AutoFilter
            .Cells.AutoFilter
        End If
    End With
        Output_Wrkbk.ActiveSheet.Range("CY1") = "Status"
        Output_Wrkbk.ActiveSheet.Range("CZ1") = "Time"

        For Each Cell In Rng
            Set Fnd = Nothing
            Set Fnd = Mapping_WrkBook.ActiveSheet.Range("A:A").Find(what:=Trim(Cell.Offset(0, 6).Value), LookAt:=xlWhole)
            ToEmail = WorksheetFunction.Clean(Trim(Cell.Offset(0, 5).Value))
            Month_Index = WorksheetFunction.Clean(Trim(Cell.Offset(0, 1).Value))
            If Not Fnd Is Nothing Then
                If Fnd.Offset(0, 1).Value = "" Then
                    If Phase_Sub_Type = "Phase-2_ADC" Then
                        SendBehalf = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 4).Value))
                        BccEmail = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 5).Value))

                        With ThisWorkbook.Sheets(Phase_Type)
                            If Fnd.Offset(0, 3) = "Yes" Then
                                TempNum = 2
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            Else
                                TempNum = 4
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            End If
                        End With
                    Else
                        SendBehalf = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 4).Value))
                        BccEmail = WorksheetFunction.Clean(Trim(Fnd.Offset(0, 5).Value))
                        With ThisWorkbook.Sheets(Phase_Type)
                            If Fnd.Offset(0, 3) = "Yes" Then
                                TempNum = 3
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            Else
                                TempNum = 5
                                Subj = .Cells(3, TempNum).Value
                                Mail_Body = .Cells(4, TempNum).Value
                            End If
                        End With
                    End If

                    Mail_Body = Replace(Mail_Body, "{FirstName}", WorksheetFunction.Proper(Trim(Cell.Offset(0, 3).Value)))

                    Set OutMail = OutApp.CreateItem(0)
                    'On Error Resume Next comment
                    With OutMail

                        .SentOnBehalfOfName = SendBehalf
                        .To = ToEmail
                        .CC = ""
                        .BCC = BccEmail
                        .Subject = Subj
                        .HTMLBody = Mail_Body & StrSignature
                        .Display
                        Application.Wait (Now + TimeValue("0:00:03"))
                        DoEvents
                        SendKeys "%(s)"
                        
                    End With
                    'On Error GoTo 0
                 Output_Wrkbk.ActiveSheet.Range("CX" & Cell.Row) = ThisWorkbook.Sheets(Phase_Type).Cells(1, TempNum).Value
                 Output_Wrkbk.ActiveSheet.Range("CY" & Cell.Row) = Now
                 
                Else
                    Set Fnd = Nothing
                    Output_Wrkbk.ActiveSheet.Range("CY" & Cell.Row) = "No Email"
                    With Summary_Workbook.Sheets(Phase_Sub_Type)
                        If .FilterMode = True Then
                            .Cells.AutoFilter
                            .Cells.AutoFilter
                        End If
                        LastRow_Del = .Cells(.Rows.Count, "A").End(xlUp).Row

                        .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=1, Criteria1:=Cell.Value
                        .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=2, Criteria1:=Month_Index

                        On Error Resume Next
                            Set Rng_Del = .Range("A2:A" & LastRow_Del).SpecialCells(xlCellTypeVisible)
                            If Err.Number = 0 Then
                                For Each Cell_Del In Rng_Del
                                    .Rows(Cell_Del.Row).EntireRow.Delete
                                Next Cell_Del
                            End If
                        On Error GoTo 0

                        If .FilterMode = True Then
                            .Cells.AutoFilter
                            .Cells.AutoFilter
                        End If
                        Set Rng_Del = Nothing
                    End With
                End If
            Else
                Set Fnd = Nothing
                Output_Wrkbk.ActiveSheet.Range("CY" & Cell.Row) = "No Email"

                With Summary_Workbook.Sheets(Phase_Sub_Type)

                        If .FilterMode = True Then
                            .Cells.AutoFilter
                            .Cells.AutoFilter
                        End If

                    LastRow_Del = .Cells(.Rows.Count, "A").End(xlUp).Row
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=1, Criteria1:=Cell.Value
                    .Range("$A$1:$Y$" & LastRow_Del).AutoFilter Field:=2, Criteria1:=Month_Index

                        On Error Resume Next
                            Set Rng_Del = .Range("A2:A" & LastRow_Del).SpecialCells(xlCellTypeVisible)
                            If Err.Number = 0 Then
                                For Each Cell_Del In Rng_Del
                                    .Rows(Cell_Del.Row).EntireRow.Delete
                                Next Cell_Del
                            End If
                        On Error GoTo 0

                        If .FilterMode = True Then
                            .Cells.AutoFilter
                            .Cells.AutoFilter
                        End If

                    Set Rng_Del = Nothing
                End With
            End If
            Set Fnd = Nothing
            Set OutMail = Nothing
        Next Cell

End Sub

