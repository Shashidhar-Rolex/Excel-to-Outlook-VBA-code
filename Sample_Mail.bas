Attribute VB_Name = "Sample_Mail"
Option Explicit

Sub Check_Mails_Phase1()

    Dim Rng As Range: Dim Cell As Range
    Dim OutApp As Object: Dim OutMail As Object
    Dim Mail_Body As String: Dim Subj As String

    'Sign related
    Dim sPath As String: Dim Sign As String: Dim signImageFolderName As String
    Dim completeFolderPath As String: Dim StrSignature As String:
    
    'Add Signature
    Sign = "jj"
    sPath = Environ("appdata") & "\Microsoft\Signatures\" & Sign & ".htm"
    signImageFolderName = Sign & "_files"
    completeFolderPath = Environ("appdata") & "\Microsoft\Signatures\" & signImageFolderName

    StrSignature = GetSignature(sPath)
    StrSignature = VBA.Replace(StrSignature, """" & signImageFolderName, """" & completeFolderPath)
    
    
    Set OutApp = CreateObject("Outlook.Application")
    
    Set Rng = ThisWorkbook.Sheets("Template-Phase1").Range("B3:E3")
    
    For Each Cell In Rng
            Mail_Body = Cell.Offset(1, 0).Value
            Subj = Cell.Value
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .SentOnBehalfOfName = "" 'SendBehalf
            .To = Application.UserName 'ToEmail
            .CC = ""
            .BCC = "" 'BccEmail
            .Subject = Subj
            .HTMLBody = Mail_Body & StrSignature
            .Display
            Application.Wait (Now + TimeValue("0:00:03"))
            DoEvents
            'SendKeys "^{ENTER}"
            SendKeys "%(s)"
        End With
    Next Cell
    
End Sub
Sub Check_Mails_Phase2()

    Dim Rng As Range: Dim Cell As Range
    Dim OutApp As Object: Dim OutMail As Object
    Dim Mail_Body As String: Dim Subj As String

    'Sign related
    Dim sPath As String: Dim Sign As String: Dim signImageFolderName As String
    Dim completeFolderPath As String: Dim StrSignature As String:
    
    'Add Signature
    Sign = "jj"
    sPath = Environ("appdata") & "\Microsoft\Signatures\" & Sign & ".htm"
    signImageFolderName = Sign & "_files"
    completeFolderPath = Environ("appdata") & "\Microsoft\Signatures\" & signImageFolderName

    StrSignature = GetSignature(sPath)
    StrSignature = VBA.Replace(StrSignature, """" & signImageFolderName, """" & completeFolderPath)
    
    
    Set OutApp = CreateObject("Outlook.Application")
    
    Set Rng = ThisWorkbook.Sheets("Template-Phase2").Range("B3:E3")
    
    For Each Cell In Rng
            Mail_Body = Cell.Offset(1, 0).Value
            Subj = Cell.Value
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .SentOnBehalfOfName = "" 'SendBehalf
            .To = Application.UserName 'ToEmail
            .CC = ""
            .BCC = "" 'BccEmail
            .Subject = Subj
            .HTMLBody = Mail_Body & StrSignature
            .Display
            Application.Wait (Now + TimeValue("0:00:03"))
            DoEvents
            'SendKeys "^{ENTER}"
            SendKeys "%(s)"
        End With
    Next Cell
    
End Sub

'Signature
Function GetSignature(fPath As String) As String
    Dim fso As Object
    Dim TSet As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set TSet = fso.GetFile(fPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.ReadAll
    TSet.Close
End Function

