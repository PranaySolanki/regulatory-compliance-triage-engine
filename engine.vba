Sub RunRegulatoryFetch()
    Dim OutlookApp As Outlook.Application
    Dim OutlookNamespace As Outlook.Namespace
    Dim Inbox As Outlook.MAPIFolder
    Dim Mail As Object
    Dim Sheet As Worksheet
    Dim ArchiveSheet As Worksheet
    Dim EmailBody As String
    Dim TargetDays As Variant
    Dim CutoffDate As Date
    Dim LastRow As Long
    Dim MatchCount As Long
    Dim DuplicateCount As Long
    
    Set Sheet = ThisWorkbook.Sheets("Dashboard")
    Set ArchiveSheet = ThisWorkbook.Sheets("Compliance_Archive")
    
    ' 1. READ TIMELINE FILTER
    TargetDays = Sheet.Range("B4").Value
    If Not IsNumeric(TargetDays) Or TargetDays <= 0 Then
        MsgBox "Please enter a valid positive number of days in cell B4 before running.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    CutoffDate = Date - CInt(TargetDays)
    
    ' 2. SAFE CLEANUP: Erase previous dashboard view from Row 7 downwards
    LastRow = Sheet.Cells(Sheet.Rows.Count, "A").End(xlUp).Row
    If LastRow >= 7 Then
        Sheet.Range("A7:G" & LastRow).ClearContents
        Sheet.Range("A7:G" & LastRow).Interior.Color = xlNone
        Sheet.Range("A7:G" & LastRow).Font.Bold = False
        Sheet.Rows("7:" & LastRow).RowHeight = Sheet.StandardHeight
    End If
    
    ' 3. INITIALIZE OUTLOOK
    On Error GoTo OutlookError
    Set OutlookApp = New Outlook.Application
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set Inbox = OutlookNamespace.GetDefaultFolder(olFolderInbox)
    On Error GoTo 0
    
    MatchCount = 0
    DuplicateCount = 0
    
    Dim FolderItems As Outlook.Items
    Set FolderItems = Inbox.Items
    FolderItems.Sort "[ReceivedTime]", False ' Oldest to Newest loop
    
    ' 4. PROCESSING LOOP
    For Each Mail In FolderItems
        If TypeOf Mail Is MailItem Then
            
            If Mail.ReceivedTime >= CutoffDate Then
                
                If InStr(1, Mail.Subject, "RBI", vbTextCompare) > 0 Or _
                   InStr(1, Mail.Subject, "CERT", vbTextCompare) > 0 Then
                    
                    EmailBody = Mail.Body
                    MatchCount = MatchCount + 1
                    
                    ' --- DUPLICATE CHECK ENGINE ---
                    Dim IsDuplicate As Boolean
                    IsDuplicate = False
                    
                    Dim ArchiveLastRow As Long
                    ArchiveLastRow = ArchiveSheet.Cells(ArchiveSheet.Rows.Count, "A").End(xlUp).Row
                    
                    If ArchiveLastRow >= 7 Then
                        Dim CheckRow As Long
                        For CheckRow = 7 To ArchiveLastRow
                            If ArchiveSheet.Cells(CheckRow, 5).Value = Mail.Subject Then
                                IsDuplicate = True
                                DuplicateCount = DuplicateCount + 1
                                Exit For
                            End If
                        Next CheckRow
                    End If
                    
                    ' --- EVALUATE ATTACHMENTS SECURELY ---
                    Dim HasAttachment As String
                    If Mail.Attachments.Count > 0 Then
                        HasAttachment = "Yes"
                    Else
                        HasAttachment = "No"
                    End If
                    
                    ' --- WRITE TO ACTIVE DASHBOARD VIEW (NOW INSERTS NEWEST AT TOP) ---
                    Sheet.Rows(7).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                    With Sheet.Range("A7:G7")
                        .Interior.Color = xlNone
                        .Font.Bold = False
                        .Font.Color = RGB(0, 0, 0)
                        .VerticalAlignment = xlTop
                        .WrapText = True
                    End With
                    
                    Sheet.Cells(7, 1).Value = Mail.ReceivedTime
                    Sheet.Cells(7, 2).Value = Mail.SenderName & " (" & Mail.SenderEmailAddress & ")"
                    Sheet.Cells(7, 3).Value = Mail.To
                    Sheet.Cells(7, 4).Value = HasAttachment
                    Sheet.Cells(7, 5).Value = Mail.Subject
                    Sheet.Cells(7, 6).Value = ExtractSectionBlock(EmailBody, "Overview")
                    Sheet.Cells(7, 7).Value = ExtractSectionBlock(EmailBody, "Recommendation")
                    Sheet.Rows(7).AutoFit
                    
                    ' --- WRITE TO ARCHIVE SHEET (NEWEST AT TOP + NO DUPLICATES) ---
                    If Not IsDuplicate Then
                        ArchiveSheet.Rows(7).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                        
                        With ArchiveSheet.Range("A7:G7")
                            .Interior.Color = xlNone
                            .Font.Bold = False
                            .Font.Color = RGB(0, 0, 0)
                            .VerticalAlignment = xlTop
                            .WrapText = True
                        End With
                        
                        ArchiveSheet.Cells(7, 1).Value = Mail.ReceivedTime
                        ArchiveSheet.Cells(7, 2).Value = Sheet.Cells(7, 2).Value
                        ArchiveSheet.Cells(7, 3).Value = Sheet.Cells(7, 3).Value
                        ArchiveSheet.Cells(7, 4).Value = Sheet.Cells(7, 4).Value
                        ArchiveSheet.Cells(7, 5).Value = Mail.Subject
                        ArchiveSheet.Cells(7, 6).Value = Sheet.Cells(7, 6).Value
                        ArchiveSheet.Cells(7, 7).Value = Sheet.Cells(7, 7).Value
                        ArchiveSheet.Rows(7).AutoFit
                    End If
                    
                End If
            End If
        End If
    Next Mail
    
    ' Format Dimensions
    Sheet.Columns("A:D").AutoFit
    Sheet.Columns("E").ColumnWidth = 30
    Sheet.Columns("F:G").ColumnWidth = 55
    
    ArchiveSheet.Columns("A:D").AutoFit
    ArchiveSheet.Columns("E").ColumnWidth = 30
    ArchiveSheet.Columns("F:G").ColumnWidth = 55
    
    MsgBox "Secure Scan Report:" & vbCrLf & _
           "---------------------------" & vbCrLf & _
           "• Total Advisories Scanned: " & MatchCount & vbCrLf & _
           "• New Records Archived at Top: " & (MatchCount - DuplicateCount) & vbCrLf & _
           "• Duplicate Records Skipped Safely: " & DuplicateCount, vbInformation, "Sync Status"
    Exit Sub

OutlookError:
    MsgBox "Could not connect to Outlook.", vbCritical, "Connection Failed"
End Sub

Function ExtractSectionBlock(BodyText As String, SectionType As String) As String
    Dim StartPos As Long, EndPos As Long
    Dim CleanBody As String
    Dim Headers As Variant, SignOffs As Variant
    Dim h As Variant, s As Variant
    
    CleanBody = BodyText
    ExtractSectionBlock = ""
    
    If SectionType = "Overview" Then
        Headers = Array("Overview:", "Description:", "Background:", "Summary:")
        SignOffs = Array("Recommendation", "Solution:", "Action Required:", "Action Points:", "Fix:", "Regards", "Sincerely", "Directions")
    ElseIf SectionType = "Recommendation" Then
        Headers = Array("Recommendation:", "Solution:", "Action Points:", "Directions:", "Advisory:")
        SignOffs = Array("Regards", "Sincerely", "Disclaimer:", "Yours faithfully", "Note:", "Chief General Manager")
    End If
    
    For Each h In Headers
        StartPos = InStr(1, CleanBody, h, vbTextCompare)
        If StartPos > 0 Then
            StartPos = StartPos + Len(h)
            Exit For
        End If
    Next h
    
    If StartPos > 0 Then
        EndPos = Len(CleanBody)
        For Each s In SignOffs
            Dim TempEnd As Long
            TempEnd = InStr(StartPos, CleanBody, s, vbTextCompare)
            If TempEnd > 0 And TempEnd < EndPos Then
                EndPos = TempEnd
            End If
        Next s
        ExtractSectionBlock = Trim(Mid(CleanBody, StartPos, EndPos - StartPos))
    End If
    
    If ExtractSectionBlock = "" Then
        If SectionType = "Overview" Then
            ExtractSectionBlock = "Header not identified. Raw text snippet: " & Left(BodyText, 250)
        Else
            ExtractSectionBlock = "No structured recommendation block found. Please check raw email body."
        End If
    End If
End Function

