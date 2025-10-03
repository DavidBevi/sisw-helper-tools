Sub Calibration_UnifyHeaders_CopySerial()
    ' ••• DAVID BEVI's SHORTCUT •••
    ' (1) Unify all HEADERS & FOOTERS
    ' (2) Copy SERIAL-NUMBER from file-name to HEADER & FIRST-PAGE
    
    Dim doc As Document
    Set doc = ActiveDocument

    Dim i As Long
    Dim firstSection As Section
    Set firstSection = doc.Sections(1)
    
    ' MACRO OPTIONS: single_undo and pause_screen_updating
    Application.UndoRecord.StartCustomRecord ("Calibration_Header")
    Application.ScreenUpdating = False
    
    ' Step 1: If there's a section-break in last 2 chars remove it
    If InStr(ActiveDocument.Range(ActiveDocument.Content.End - 2, ActiveDocument.Content.End).Text, Chr(12)) > 0 Then
        ActiveDocument.Range(ActiveDocument.Content.End - 2, ActiveDocument.Content.End).Delete
    End If

    ' Step 2: Unify all headers and footers to the first section headers/footers
    For i = 2 To doc.Sections.Count
        ' Link headers and footers to the first
        doc.Sections(i).Headers(wdHeaderFooterPrimary).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterPrimary).LinkToPrevious = True
        doc.Sections(i).Headers(wdHeaderFooterFirstPage).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterFirstPage).LinkToPrevious = True
        doc.Sections(i).Headers(wdHeaderFooterEvenPages).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterEvenPages).LinkToPrevious = True
    Next i

    ' Refresh headers/footers links
    doc.Fields.Update

    ' Step 3: get header text (from first section, since all headers are now linked)
    Dim headerRange As Range
    Set headerRange = firstSection.Headers(wdHeaderFooterPrimary).Range
    Dim headerText As String
    headerText = headerRange.Text

    ' Define pattern: 4 digits followed by a dash (any dash)
    ' Dashes: Hyphen-minus "-", En-dash "–" (U+2013), Em-dash "—" (U+2014), Minus sign "-" (U+2212)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "(\d{4})([-\u2013\u2014\u2212])"
        .Global = False
        .IgnoreCase = True
    End With

    If Not regex.test(headerText) Then
        MsgBox "Header does not contain 4 digits followed by a dash." & vbCrLf & _
               "File: " & doc.Name & vbCrLf & _
               "Header text: " & headerText, vbExclamation, "Calibration_Header"
        Exit Sub
    End If

    ' Extract match: NNNN and the dash character
    Dim match As Object
    Set match = regex.Execute(headerText)(0)
    Dim fourDigits As String
    Dim dashChar As String
    fourDigits = match.SubMatches(0)
    dashChar = match.SubMatches(1)

    ' Step 4: Extract digits from filename
    Dim filename As String
    filename = doc.Name

    ' Regex for 10 digits in a row
    Dim regex10 As Object
    Set regex10 = CreateObject("VBScript.RegExp")
    With regex10
        .Pattern = "\d{10}"
        .Global = False
    End With

    Dim eightDigits As String
    If regex10.test(filename) Then
        Dim match10 As Object
        Set match10 = regex10.Execute(filename)(0)
        ' take last 8 digits from the 10-digit string
        eightDigits = Right(match10.Value, 8)
    Else
        ' Regex for 8 digits
        Dim regex8file As Object
        Set regex8file = CreateObject("VBScript.RegExp")
        With regex8file
            .Pattern = "\d{8}"
            .Global = False
        End With
        If regex8file.test(filename) Then
            eightDigits = regex8file.Execute(filename)(0).Value
        Else
            MsgBox "Filename does not contain 8 or 10 digit number." & vbCrLf & _
                   "File: " & filename, vbExclamation, "Calibration_Header"
            Exit Sub
        End If
    End If

    ' Step 5: Insert the 8 digits after NNNN- in the header, preserving formatting
    ' Find the position of the pattern NNNN + dash in header range
    Dim pos As Long
    pos = InStr(headerText, fourDigits & dashChar)
    If pos = 0 Then
        MsgBox "Pattern not found in header text after regex match. Unexpected.", vbCritical
        Exit Sub
    End If

    ' CHECK: do NOT append to NNNN- if the 8 digits are already present
    Dim nextEight As String
    nextEight = Mid(headerText, pos + Len(fourDigits & dashChar), 8)
    Dim rx8 As Object
    Set rx8 = CreateObject("VBScript.RegExp")
    With rx8
        .Pattern = "^\d{8}$"
        .Global = False
    End With

    If rx8.test(nextEight) Then
        ' already has 8 digits after the dash: skip header insertion
        GoTo SkipHeaderInsertion
    End If

    ' Move to position just after the dash to insert the 8 digits
    Dim insertPos As Long
    insertPos = pos + Len(fourDigits & dashChar) - 1

    ' Get Range object for insertion point inside headerRange
    Dim insertRange As Range
    Set insertRange = headerRange.Duplicate
    insertRange.Start = headerRange.Start + insertPos
    insertRange.End = insertRange.Start

    ' Insert the 8 digits text preserving the formatting of the dash character
    ' We'll copy the formatting of the dash character before inserting

    ' Copy formatting of the dash character
    Dim dashRange As Range
    Set dashRange = headerRange.Duplicate
    dashRange.Start = headerRange.Start + pos + Len(fourDigits) - 1
    dashRange.End = dashRange.Start + 1

    ' Insert the digits with dash formatting
    insertRange.FormattedText = dashRange.FormattedText
    insertRange.Text = eightDigits

SkipHeaderInsertion:
    ' Step 6: Insert the 8 digits after "Serial number(s) : " in main document text
    Dim bodyRange As Range
    Set bodyRange = doc.Content
    
    Dim findText As String
    findText = "Serial number(s)" & vbTab & ": "
    
    Dim found As Boolean
    found = bodyRange.Find.Execute(findText:=findText, MatchCase:=False, MatchWholeWord:=False)
    
    If found Then
        ' Move cursor to just after the match
        Dim insertBodyRange As Range
        Set insertBodyRange = bodyRange.Duplicate
        insertBodyRange.Start = bodyRange.Start + Len(findText)
        insertBodyRange.End = insertBodyRange.Start
    
        ' CHECK: do NOT insert if the next 8 characters after the colon are already 8 digits
        Dim afterPos As Long
        afterPos = bodyRange.Start + Len(findText)
        
        Dim testEndPos As Long
        testEndPos = afterPos + 8
        If testEndPos > doc.Content.End Then testEndPos = doc.Content.End
        
        Dim testRange As Range
        Set testRange = doc.Range(Start:=afterPos, End:=testEndPos)
        
        Dim testText As String
        testText = Left(testRange.Text, 8)
        
        Dim rx8body As Object
        Set rx8body = CreateObject("VBScript.RegExp")
        With rx8body
            .Pattern = "^\d{8}$"
            .Global = False
        End With
        
        If Not rx8body.test(testText) Then
            ' Optional: copy formatting from the last character in the matched text
            Dim formatRange As Range
            Set formatRange = bodyRange.Duplicate
            formatRange.Start = insertBodyRange.Start - 1
            formatRange.End = insertBodyRange.Start
        
            insertBodyRange.FormattedText = formatRange.FormattedText
            insertBodyRange.Text = eightDigits
        End If
    End If

    ' MACRO OPTIONS: closure of pause_screen_updating and single_undo
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
    
End Sub



Sub ColRezise_inSection()
    Dim sec As Section
    Dim firstTable As Table
    Dim tbl As Table
    Dim totalWidth As Single
    Dim col1Width As Single
    Dim col2Width As Single
    Dim r As Long
    
    ' Get section where cursor is
    Set sec = Selection.Sections(1)
    If sec.Range.Tables.Count = 0 Then Exit Sub
    
    ' Reference to the first table in the section
    Set firstTable = sec.Range.Tables(1)
    
    ' Store overall table width and column widths
    totalWidth = firstTable.Cell(1, 1).Width
    col1Width = firstTable.Cell(2, 1).Width
    col2Width = firstTable.Cell(2, 2).Width
    
    ' Start single undo record
    Application.UndoRecord.StartCustomRecord ("ColRezise_inSection")
    Application.ScreenUpdating = False
    
    ' Apply to all other tables in the section
    For Each tbl In sec.Range.Tables
        If Not tbl Is firstTable Then
            tbl.Cell(1, 1).Width = totalWidth
            For r = 2 To tbl.Rows.Count
                If tbl.Rows(r).Cells.Count >= 2 Then
                    tbl.Rows(r).Cells(1).Width = col1Width
                    tbl.Rows(r).Cells(2).Width = col2Width
                End If
            Next r
        End If
    Next tbl
        
    ' End single undo record
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
    
End Sub


