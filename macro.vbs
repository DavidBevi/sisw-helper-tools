Sub Calibration_Header()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim i As Long
    Dim firstSection As Section
    Set firstSection = doc.Sections(1)

    ' Step 1: Unify all headers and footers to the first section headers/footers
    For i = 2 To doc.Sections.count
        ' Link headers and footers to the first section
        doc.Sections(i).Headers(wdHeaderFooterPrimary).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterPrimary).LinkToPrevious = True
        doc.Sections(i).Headers(wdHeaderFooterFirstPage).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterFirstPage).LinkToPrevious = True
        doc.Sections(i).Headers(wdHeaderFooterEvenPages).LinkToPrevious = True
        doc.Sections(i).Footers(wdHeaderFooterEvenPages).LinkToPrevious = True
    Next i

    ' Refresh headers/footers links
    doc.Fields.Update

    ' Step 2: Get header text of first section primary header
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

    ' Step 3: Extract digits from filename
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
        Dim regex8 As Object
        Set regex8 = CreateObject("VBScript.RegExp")
        With regex8
            .Pattern = "\d{8}"
            .Global = False
        End With
        If regex8.test(filename) Then
            eightDigits = regex8.Execute(filename)(0).Value
        Else
            MsgBox "Filename does not contain 8 or 10 digit number." & vbCrLf & _
                   "File: " & filename, vbExclamation, "Calibration_Header"
            Exit Sub
        End If
    End If

    ' Step 4: Insert the 8 digits after NNNN- in the header, preserving formatting

    ' Find the position of the pattern NNNN + dash in header range
    Dim pos As Long
    pos = InStr(headerText, fourDigits & dashChar)
    If pos = 0 Then
        MsgBox "Pattern not found in header text after regex match. Unexpected.", vbCritical
        Exit Sub
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

    ' Alternatively, insert text and apply formatting manually:
    'insertRange.Text = eightDigits
    'insertRange.Font.Name = dashRange.Font.Name
    'insertRange.Font.Size = dashRange.Font.Size
    '... (other font properties if needed)
    
    ' Step 5: Insert the 8 digits after "Serial number(s) : " in main document text
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
    
        ' Optional: copy formatting from the last character in the matched text
        Dim formatRange As Range
        Set formatRange = bodyRange.Duplicate
        formatRange.Start = insertBodyRange.Start - 1
        formatRange.End = insertBodyRange.Start
    
        insertBodyRange.FormattedText = formatRange.FormattedText
        insertBodyRange.Text = eightDigits
End If
    
End Sub


Sub Sections()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim para As Paragraph
    Dim sectionTitles As Collection
    Set sectionTitles = New Collection

    Dim paraStyles As Collection
    Set paraStyles = New Collection

    Dim paraIndices As Collection
    Set paraIndices = New Collection

    Dim paraCounter As Long
    paraCounter = 0

    ' Collect Heading 1 and Heading 2 paragraphs, save their paragraph indices
    For Each para In doc.Paragraphs
        paraCounter = paraCounter + 1
        If para.Style = "Heading 1" Or para.Style = "Heading 2" Then
            sectionTitles.Add Trim(para.Range.Text)
            paraStyles.Add para.Style
            paraIndices.Add paraCounter ' Manual paragraph index
        End If
    Next para

    ' Find the index of "System configuration" (case-insensitive)
    Dim targetIndex As Long
    targetIndex = 0
    Dim i As Long
    For i = 1 To sectionTitles.count
        If LCase(sectionTitles(i)) Like "system configuration*" Then
            targetIndex = i
            Exit For
        End If
    Next i

    If targetIndex = 0 Then
        MsgBox "'System configuration' section not found.", vbExclamation
        Exit Sub
    End If

    ' Prepare message listing all sections/subsections after "System configuration"
    Dim msg As String
    msg = "Sections and subsections after 'System configuration':" & vbCrLf & vbCrLf

    Dim tblCount As Long
    Dim rngSubsection As Range
    Dim startParaIndex As Long
    Dim nextHeadingIndex As Long

    For i = targetIndex + 1 To sectionTitles.count
        If paraStyles(i) = "Heading 2" Then
            ' Calculate subsection range from current Heading 2 to the next Heading 1 or Heading 2 or document end
            startParaIndex = paraIndices(i)
            nextHeadingIndex = doc.Paragraphs.count + 1 ' default: end of doc

            Dim j As Long
            For j = startParaIndex + 1 To doc.Paragraphs.count
                If doc.Paragraphs(j).Style = "Heading 1" Or doc.Paragraphs(j).Style = "Heading 2" Then
                    nextHeadingIndex = j
                    Exit For
                End If
            Next j

            If nextHeadingIndex > doc.Paragraphs.count Then
                Set rngSubsection = doc.Range(Start:=doc.Paragraphs(startParaIndex).Range.Start, _
                                             End:=doc.Content.End)
            Else
                Set rngSubsection = doc.Range(Start:=doc.Paragraphs(startParaIndex).Range.Start, _
                                             End:=doc.Paragraphs(nextHeadingIndex).Range.Start - 1)
            End If

            tblCount = rngSubsection.Tables.count

            msg = msg & (i - targetIndex) & ". " & sectionTitles(i) & " (Tables: " & tblCount & ")" & vbCrLf
        Else
            msg = msg & (i - targetIndex) & ". " & sectionTitles(i) & vbCrLf
        End If
    Next i

    ' Show MsgBox with OK (go to next section) or Cancel (stop)
    Dim answer As VbMsgBoxResult
    answer = MsgBox(msg, vbOKCancel + vbInformation, "Sections After System Configuration")

    If answer = vbCancel Then Exit Sub

    ' Target section title to jump to (the first one after System configuration)
    Dim targetTitle As String
    targetTitle = sectionTitles(targetIndex + 1)

    ' Find paragraph object for target title and the previous heading paragraph
    Dim targetPara As Paragraph
    Dim prevPara As Paragraph
    Set targetPara = Nothing
    Set prevPara = Nothing

    For Each para In doc.Paragraphs
        If (para.Style = "Heading 1" Or para.Style = "Heading 2") Then
            If Trim(para.Range.Text) = targetTitle Then
                Set targetPara = para
                Exit For
            End If
        End If
    Next para

    If targetPara Is Nothing Then
        MsgBox "Next section heading not found.", vbExclamation
        Exit Sub
    End If

    ' Find previous heading paragraph (Heading 1 or Heading 2) before targetPara
    Dim prevParaCandidate As Paragraph
    For Each para In doc.Paragraphs
        If (para.Style = "Heading 1" Or para.Style = "Heading 2") Then
            If para.Range.Start < targetPara.Range.Start Then
                Set prevParaCandidate = para
            Else
                Exit For
            End If
        End If
    Next para
    Set prevPara = prevParaCandidate

    ' Scroll workaround:
    If Not prevPara Is Nothing Then
        prevPara.Range.Select
        ActiveWindow.ScrollIntoView Selection.Range, True ' Scroll previous heading to top
    Else
        targetPara.Range.Select
        ActiveWindow.ScrollIntoView Selection.Range, True ' Scroll target heading to top if no previous heading
    End If

    ' Now select target heading and scroll so it is visible (stabilize)
    targetPara.Range.Select
    ActiveWindow.ScrollIntoView targetPara.Range, False

End Sub
