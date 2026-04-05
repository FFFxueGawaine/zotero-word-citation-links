Attribute VB_Name = "ZoteroWordHyperlinks"
Option Explicit

Private Const BIB_BOOKMARK As String = "ZOTERO_BIBL_ROOT"
Private Const REF_BOOKMARK_PREFIX As String = "ZOTERO_REF_"
Private Const UNLINKED_CITATION_COLOR As Long = -16777216
Private Const LINK_TARGET_PREFIX As String = "ZWL_COLOR="

Public Sub ZoteroCreateCitationLinks(Optional ByVal control As Variant)
    ApplyZoteroCitationLinksAuto
End Sub

Public Sub ZoteroRemoveCitationLinks(Optional ByVal control As Variant)
    RemoveManagedCitationLinks
End Sub

Public Sub ZoteroLinkCitationNumeric()
    ApplyZoteroCitationLinksManual True
End Sub

Public Sub ZoteroLinkCitationWholeField()
    ApplyZoteroCitationLinksManual False
End Sub

Private Sub ApplyZoteroCitationLinksAuto()
    ApplyZoteroCitationLinks True, False
End Sub

Private Sub ApplyZoteroCitationLinksManual(ByVal numericMode As Boolean)
    ApplyZoteroCitationLinks False, numericMode
End Sub

Private Sub RemoveManagedCitationLinks()
    Dim i As Long
    Dim hl As Hyperlink
    Dim bmName As String

    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        Set hl = ActiveDocument.Hyperlinks(i)
        If IsManagedBookmarkName(hl.SubAddress) Then
            RemoveHyperlinkSafely hl
        End If
    Next i

    For i = ActiveDocument.Bookmarks.Count To 1 Step -1
        bmName = ActiveDocument.Bookmarks(i).Name
        If IsManagedBookmarkName(bmName) Then
            ActiveDocument.Bookmarks(bmName).Delete
        End If
    Next i
End Sub

Private Sub ApplyZoteroCitationLinks(ByVal autoDetectMode As Boolean, ByVal numericMode As Boolean)
    Dim keepStart As Long
    Dim keepEnd As Long
    Dim oldScreenUpdating As Boolean
    Dim bibRange As Range
    Dim aField As Field
    Dim fieldCode As String
    Dim titles As Collection
    Dim useNumericMode As Boolean

    On Error GoTo CleanFail

    keepStart = Selection.Start
    keepEnd = Selection.End
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False

    Set bibRange = FindZoteroBibliographyRange()
    If bibRange Is Nothing Then
        MsgBox "Zotero bibliography was not found. Please run Zotero -> Add/Edit Bibliography first.", vbExclamation
        GoTo CleanExit
    End If

    AddOrReplaceBookmark BIB_BOOKMARK, bibRange

    For Each aField In ActiveDocument.Fields
        If InStr(1, aField.Code.Text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0 Then
            fieldCode = aField.Code.Text
            Set titles = ExtractTitles(fieldCode)
            If titles.Count > 0 Then
                If autoDetectMode Then
                    useNumericMode = ShouldLinkAsNumeric(fieldCode, aField.Result.Text)
                Else
                    useNumericMode = numericMode
                End If

                If useNumericMode Then
                    LinkNumericCitationField aField, titles, bibRange
                Else
                    LinkWholeCitationField aField, CStr(titles(1)), bibRange
                End If
            End If
        End If
    Next aField

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    ActiveDocument.Range(keepStart, keepEnd).Select
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    ActiveDocument.Range(keepStart, keepEnd).Select
    MsgBox "Macro failed: " & Err.Description, vbExclamation
End Sub

Private Function ShouldLinkAsNumeric(ByVal fieldCode As String, ByVal displayText As String) As Boolean
    Dim plainCitation As String

    plainCitation = Trim$(ExtractPlainCitation(fieldCode))

    If LooksLikeAuthorYearCitation(plainCitation) Then
        ShouldLinkAsNumeric = False
        Exit Function
    End If

    If LooksLikeNumericCitation(plainCitation) Then
        ShouldLinkAsNumeric = True
        Exit Function
    End If

    ShouldLinkAsNumeric = LooksLikeNumericCitation(displayText)
End Function

Private Function LooksLikeNumericCitation(ByVal textValue As String) As Boolean
    If Len(Trim$(textValue)) = 0 Then
        Exit Function
    End If

    If ContainsLetter(textValue) Then
        LooksLikeNumericCitation = False
    Else
        LooksLikeNumericCitation = ContainsDigit(textValue)
    End If
End Function

Private Function LooksLikeAuthorYearCitation(ByVal textValue As String) As Boolean
    LooksLikeAuthorYearCitation = ContainsLetter(textValue) And ContainsDigit(textValue)
End Function

Private Function ContainsLetter(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch Like "[A-Za-z]" Then
            ContainsLetter = True
            Exit Function
        End If
    Next i
End Function

Private Function ContainsDigit(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch Like "[0-9]" Then
            ContainsDigit = True
            Exit Function
        End If
    Next i
End Function

Private Sub LinkNumericCitationField(ByVal aField As Field, ByVal titles As Collection, ByVal bibRange As Range)
    Dim tokens As Collection
    Dim searchRange As Range
    Dim anchorRange As Range
    Dim targetTitle As String
    Dim bookmarkName As String
    Dim tooltipText As String
    Dim i As Long

    Set tokens = ExtractVisibleNumericTokens(aField.Result.Text)
    If tokens.Count = 0 Then
        LinkWholeCitationField aField, CStr(titles(1)), bibRange
        Exit Sub
    End If

    Set searchRange = aField.Result.Duplicate

    For i = 1 To tokens.Count
        targetTitle = ResolveTitleForTokenIndex(titles, tokens.Count, i)
        If Len(targetTitle) = 0 Then
            Exit For
        End If

        bookmarkName = EnsureBibliographyEntryBookmark(targetTitle, bibRange, tooltipText)
        If Len(bookmarkName) = 0 Then
            GoTo NextToken
        End If

        Set anchorRange = FindTextInRange(searchRange, CStr(tokens(i)))
        If anchorRange Is Nothing Then
            Set anchorRange = FindTextInRange(aField.Result.Duplicate, CStr(tokens(i)))
        End If

        If Not anchorRange Is Nothing Then
            AddHyperlinkToRange anchorRange, bookmarkName, tooltipText
            If anchorRange.End < aField.Result.End Then
                searchRange.Start = anchorRange.End
            End If
        End If

NextToken:
    Next i
End Sub

Private Sub LinkWholeCitationField(ByVal aField As Field, ByVal title As String, ByVal bibRange As Range)
    Dim tooltipText As String
    Dim bookmarkName As String
    Dim anchorRange As Range
    Dim useWrappedTextStyle As Boolean

    bookmarkName = EnsureBibliographyEntryBookmark(title, bibRange, tooltipText)
    If Len(bookmarkName) = 0 Then
        Exit Sub
    End If

    useWrappedTextStyle = ShouldUseWrappedTextLinkRange(aField.Result.Text)
    Set anchorRange = ResolveWholeCitationAnchorRange(aField)
    AddHyperlinkToRange anchorRange, bookmarkName, tooltipText, useWrappedTextStyle
End Sub

Private Function ResolveWholeCitationAnchorRange(ByVal aField As Field) As Range
    Dim candidateRange As Range

    Set candidateRange = aField.Result.Duplicate

    If ShouldUseWrappedTextLinkRange(candidateRange.Text) Then
        Set candidateRange = ExtractAuthorDateLinkRange(candidateRange)
    End If

    Set ResolveWholeCitationAnchorRange = candidateRange
End Function

Private Function ExtractAuthorDateLinkRange(ByVal sourceRange As Range) As Range
    Dim resultRange As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim startChar As String
    Dim endChar As String

    Set resultRange = sourceRange.Duplicate
    startPos = resultRange.Start
    endPos = resultRange.End

    Do While startPos < endPos And IsWhitespaceCharacter(ActiveDocument.Range(startPos, startPos + 1).Text)
        startPos = startPos + 1
    Loop

    Do While endPos > startPos And IsWhitespaceCharacter(ActiveDocument.Range(endPos - 1, endPos).Text)
        endPos = endPos - 1
    Loop

    If endPos <= startPos Then
        Set ExtractAuthorDateLinkRange = sourceRange.Duplicate
        Exit Function
    End If

    startChar = ActiveDocument.Range(startPos, startPos + 1).Text
    endChar = ActiveDocument.Range(endPos - 1, endPos).Text

    If IsMatchingWrapper(startChar, endChar) And endPos - startPos > 2 Then
        startPos = startPos + 1
        endPos = endPos - 1
    End If

    Do While startPos < endPos And IsWhitespaceCharacter(ActiveDocument.Range(startPos, startPos + 1).Text)
        startPos = startPos + 1
    Loop

    Do While endPos > startPos And IsWhitespaceCharacter(ActiveDocument.Range(endPos - 1, endPos).Text)
        endPos = endPos - 1
    Loop

    If endPos <= startPos Then
        Set ExtractAuthorDateLinkRange = sourceRange.Duplicate
    Else
        Set ExtractAuthorDateLinkRange = ActiveDocument.Range(startPos, endPos)
    End If
End Function

Private Function IsMatchingWrapper(ByVal startChar As String, ByVal endChar As String) As Boolean
    IsMatchingWrapper = (startChar = "(" And endChar = ")") _
        Or (startChar = "[" And endChar = "]")
End Function

Private Function IsWhitespaceCharacter(ByVal textValue As String) As Boolean
    IsWhitespaceCharacter = Len(Trim$(textValue)) = 0
End Function

Private Function ShouldUseWrappedTextLinkRange(ByVal textValue As String) As Boolean
    Dim trimmedText As String
    Dim startChar As String
    Dim endChar As String

    trimmedText = Trim$(textValue)
    If Len(trimmedText) < 3 Then
        Exit Function
    End If

    If Not ContainsLetter(trimmedText) Then
        Exit Function
    End If

    startChar = Left$(trimmedText, 1)
    endChar = Right$(trimmedText, 1)
    ShouldUseWrappedTextLinkRange = IsMatchingWrapper(startChar, endChar)
End Function

Private Function ExtractVisibleNumericTokens(ByVal textValue As String) As Collection
    Dim results As New Collection
    Dim i As Long
    Dim ch As String
    Dim tokenText As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch Like "[0-9]" Then
            tokenText = tokenText & ch
        ElseIf Len(tokenText) > 0 Then
            results.Add tokenText
            tokenText = ""
        End If
    Next i

    If Len(tokenText) > 0 Then
        results.Add tokenText
    End If

    Set ExtractVisibleNumericTokens = results
End Function

Private Function ResolveTitleForTokenIndex(ByVal titles As Collection, ByVal tokenCount As Long, ByVal tokenIndex As Long) As String
    If titles.Count = 0 Then
        ResolveTitleForTokenIndex = ""
        Exit Function
    End If

    If titles.Count > tokenCount Then
        If tokenIndex = 1 Then
            ResolveTitleForTokenIndex = CStr(titles(1))
        ElseIf tokenIndex = tokenCount Then
            ResolveTitleForTokenIndex = CStr(titles(titles.Count))
        ElseIf tokenIndex <= titles.Count Then
            ResolveTitleForTokenIndex = CStr(titles(tokenIndex))
        Else
            ResolveTitleForTokenIndex = ""
        End If
    ElseIf tokenIndex <= titles.Count Then
        ResolveTitleForTokenIndex = CStr(titles(tokenIndex))
    Else
        ResolveTitleForTokenIndex = ""
    End If
End Function

Private Function FindZoteroBibliographyRange() As Range
    Dim aField As Field

    For Each aField In ActiveDocument.Fields
        If InStr(1, aField.Code.Text, "ADDIN ZOTERO_BIBL", vbTextCompare) > 0 Then
            Set FindZoteroBibliographyRange = aField.Result.Duplicate
            Exit Function
        End If
    Next aField
End Function

Private Function ExtractPlainCitation(ByVal fieldCode As String) As String
    ExtractPlainCitation = ExtractJsonValue(fieldCode, "plainCitation")
End Function

Private Function ExtractTitles(ByVal fieldCode As String) As Collection
    Dim results As New Collection
    Dim startPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim titleText As String
    Dim searchFrom As Long

    searchFrom = 1
    Do
        startPos = InStr(searchFrom, fieldCode, """title"":""", vbTextCompare)
        If startPos = 0 Then
            Exit Do
        End If

        valueStart = startPos + Len("""title"":""")
        valueEnd = FindJsonStringEnd(fieldCode, valueStart)
        If valueEnd = 0 Then
            Exit Do
        End If

        titleText = Mid$(fieldCode, valueStart, valueEnd - valueStart)
        titleText = JsonUnescape(titleText)
        If Len(titleText) > 0 Then
            results.Add titleText
        End If

        searchFrom = valueEnd + 1
    Loop

    Set ExtractTitles = results
End Function

Private Function EnsureBibliographyEntryBookmark(ByVal title As String, ByVal bibRange As Range, ByRef tooltipText As String) As String
    Dim entryRange As Range
    Dim bookmarkName As String

    Set entryRange = FindTextInRange(bibRange.Duplicate, Left$(title, 255))
    If entryRange Is Nothing Then
        EnsureBibliographyEntryBookmark = ""
        Exit Function
    End If

    Set entryRange = entryRange.Paragraphs(1).Range.Duplicate
    If entryRange.End > entryRange.Start Then
        entryRange.End = entryRange.End - 1
    End If

    tooltipText = Left$(entryRange.Text, 120)
    bookmarkName = MakeBookmarkName(title)
    AddOrReplaceBookmark bookmarkName, entryRange
    EnsureBibliographyEntryBookmark = bookmarkName
End Function

Private Function FindTextInRange(ByVal targetRange As Range, ByVal searchText As String) As Range
    Dim workRange As Range

    If Len(searchText) = 0 Then
        Exit Function
    End If

    Set workRange = targetRange.Duplicate
    With workRange.Find
        .ClearFormatting
        .Text = searchText
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    If workRange.Find.Execute Then
        Set FindTextInRange = workRange.Duplicate
    End If
End Function

Private Sub AddHyperlinkToRange(ByVal anchorRange As Range, ByVal bookmarkName As String, ByVal tooltipText As String, Optional ByVal clearDirectFormatting As Boolean = False)
    Dim i As Long
    Dim startPos As Long
    Dim linkText As String
    Dim storedColor As Long
    Dim newRange As Range
    Dim createdLink As Hyperlink
    Dim formatSnapshot As Variant

    startPos = anchorRange.Start
    linkText = anchorRange.Text
    storedColor = ResolveStoredColor(anchorRange)
    If clearDirectFormatting Then
        formatSnapshot = CaptureCharacterFormattingSnapshot(anchorRange)
    End If

    For i = anchorRange.Hyperlinks.Count To 1 Step -1
        RemoveHyperlinkSafely anchorRange.Hyperlinks(i)
    Next i

    Set newRange = ActiveDocument.Range(startPos, startPos + Len(linkText))

    Set createdLink = ActiveDocument.Hyperlinks.Add( _
        Anchor:=newRange, _
        Address:="", _
        SubAddress:=bookmarkName, _
        Target:=BuildLinkTarget(storedColor), _
        ScreenTip:=tooltipText, _
        TextToDisplay:=linkText)

    Set newRange = ActiveDocument.Range(startPos, startPos + Len(linkText))
    If clearDirectFormatting Then
        createdLink.Range.Select
        Selection.ClearCharacterDirectFormatting
        Set newRange = createdLink.Range.Duplicate
        RestoreCharacterFormattingFromSnapshot newRange, formatSnapshot
    End If
    ApplyLinkedCitationAppearance newRange
End Sub

Private Sub ApplyLinkedCitationAppearance(ByVal targetRange As Range)
    Dim i As Long

    targetRange.Font.Color = vbBlue
    targetRange.Font.Underline = wdUnderlineNone

    For i = 1 To targetRange.Characters.Count
        With targetRange.Characters(i).Font
            .Color = vbBlue
            .Underline = wdUnderlineNone
        End With
    Next i
End Sub

Private Sub ApplyUnlinkedCitationAppearance(ByVal targetRange As Range, Optional ByVal hasStoredColor As Boolean = False, Optional ByVal storedColor As Long = 0)
    Dim i As Long
    Dim targetColor As Long

    If hasStoredColor Then
        targetColor = storedColor
    Else
        targetColor = UNLINKED_CITATION_COLOR
    End If

    targetRange.Font.Color = targetColor
    targetRange.Font.Underline = wdUnderlineNone

    For i = 1 To targetRange.Characters.Count
        With targetRange.Characters(i).Font
            .Color = targetColor
            .Underline = wdUnderlineNone
        End With
    Next i
End Sub

Private Sub RemoveHyperlinkSafely(ByVal hl As Hyperlink)
    Dim displayText As String
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim targetStart As Long
    Dim hadToRewrite As Boolean
    Dim storedColor As Long
    Dim hasStoredColor As Boolean

    Set sourceRange = hl.Range.Duplicate
    Set targetRange = hl.Range.Duplicate
    targetStart = targetRange.Start
    displayText = targetRange.Text
    hasStoredColor = TryParseStoredColor(hl.Target, storedColor)

    On Error Resume Next
    hl.Delete
    Err.Clear
    On Error GoTo 0

    If targetRange.Hyperlinks.Count > 0 Then
        hadToRewrite = True
        targetRange.Text = displayText
    End If

    Set targetRange = ActiveDocument.Range(targetStart, targetStart + Len(displayText))
    If hadToRewrite Then
        RestoreCharacterFormatting targetRange, sourceRange
    End If
    ApplyUnlinkedCitationAppearance targetRange, hasStoredColor, storedColor
End Sub

Private Function ResolveStoredColor(ByVal anchorRange As Range) As Long
    Dim existingLink As Hyperlink
    Dim existingColor As Long

    If anchorRange.Hyperlinks.Count > 0 Then
        Set existingLink = anchorRange.Hyperlinks(1)
        If TryParseStoredColor(existingLink.Target, existingColor) Then
            ResolveStoredColor = existingColor
            Exit Function
        End If
    End If

    ResolveStoredColor = GetPrimaryColor(anchorRange)
End Function

Private Function GetPrimaryColor(ByVal targetRange As Range) As Long
    On Error GoTo Fallback

    If targetRange.Characters.Count > 0 Then
        GetPrimaryColor = CLng(targetRange.Characters(1).Font.Color)
    Else
        GetPrimaryColor = UNLINKED_CITATION_COLOR
    End If
    Exit Function

Fallback:
    GetPrimaryColor = UNLINKED_CITATION_COLOR
End Function

Private Function BuildLinkTarget(ByVal storedColor As Long) As String
    BuildLinkTarget = LINK_TARGET_PREFIX & CStr(storedColor)
End Function

Private Function TryParseStoredColor(ByVal targetValue As String, ByRef storedColor As Long) As Boolean
    Dim rawValue As String

    If Left$(targetValue, Len(LINK_TARGET_PREFIX)) <> LINK_TARGET_PREFIX Then
        Exit Function
    End If

    rawValue = Mid$(targetValue, Len(LINK_TARGET_PREFIX) + 1)
    If Len(rawValue) = 0 Then
        Exit Function
    End If

    On Error GoTo ParseFail
    storedColor = CLng(rawValue)
    TryParseStoredColor = True
    Exit Function

ParseFail:
    TryParseStoredColor = False
End Function

Private Sub RestoreCharacterFormatting(ByVal targetRange As Range, ByVal sourceRange As Range)
    Dim i As Long
    Dim charCount As Long

    CopyFontFormatting targetRange.Font, sourceRange.Font

    charCount = targetRange.Characters.Count
    If sourceRange.Characters.Count < charCount Then
        charCount = sourceRange.Characters.Count
    End If

    For i = 1 To charCount
        CopyFontFormatting targetRange.Characters(i).Font, sourceRange.Characters(i).Font
    Next i
End Sub

Private Sub CopyFontFormatting(ByVal targetFont As Font, ByVal sourceFont As Font)
    With targetFont
        .Name = sourceFont.Name
        .Size = sourceFont.Size
        .Bold = sourceFont.Bold
        .Italic = sourceFont.Italic
        .Superscript = sourceFont.Superscript
        .Subscript = sourceFont.Subscript
        .Position = sourceFont.Position
        .Scaling = sourceFont.Scaling
        .Spacing = sourceFont.Spacing
        .SmallCaps = sourceFont.SmallCaps
        .AllCaps = sourceFont.AllCaps
        .StrikeThrough = sourceFont.StrikeThrough
        .DoubleStrikeThrough = sourceFont.DoubleStrikeThrough
        .Hidden = sourceFont.Hidden
        .Outline = sourceFont.Outline
        .Emboss = sourceFont.Emboss
        .Shadow = sourceFont.Shadow
        .Kerning = sourceFont.Kerning
    End With
End Sub

Private Function CaptureCharacterFormattingSnapshot(ByVal sourceRange As Range) As Variant
    Dim snapshot() As Variant
    Dim charCount As Long
    Dim i As Long

    charCount = sourceRange.Characters.Count
    ReDim snapshot(1 To charCount, 1 To 15)

    For i = 1 To charCount
        With sourceRange.Characters(i).Font
            snapshot(i, 1) = .Name
            snapshot(i, 2) = .Size
            snapshot(i, 3) = .Bold
            snapshot(i, 4) = .Italic
            snapshot(i, 5) = .Superscript
            snapshot(i, 6) = .Subscript
            snapshot(i, 7) = .Position
            snapshot(i, 8) = .Scaling
            snapshot(i, 9) = .Spacing
            snapshot(i, 10) = .SmallCaps
            snapshot(i, 11) = .AllCaps
            snapshot(i, 12) = .StrikeThrough
            snapshot(i, 13) = .DoubleStrikeThrough
            snapshot(i, 14) = .Hidden
            snapshot(i, 15) = .Kerning
        End With
    Next i

    CaptureCharacterFormattingSnapshot = snapshot
End Function

Private Sub RestoreCharacterFormattingFromSnapshot(ByVal targetRange As Range, ByVal snapshot As Variant)
    Dim i As Long
    Dim charCount As Long

    On Error GoTo SnapshotMissing
    charCount = UBound(snapshot, 1)
    If targetRange.Characters.Count < charCount Then
        charCount = targetRange.Characters.Count
    End If

    For i = 1 To charCount
        With targetRange.Characters(i).Font
            .Name = snapshot(i, 1)
            .Size = snapshot(i, 2)
            .Bold = snapshot(i, 3)
            .Italic = snapshot(i, 4)
            .Superscript = snapshot(i, 5)
            .Subscript = snapshot(i, 6)
            .Position = snapshot(i, 7)
            .Scaling = snapshot(i, 8)
            .Spacing = snapshot(i, 9)
            .SmallCaps = snapshot(i, 10)
            .AllCaps = snapshot(i, 11)
            .StrikeThrough = snapshot(i, 12)
            .DoubleStrikeThrough = snapshot(i, 13)
            .Hidden = snapshot(i, 14)
            .Kerning = snapshot(i, 15)
        End With
    Next i
    Exit Sub

SnapshotMissing:
    Err.Clear
End Sub

Private Sub AddOrReplaceBookmark(ByVal bookmarkName As String, ByVal bookmarkRange As Range)
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        ActiveDocument.Bookmarks(bookmarkName).Delete
    End If
    ActiveDocument.Bookmarks.Add Name:=bookmarkName, Range:=bookmarkRange
End Sub

Private Function MakeBookmarkName(ByVal sourceText As String) As String
    Dim i As Long
    Dim ch As String
    Dim cleaned As String

    sourceText = Trim$(sourceText)
    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            cleaned = cleaned & ch
        Else
            cleaned = cleaned & "_"
        End If
    Next i

    Do While InStr(cleaned, "__") > 0
        cleaned = Replace(cleaned, "__", "_")
    Loop

    If Len(cleaned) = 0 Then
        cleaned = "REF"
    End If

    MakeBookmarkName = Left$(REF_BOOKMARK_PREFIX & cleaned, 40)
End Function

Private Function IsManagedBookmarkName(ByVal bookmarkName As String) As Boolean
    IsManagedBookmarkName = (bookmarkName = BIB_BOOKMARK) Or _
                            (Left$(bookmarkName, Len(REF_BOOKMARK_PREFIX)) = REF_BOOKMARK_PREFIX)
End Function

Private Function ExtractJsonValue(ByVal fieldCode As String, ByVal keyName As String) As String
    Dim keyText As String
    Dim valueStart As Long
    Dim valueEnd As Long

    keyText = """" & keyName & """:"""
    valueStart = InStr(1, fieldCode, keyText, vbTextCompare)
    If valueStart = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    valueStart = valueStart + Len(keyText)
    valueEnd = FindJsonStringEnd(fieldCode, valueStart)
    If valueEnd = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    ExtractJsonValue = JsonUnescape(Mid$(fieldCode, valueStart, valueEnd - valueStart))
End Function

Private Function FindJsonStringEnd(ByVal textValue As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim currentChar As String
    Dim escapeActive As Boolean

    For i = startPos To Len(textValue)
        currentChar = Mid$(textValue, i, 1)
        If escapeActive Then
            escapeActive = False
        ElseIf currentChar = "\" Then
            escapeActive = True
        ElseIf currentChar = """" Then
            FindJsonStringEnd = i
            Exit Function
        End If
    Next i
End Function

Private Function JsonUnescape(ByVal textValue As String) As String
    textValue = Replace(textValue, "\\", "\")
    textValue = Replace(textValue, "\" & Chr$(34), Chr$(34))
    textValue = Replace(textValue, "\/", "/")
    textValue = Replace(textValue, "\u2013", ChrW$(&H2013))
    textValue = Replace(textValue, "\u2014", ChrW$(&H2014))
    textValue = Replace(textValue, "\u2018", "'")
    textValue = Replace(textValue, "\u2019", "'")
    textValue = Replace(textValue, "\u201C", ChrW$(&H201C))
    textValue = Replace(textValue, "\u201D", ChrW$(&H201D))
    textValue = Replace(textValue, "\u00A0", " ")
    JsonUnescape = textValue
End Function
