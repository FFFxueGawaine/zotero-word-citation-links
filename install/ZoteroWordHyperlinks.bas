Attribute VB_Name = "ZoteroWordHyperlinks"
Option Explicit

Private Const BIB_BOOKMARK As String = "ZOTERO_BIBL_ROOT"
Private Const REF_BOOKMARK_PREFIX As String = "ZOTERO_REF_"
Private Const DEFAULT_LINK_COLOR As Long = vbBlue
Private Const UNLINKED_CITATION_COLOR As Long = -16777216
Private Const LINK_COLOR_VARIABLE As String = "ZWL_LINK_COLOR"
Private Const LEGACY_LINK_TARGET_PREFIX As String = "ZWL_COLOR="
Private Const FORMAT_TARGET_PREFIX As String = "ZWL_FMT="
Private Const FORMAT_VARIABLE_PREFIX As String = "ZWL_FMT_"
Private Const NEXT_FORMAT_ID_VARIABLE As String = "ZWL_NEXT_FMT_ID"
Private Const LINK_STYLE_NAME As String = "Zotero Citation Link"

Public Sub ZoteroCreateCitationLinks(Optional ByVal control As Variant)
    ApplyZoteroCitationLinksAuto
End Sub

Public Sub ZoteroRemoveCitationLinks(Optional ByVal control As Variant)
    RemoveManagedCitationLinks
End Sub

Public Sub ZoteroSetLinkColor(Optional ByVal control As Variant)
    MsgBox "Link appearance is now controlled by the current document character style '" & LINK_STYLE_NAME & "'." & vbCrLf & vbCrLf & _
        "Open the Styles pane in Word and edit that style to change the link font, size, color, or other character formatting.", vbInformation
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
    RemoveManagedCitationArtifacts ActiveDocument
End Sub

Private Sub RemoveManagedCitationArtifacts(ByVal targetDocument As Document)
    Dim i As Long
    Dim hl As Hyperlink
    Dim bmName As String

    For i = targetDocument.Hyperlinks.Count To 1 Step -1
        Set hl = targetDocument.Hyperlinks(i)
        If IsManagedBookmarkName(hl.SubAddress) Then
            RemoveHyperlinkSafely targetDocument, hl
        End If
    Next i

    For i = targetDocument.Bookmarks.Count To 1 Step -1
        bmName = targetDocument.Bookmarks(i).Name
        If IsManagedBookmarkName(bmName) Then
            targetDocument.Bookmarks(bmName).Delete
        End If
    Next i

    DeleteManagedFormattingVariables targetDocument
End Sub

Private Sub SetDefaultLinkColorInteractive()
    Dim currentColor As Long
    Dim selectedColor As Long
    Dim selectedLabel As String

    currentColor = GetConfiguredLinkColor()

    If Not TryPromptForLinkColorWithOfficeDialog(selectedColor, selectedLabel) Then
        If Not PromptForLinkColorSelection(selectedColor, selectedLabel) Then
            Exit Sub
        End If
    End If

    If selectedColor = currentColor Then
        MsgBox "Default link color was not changed.", vbInformation
        Exit Sub
    End If

    On Error GoTo SaveFailed
    SaveConfiguredLinkColor selectedColor
    MsgBox "Default link color updated to " & selectedLabel & "." & vbCrLf & vbCrLf & _
        "The new color will be used the next time you run Create Citation Links.", vbInformation
    Exit Sub

SaveFailed:
    MsgBox "Unable to save the link color setting: " & Err.Description, vbExclamation
End Sub

Private Function TryPromptForLinkColorWithOfficeDialog(ByRef selectedColor As Long, ByRef selectedLabel As String) As Boolean
    Dim originalDoc As Document
    Dim scratchDoc As Document
    Dim scratchRange As Range
    Dim originalStart As Long
    Dim originalEnd As Long
    Dim currentColor As Long
    Dim chosenColor As Long
    Dim hadOriginalDoc As Boolean

    On Error GoTo NativeDialogFailed

    currentColor = GetConfiguredLinkColor()

    If Documents.Count > 0 Then
        Set originalDoc = ActiveDocument
        hadOriginalDoc = True
        originalStart = Selection.Start
        originalEnd = Selection.End
    End If

    Set scratchDoc = Documents.Add
    Set scratchRange = scratchDoc.Range(0, 0)
    scratchRange.Text = "Zotero Link Color"
    Set scratchRange = scratchDoc.Range(0, Len("Zotero Link Color"))
    scratchRange.Select
    Selection.Font.Color = currentColor
    Selection.Font.Underline = wdUnderlineNone

    Application.CommandBars.ExecuteMso "FontColorMoreColorsDialog"

    chosenColor = CLng(Selection.Font.Color)
    If chosenColor = wdUndefined Then
        GoTo NativeDialogFailed
    End If

    If scratchDoc Is Nothing Then
        GoTo NativeDialogFailed
    End If

    If chosenColor = currentColor Then
        GoTo NativeDialogFailed
    End If

    selectedColor = chosenColor
    selectedLabel = DescribeColorValue(chosenColor)
    TryPromptForLinkColorWithOfficeDialog = True

NativeDialogCleanup:
    On Error Resume Next
    If Not scratchDoc Is Nothing Then
        scratchDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If hadOriginalDoc Then
        originalDoc.Activate
        originalDoc.Range(originalStart, originalEnd).Select
    End If
    On Error GoTo 0
    Exit Function

NativeDialogFailed:
    TryPromptForLinkColorWithOfficeDialog = False
    Resume NativeDialogCleanup
End Function

Private Function PromptForLinkColorSelection(ByRef selectedColor As Long, ByRef selectedLabel As String) As Boolean
    Dim promptText As String
    Dim userChoice As String

    promptText = "Choose the default citation link color." & vbCrLf & vbCrLf & _
        "Current: " & DescribeColorValue(GetConfiguredLinkColor()) & vbCrLf & vbCrLf & _
        "1 - Blue" & vbCrLf & _
        "2 - Black" & vbCrLf & _
        "3 - Dark Red" & vbCrLf & _
        "4 - Dark Green" & vbCrLf & _
        "5 - Orange" & vbCrLf & _
        "6 - Custom RGB" & vbCrLf & vbCrLf & _
        "Enter a number, or leave blank to cancel."

    Do
        userChoice = Trim$(InputBox(promptText, "Set Link Color", "1"))
        If Len(userChoice) = 0 Then
            Exit Function
        End If

        Select Case userChoice
            Case "1"
                selectedColor = DEFAULT_LINK_COLOR
                selectedLabel = "Blue"
                PromptForLinkColorSelection = True
                Exit Function
            Case "2"
                selectedColor = RGB(0, 0, 0)
                selectedLabel = "Black"
                PromptForLinkColorSelection = True
                Exit Function
            Case "3"
                selectedColor = RGB(192, 0, 0)
                selectedLabel = "Dark Red"
                PromptForLinkColorSelection = True
                Exit Function
            Case "4"
                selectedColor = RGB(0, 112, 60)
                selectedLabel = "Dark Green"
                PromptForLinkColorSelection = True
                Exit Function
            Case "5"
                selectedColor = RGB(230, 120, 0)
                selectedLabel = "Orange"
                PromptForLinkColorSelection = True
                Exit Function
            Case "6"
                If PromptForCustomRgbColor(selectedColor, selectedLabel) Then
                    PromptForLinkColorSelection = True
                End If
                Exit Function
            Case Else
                MsgBox "Please enter 1, 2, 3, 4, 5, or 6.", vbExclamation
        End Select
    Loop
End Function

Private Function PromptForCustomRgbColor(ByRef selectedColor As Long, ByRef selectedLabel As String) As Boolean
    Dim inputValue As String
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    Do
        inputValue = Trim$(InputBox( _
            "Enter a custom RGB color as R,G,B." & vbCrLf & _
            "Example: 220,20,60" & vbCrLf & vbCrLf & _
            "Leave blank to cancel.", _
            "Custom Link Color"))

        If Len(inputValue) = 0 Then
            Exit Function
        End If

        If TryParseRgbInput(inputValue, redValue, greenValue, blueValue) Then
            selectedColor = RGB(redValue, greenValue, blueValue)
            selectedLabel = "Custom RGB (" & redValue & "," & greenValue & "," & blueValue & ")"
            PromptForCustomRgbColor = True
            Exit Function
        End If

        MsgBox "Please enter three integers between 0 and 255, for example 220,20,60.", vbExclamation
    Loop
End Function

Private Function TryParseRgbInput(ByVal inputValue As String, ByRef redValue As Long, ByRef greenValue As Long, ByRef blueValue As Long) As Boolean
    Dim parts() As String

    parts = Split(inputValue, ",")
    If UBound(parts) - LBound(parts) <> 2 Then
        Exit Function
    End If

    If Not TryParseRgbPart(parts(0), redValue) Then
        Exit Function
    End If

    If Not TryParseRgbPart(parts(1), greenValue) Then
        Exit Function
    End If

    If Not TryParseRgbPart(parts(2), blueValue) Then
        Exit Function
    End If

    TryParseRgbInput = True
End Function

Private Function TryParseRgbPart(ByVal rawValue As String, ByRef componentValue As Long) As Boolean
    Dim cleanedValue As String

    cleanedValue = Trim$(rawValue)
    If Len(cleanedValue) = 0 Then
        Exit Function
    End If

    If Not cleanedValue Like "[0-9]*" Then
        Exit Function
    End If

    On Error GoTo ParseFailed
    componentValue = CLng(cleanedValue)
    If componentValue < 0 Or componentValue > 255 Then
        Exit Function
    End If

    TryParseRgbPart = True
    Exit Function

ParseFailed:
    TryParseRgbPart = False
End Function

Private Function GetConfiguredLinkColor() As Long
    Dim configuredColor As Long

    If TryGetConfiguredLinkColor(configuredColor) Then
        GetConfiguredLinkColor = configuredColor
    Else
        GetConfiguredLinkColor = DEFAULT_LINK_COLOR
    End If
End Function

Private Function TryGetConfiguredLinkColor(ByRef configuredColor As Long) As Boolean
    Dim rawValue As String

    On Error GoTo MissingValue
    rawValue = Trim$(ThisDocument.Variables(LINK_COLOR_VARIABLE).Value)
    If Len(rawValue) = 0 Then
        Exit Function
    End If

    configuredColor = CLng(rawValue)
    TryGetConfiguredLinkColor = True
    Exit Function

MissingValue:
    TryGetConfiguredLinkColor = False
End Function

Private Sub SaveConfiguredLinkColor(ByVal colorValue As Long)
    Dim colorText As String

    colorText = CStr(colorValue)

    On Error Resume Next
    ThisDocument.Variables(LINK_COLOR_VARIABLE).Value = colorText
    If Err.Number <> 0 Then
        Err.Clear
        ThisDocument.Variables.Add Name:=LINK_COLOR_VARIABLE, Value:=colorText
    End If
    On Error GoTo 0

    ThisDocument.Save
End Sub

Private Function DescribeColorValue(ByVal colorValue As Long) As String
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    Dim presetName As String

    GetRgbParts colorValue, redValue, greenValue, blueValue
    presetName = GetPresetColorName(colorValue)

    If Len(presetName) > 0 Then
        DescribeColorValue = presetName & " (" & redValue & "," & greenValue & "," & blueValue & ")"
    Else
        DescribeColorValue = "RGB (" & redValue & "," & greenValue & "," & blueValue & ")"
    End If
End Function

Private Function GetPresetColorName(ByVal colorValue As Long) As String
    Select Case colorValue
        Case DEFAULT_LINK_COLOR
            GetPresetColorName = "Blue"
        Case RGB(0, 0, 0)
            GetPresetColorName = "Black"
        Case RGB(192, 0, 0)
            GetPresetColorName = "Dark Red"
        Case RGB(0, 112, 60)
            GetPresetColorName = "Dark Green"
        Case RGB(230, 120, 0)
            GetPresetColorName = "Orange"
    End Select
End Function

Private Sub GetRgbParts(ByVal colorValue As Long, ByRef redValue As Long, ByRef greenValue As Long, ByRef blueValue As Long)
    redValue = colorValue And &HFF&
    greenValue = (colorValue \ &H100&) And &HFF&
    blueValue = (colorValue \ &H10000) And &HFF&
End Sub

Private Sub ApplyZoteroCitationLinks(ByVal autoDetectMode As Boolean, ByVal numericMode As Boolean)
    Dim targetDocument As Document
    Dim keepStart As Long
    Dim keepEnd As Long
    Dim oldScreenUpdating As Boolean
    Dim bibRange As Range
    Dim aField As Field
    Dim fieldCode As String
    Dim titles As Collection
    Dim useNumericMode As Boolean

    On Error GoTo CleanFail

    Set targetDocument = ActiveDocument
    keepStart = Selection.Start
    keepEnd = Selection.End
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False

    RemoveManagedCitationArtifacts targetDocument

    Set bibRange = FindZoteroBibliographyRange(targetDocument)
    If bibRange Is Nothing Then
        MsgBox "Zotero bibliography was not found. Please run Zotero -> Add/Edit Bibliography first.", vbExclamation
        GoTo CleanExit
    End If

    AddOrReplaceBookmark targetDocument, BIB_BOOKMARK, bibRange

    For Each aField In targetDocument.Fields
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
                    LinkNumericCitationField targetDocument, aField, titles, bibRange
                Else
                    LinkWholeCitationField targetDocument, aField, CStr(titles(1)), bibRange
                End If
            End If
        End If
    Next aField

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    targetDocument.Range(keepStart, keepEnd).Select
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    targetDocument.Range(keepStart, keepEnd).Select
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

Private Sub LinkNumericCitationField(ByVal targetDocument As Document, ByVal aField As Field, ByVal titles As Collection, ByVal bibRange As Range)
    Dim tokens As Collection
    Dim searchRange As Range
    Dim anchorRange As Range
    Dim targetTitle As String
    Dim bookmarkName As String
    Dim tooltipText As String
    Dim i As Long

    Set tokens = ExtractVisibleNumericTokens(aField.Result.Text)
    If tokens.Count = 0 Then
        LinkWholeCitationField targetDocument, aField, CStr(titles(1)), bibRange
        Exit Sub
    End If

    Set searchRange = aField.Result.Duplicate

    For i = 1 To tokens.Count
        targetTitle = ResolveTitleForTokenIndex(titles, tokens.Count, i)
        If Len(targetTitle) = 0 Then
            Exit For
        End If

        bookmarkName = EnsureBibliographyEntryBookmark(targetDocument, targetTitle, bibRange, tooltipText)
        If Len(bookmarkName) = 0 Then
            GoTo NextToken
        End If

        Set anchorRange = FindTextInRange(searchRange, CStr(tokens(i)))
        If anchorRange Is Nothing Then
            Set anchorRange = FindTextInRange(aField.Result.Duplicate, CStr(tokens(i)))
        End If

        If Not anchorRange Is Nothing Then
            AddHyperlinkToRange targetDocument, anchorRange, bookmarkName, tooltipText
            If anchorRange.End < aField.Result.End Then
                searchRange.Start = anchorRange.End
            End If
        End If

NextToken:
    Next i
End Sub

Private Sub LinkWholeCitationField(ByVal targetDocument As Document, ByVal aField As Field, ByVal title As String, ByVal bibRange As Range)
    Dim tooltipText As String
    Dim bookmarkName As String
    Dim anchorRange As Range
    Dim useWrappedTextStyle As Boolean

    bookmarkName = EnsureBibliographyEntryBookmark(targetDocument, title, bibRange, tooltipText)
    If Len(bookmarkName) = 0 Then
        Exit Sub
    End If

    useWrappedTextStyle = ShouldUseWrappedTextLinkRange(aField.Result.Text)
    Set anchorRange = ResolveWholeCitationAnchorRange(targetDocument, aField)
    AddHyperlinkToRange targetDocument, anchorRange, bookmarkName, tooltipText, useWrappedTextStyle
End Sub

Private Function ResolveWholeCitationAnchorRange(ByVal targetDocument As Document, ByVal aField As Field) As Range
    Dim candidateRange As Range

    Set candidateRange = aField.Result.Duplicate

    If ShouldUseWrappedTextLinkRange(candidateRange.Text) Then
        Set candidateRange = ExtractAuthorDateLinkRange(targetDocument, candidateRange)
    End If

    Set ResolveWholeCitationAnchorRange = candidateRange
End Function

Private Function ExtractAuthorDateLinkRange(ByVal targetDocument As Document, ByVal sourceRange As Range) As Range
    Dim resultRange As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim startChar As String
    Dim endChar As String

    Set resultRange = sourceRange.Duplicate
    startPos = resultRange.Start
    endPos = resultRange.End

    Do While startPos < endPos And IsWhitespaceCharacter(targetDocument.Range(startPos, startPos + 1).Text)
        startPos = startPos + 1
    Loop

    Do While endPos > startPos And IsWhitespaceCharacter(targetDocument.Range(endPos - 1, endPos).Text)
        endPos = endPos - 1
    Loop

    If endPos <= startPos Then
        Set ExtractAuthorDateLinkRange = sourceRange.Duplicate
        Exit Function
    End If

    startChar = targetDocument.Range(startPos, startPos + 1).Text
    endChar = targetDocument.Range(endPos - 1, endPos).Text

    If IsMatchingWrapper(startChar, endChar) And endPos - startPos > 2 Then
        startPos = startPos + 1
        endPos = endPos - 1
    End If

    Do While startPos < endPos And IsWhitespaceCharacter(targetDocument.Range(startPos, startPos + 1).Text)
        startPos = startPos + 1
    Loop

    Do While endPos > startPos And IsWhitespaceCharacter(targetDocument.Range(endPos - 1, endPos).Text)
        endPos = endPos - 1
    Loop

    If endPos <= startPos Then
        Set ExtractAuthorDateLinkRange = sourceRange.Duplicate
    Else
        Set ExtractAuthorDateLinkRange = targetDocument.Range(startPos, endPos)
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

Private Function FindZoteroBibliographyRange(ByVal targetDocument As Document) As Range
    Dim aField As Field

    For Each aField In targetDocument.Fields
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

Private Function EnsureBibliographyEntryBookmark(ByVal targetDocument As Document, ByVal title As String, ByVal bibRange As Range, ByRef tooltipText As String) As String
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
    AddOrReplaceBookmark targetDocument, bookmarkName, entryRange
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

Private Sub AddHyperlinkToRange(ByVal targetDocument As Document, ByVal anchorRange As Range, ByVal bookmarkName As String, ByVal tooltipText As String, Optional ByVal clearDirectFormatting As Boolean = False)
    Dim startPos As Long
    Dim linkText As String
    Dim newRange As Range
    Dim createdLink As Hyperlink
    Dim formatSnapshot As Variant
    Dim metadataKey As String
    Dim citationStyle As Style
    Dim stageText As String

    On Error GoTo LinkFail

    stageText = "capture snapshot"
    startPos = anchorRange.Start
    linkText = anchorRange.Text
    formatSnapshot = CaptureCharacterFormattingSnapshot(anchorRange)
    stageText = "ensure style"
    Set citationStyle = EnsureCitationLinkStyle(targetDocument, anchorRange)
    stageText = "save metadata"
    metadataKey = SaveFormattingSnapshot(targetDocument, formatSnapshot)

    stageText = "create hyperlink range"
    Set newRange = targetDocument.Range(startPos, startPos + Len(linkText))

    stageText = "insert hyperlink"
    Set createdLink = targetDocument.Hyperlinks.Add( _
        Anchor:=newRange, _
        Address:="", _
        SubAddress:=bookmarkName, _
        Target:=BuildLinkTarget(metadataKey), _
        ScreenTip:=tooltipText, _
        TextToDisplay:=linkText)

    stageText = "clear direct formatting"
    Set newRange = createdLink.Range.Duplicate
    ClearCharacterDirectFormattingRange newRange
    stageText = "apply citation style"
    ApplyCitationLinkStyle targetDocument, newRange, citationStyle
    Exit Sub

LinkFail:
    Err.Raise Err.Number, "ZoteroWordHyperlinks.AddHyperlinkToRange", "AddHyperlinkToRange failed during " & stageText & ": " & Err.Description
End Sub

Private Sub ClearCharacterDirectFormattingRange(ByVal targetRange As Range)
    targetRange.Select
    Selection.ClearCharacterDirectFormatting
End Sub

Private Function EnsureCitationLinkStyle(ByVal targetDocument As Document, ByVal sourceRange As Range) As Style
    Dim citationStyle As Style

    On Error GoTo StyleFail

    On Error Resume Next
    Set citationStyle = targetDocument.Styles(LINK_STYLE_NAME)
    On Error GoTo 0

    If citationStyle Is Nothing Then
        Set citationStyle = targetDocument.Styles.Add(Name:=LINK_STYLE_NAME, Type:=wdStyleTypeCharacter)
        CopyFontFormatting citationStyle.Font, sourceRange.Font
    End If

    Set EnsureCitationLinkStyle = citationStyle
    Exit Function

StyleFail:
    Err.Raise Err.Number, "ZoteroWordHyperlinks.EnsureCitationLinkStyle", "EnsureCitationLinkStyle failed: " & Err.Description
End Function

Private Sub ApplyCitationLinkStyle(ByVal targetDocument As Document, ByVal targetRange As Range, ByVal citationStyle As Style)
    On Error Resume Next
    targetRange.Style = citationStyle
    On Error GoTo 0
End Sub

Private Sub ClearCitationLinkCharacterStyle(ByVal targetDocument As Document, ByVal targetRange As Range)
    On Error Resume Next
    targetRange.Style = targetDocument.Styles(wdStyleDefaultParagraphFont)
    On Error GoTo 0
End Sub

Private Function SaveFormattingSnapshot(ByVal targetDocument As Document, ByVal snapshot As Variant) As String
    Dim metadataKey As String

    metadataKey = GetNextFormattingMetadataKey(targetDocument)
    SetDocumentVariable targetDocument, metadataKey, SerializeCharacterFormattingSnapshot(snapshot)
    SaveFormattingSnapshot = metadataKey
End Function

Private Function GetNextFormattingMetadataKey(ByVal targetDocument As Document) As String
    Dim nextId As Long

    nextId = GetNextDocumentCounterValue(targetDocument, NEXT_FORMAT_ID_VARIABLE)
    SetDocumentVariable targetDocument, NEXT_FORMAT_ID_VARIABLE, CStr(nextId)
    GetNextFormattingMetadataKey = FORMAT_VARIABLE_PREFIX & Format$(nextId, "0000000000")
End Function

Private Function GetNextDocumentCounterValue(ByVal targetDocument As Document, ByVal variableName As String) As Long
    Dim currentValue As Long

    On Error GoTo MissingValue
    currentValue = CLng(targetDocument.Variables(variableName).Value)
    GetNextDocumentCounterValue = currentValue + 1
    Exit Function

MissingValue:
    GetNextDocumentCounterValue = 1
End Function

Private Sub SetDocumentVariable(ByVal targetDocument As Document, ByVal variableName As String, ByVal variableValue As String)
    On Error Resume Next
    targetDocument.Variables(variableName).Value = variableValue
    If Err.Number <> 0 Then
        Err.Clear
        targetDocument.Variables.Add Name:=variableName, Value:=variableValue
    End If
    On Error GoTo 0
End Sub

Private Sub DeleteManagedFormattingVariables(ByVal targetDocument As Document)
    Dim i As Long
    Dim variableName As String

    For i = targetDocument.Variables.Count To 1 Step -1
        variableName = targetDocument.Variables(i).Name
        If Left$(variableName, Len(FORMAT_VARIABLE_PREFIX)) = FORMAT_VARIABLE_PREFIX Then
            targetDocument.Variables(i).Delete
        End If
    Next i
End Sub

Private Function SerializeCharacterFormattingSnapshot(ByVal snapshot As Variant) As String
    Dim charCount As Long
    Dim propertyCount As Long
    Dim i As Long
    Dim j As Long
    Dim recordText As String
    Dim resultText As String

    On Error GoTo SnapshotMissing
    charCount = UBound(snapshot, 1)
    propertyCount = UBound(snapshot, 2)

    For i = 1 To charCount
        recordText = ""
        For j = 1 To propertyCount
            If j > 1 Then
                recordText = recordText & "|"
            End If
            recordText = recordText & EscapeSnapshotValue(CStr(snapshot(i, j)))
        Next j

        If i > 1 Then
            resultText = resultText & "~"
        End If
        resultText = resultText & recordText
    Next i

    SerializeCharacterFormattingSnapshot = resultText
    Exit Function

SnapshotMissing:
    SerializeCharacterFormattingSnapshot = ""
End Function

Private Function TryLoadFormattingSnapshot(ByVal targetDocument As Document, ByVal targetValue As String, ByRef snapshot As Variant) As Boolean
    Dim metadataKey As String
    Dim serializedSnapshot As String

    If Not TryParseFormatMetadataKey(targetValue, metadataKey) Then
        Exit Function
    End If

    On Error GoTo MissingValue
    serializedSnapshot = targetDocument.Variables(metadataKey).Value
    If Len(serializedSnapshot) = 0 Then
        Exit Function
    End If

    snapshot = DeserializeCharacterFormattingSnapshot(serializedSnapshot)
    TryLoadFormattingSnapshot = True
    Exit Function

MissingValue:
    TryLoadFormattingSnapshot = False
End Function

Private Function DeserializeCharacterFormattingSnapshot(ByVal serializedSnapshot As String) As Variant
    Dim records() As String
    Dim values() As String
    Dim snapshot() As Variant
    Dim i As Long
    Dim j As Long

    records = Split(serializedSnapshot, "~")
    ReDim snapshot(1 To UBound(records) - LBound(records) + 1, 1 To 17)

    For i = LBound(records) To UBound(records)
        values = Split(records(i), "|")
        For j = 0 To 16
            If j <= UBound(values) Then
                snapshot(i - LBound(records) + 1, j + 1) = UnescapeSnapshotValue(values(j))
            Else
                snapshot(i - LBound(records) + 1, j + 1) = ""
            End If
        Next j
    Next i

    DeserializeCharacterFormattingSnapshot = snapshot
End Function

Private Function EscapeSnapshotValue(ByVal rawValue As String) As String
    rawValue = Replace(rawValue, "%", "%25")
    rawValue = Replace(rawValue, "|", "%7C")
    rawValue = Replace(rawValue, "~", "%7E")
    rawValue = Replace(rawValue, vbCr, "%0D")
    rawValue = Replace(rawValue, vbLf, "%0A")
    EscapeSnapshotValue = rawValue
End Function

Private Function UnescapeSnapshotValue(ByVal rawValue As String) As String
    rawValue = Replace(rawValue, "%0A", vbLf)
    rawValue = Replace(rawValue, "%0D", vbCr)
    rawValue = Replace(rawValue, "%7E", "~")
    rawValue = Replace(rawValue, "%7C", "|")
    rawValue = Replace(rawValue, "%25", "%")
    UnescapeSnapshotValue = rawValue
End Function

Private Sub ApplyLinkedCitationAppearance(ByVal targetRange As Range)
    Dim i As Long
    Dim targetColor As Long

    targetColor = GetConfiguredLinkColor()

    targetRange.Font.Color = targetColor
    targetRange.Font.Underline = wdUnderlineNone

    For i = 1 To targetRange.Characters.Count
        With targetRange.Characters(i).Font
            .Color = targetColor
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

Private Sub RemoveHyperlinkSafely(ByVal targetDocument As Document, ByVal hl As Hyperlink)
    Dim displayText As String
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim targetStart As Long
    Dim hadToRewrite As Boolean
    Dim storedColor As Long
    Dim hasStoredColor As Boolean
    Dim formatSnapshot As Variant
    Dim hasFormatSnapshot As Boolean

    Set sourceRange = hl.Range.Duplicate
    Set targetRange = hl.Range.Duplicate
    targetStart = targetRange.Start
    displayText = targetRange.Text
    hasFormatSnapshot = TryLoadFormattingSnapshot(targetDocument, hl.Target, formatSnapshot)
    hasStoredColor = TryParseStoredColor(hl.Target, storedColor)

    On Error Resume Next
    hl.Delete
    Err.Clear
    On Error GoTo 0

    If targetRange.Hyperlinks.Count > 0 Then
        hadToRewrite = True
        targetRange.Text = displayText
    End If

    Set targetRange = targetDocument.Range(targetStart, targetStart + Len(displayText))
    If hasFormatSnapshot Then
        ClearCitationLinkCharacterStyle targetDocument, targetRange
        RestoreCharacterFormattingFromSnapshot targetRange, formatSnapshot
    Else
        If hadToRewrite Then
            RestoreCharacterFormatting targetRange, sourceRange
        End If
        ApplyUnlinkedCitationAppearance targetRange, hasStoredColor, storedColor
    End If
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

Private Function BuildLinkTarget(ByVal metadataKey As String) As String
    BuildLinkTarget = FORMAT_TARGET_PREFIX & metadataKey
End Function

Private Function TryParseFormatMetadataKey(ByVal targetValue As String, ByRef metadataKey As String) As Boolean
    If Left$(targetValue, Len(FORMAT_TARGET_PREFIX)) <> FORMAT_TARGET_PREFIX Then
        Exit Function
    End If

    metadataKey = Mid$(targetValue, Len(FORMAT_TARGET_PREFIX) + 1)
    If Len(metadataKey) = 0 Then
        Exit Function
    End If

    TryParseFormatMetadataKey = True
End Function

Private Function TryParseStoredColor(ByVal targetValue As String, ByRef storedColor As Long) As Boolean
    Dim rawValue As String

    If Left$(targetValue, Len(LEGACY_LINK_TARGET_PREFIX)) <> LEGACY_LINK_TARGET_PREFIX Then
        Exit Function
    End If

    rawValue = Mid$(targetValue, Len(LEGACY_LINK_TARGET_PREFIX) + 1)
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
        .Color = sourceFont.Color
        .Underline = sourceFont.Underline
    End With
End Sub

Private Function CaptureCharacterFormattingSnapshot(ByVal sourceRange As Range) As Variant
    Dim snapshot() As Variant
    Dim charCount As Long
    Dim i As Long

    charCount = sourceRange.Characters.Count
    ReDim snapshot(1 To charCount, 1 To 17)

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
            snapshot(i, 16) = .Color
            snapshot(i, 17) = .Underline
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
            .Color = snapshot(i, 16)
            .Underline = snapshot(i, 17)
        End With
    Next i
    Exit Sub

SnapshotMissing:
    Err.Clear
End Sub

Private Sub AddOrReplaceBookmark(ByVal targetDocument As Document, ByVal bookmarkName As String, ByVal bookmarkRange As Range)
    If targetDocument.Bookmarks.Exists(bookmarkName) Then
        targetDocument.Bookmarks(bookmarkName).Delete
    End If
    targetDocument.Bookmarks.Add Name:=bookmarkName, Range:=bookmarkRange
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
