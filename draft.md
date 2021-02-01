# Define styles for a new Word&nbsp;template: a short example

To make a new Word template means defining many styles. The easiest way I've found is to write a description of the styles, then run a macro that defines styles like the description.

For example, suppose you want a template that matches the format of an ANSI standard, "[Scientific and Technical Reports&mdash;Preparation, Presentation, and Preservation](https://www.niso.org/publications/z39.18-2005-r2010)." Measure its margins, compare the font sizes to some samples, and write a description. Use the format shown below, with a style name at the beginning of the line and commas between the specifications. 

## Sample style descriptions

*General settings . . . . . . . . . . . .*

Margins, 1&quot; left, 1.25&quot; right, 1&quot; top, 1&quot; bottom, 0.5&quot; header, 0.5&quot; footer, mirror margins.

Defaults for all defined styles, Body font, auto color, not italic, not bold, 10 pt size, 1.04 line spacing, 0 pt before, 0 pt after, 0&quot; left indent, 0&quot; right indent, left aligned, based on no style, followed by Body Text, normal character spacing, no kerning.

*Headings . . . . . . . . . . . .*

Heading 1 style, &quot;%1&quot; number, Headings font, bold, black color, 14 pt size, 24 pt before, 18 pt after, 2.25 pt bottom border, keep with next.

Heading 2 style, &quot;%1.%2&quot; number, Headings font, bold, black color, 12 pt size, 24 pt before, 12 pt after, 0.5 top border, keep with next.

Heading 3 style, &quot;%1.%2.%3&quot; number, Headings font, bold, black color, 11 pt size, 18 pt before, 6 pt after, keep with next.

Heading 4 style, &quot;%1.%2.%3.%4&quot; number, Headings font, bold, black color, 10 pt size, 18 pt before, 6 pt after, keep with next.

ListHeadings list, Heading 1, Heading 2, Heading 3, Heading 4.

ListHeadings number defaults, Headings number font, auto number color, 0&quot; number indent, 0.5&quot; text indent, tab after number.

*Other styles . . . . . . . . . . . .*

Body Text style, 12 pt after, 0.5&quot; left indent.

Header style, Headings font, bold, 11 pt size, no tabs.

Footer style, Headings font, 9 pt size, 0.13&quot; right indent, clear tabs, right tab.

Caption style, Headings font, bold, 6 pt after, centered.

Table Text style, 8.5 pt size, single line spacing.

List Bullet style, &#8226; bullet, 0.5&quot; bullet indent, 0.75&quot; text indent, 6 pt after, space between, followed by List Bullet.

ListBullets list, List Bullet.

ListBullets bullet defaults, Body bullet font, tab after bullet.

## Apply the style descriptions

To apply style descriptions, open a new Word document, set the style defaults, paste the style descriptions, paste and run the style macro, and save the Word document.

###  Open a document without styles

1. Type **winword /a /w** in the Windows taskbar and press **Enter**. \
<span style='font-size:small; color:darkgray;'>&#128712; The /a switch opens Word without opening your Normal template, which might have custom styles. The /w switch opens a new blank document. For more info see [Command-line switches for Microsoft Office products](https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6).</span>

### Set the style defaults

1. In Word, click the **Design** menu and click **Fonts**.
1. Select the theme fonts. For these sample styles, click **Arial**.
1. Click the <img src='far/lightbulb.svg' alt='light bulb' height='12'>&ensp;**Tell Me** box, type **styles**, and press **Enter**.
1. Click **Manage Styles**, the third button in the Styles pane.
1. Click the **Set Defaults** tab in the Manage Styles window. 
    1. Select a font size. For these sample styles, select **10**. 
    1. For the paragraph spacing after, select **0 pt**. 
    1. Select the line spacing. For these sample styles, leave **Multiple** and type **1.04**.
    1. Click **OK**. 

### Add the style descriptions and macro

1. Copy the style descriptions (see above).
1. Right-click the Word document and select the paste option **Keep Text Only**.
1. Click the <img src='far/lightbulb.svg' alt='light bulb' height='12'>&ensp;**Tell Me** box, type **visual basic**, and press **Enter**.
    1. In the Visual Basic window, click **Insert** and **Module**.
    1. Copy the macro (see below) and paste in Visual Basic.
    1. Click **File** and **Close and Return to Microsoft Word**.
1. In Word, click the **View** menu and click **Macros**.
1. Select the macro **sctApplySpecs** and click **Run**.

### Save the file

1. Click **File** and **Save As**.
1. Click **Browse**.
1. Select a folder.
1. Type a file name. For the sample styles, type **Sample standard styles**.
1. Select a file type. \
&bull; To start a document, leave **Word Document (\*.docx)** as the file type. \
&bull; To start a template, leave **Word Template (\*.dotx)**. \
&bull; To make further style changes with the macro, select **Word Macro-Enabled Document (\*.docm)**.
1. Click **Save.**

## Macro to read the style descriptions

```vba
Const sctSpecs As Long = 20
Const strDefaultStyleGallery As String = "Normal, No Spacing, Heading 1, " _
    & "Heading 2, Heading 3, Heading 4, Heading 5, Heading 6, Heading 7, " _
    & "Heading 8, Heading 9, Title, Subtitle, Subtle Emphasis, Emphasis, " _
    & "Intense Emphasis, Strong, Quote, Intense Quote, Subtle Reference, " _
    & "Intense Reference, Book Title, List Paragraph, Caption, TOC Heading"

Sub sctApplySpecs()
    Dim rngParas As Range, arrParas() As String
    Dim strPara As String, lngPara As Long, lngListPara As Long
    Dim strLabel As String, strLabelLow As String
    Dim arrSpecs() As String, strSpec As String
    Dim strSpecLow As String, lngSpec As Long, dblSpec As Double
    Dim arrStyles() As String, lngStyles As Long, strStyle As String
    Dim arrDefaultStyleGallery() As String
    Dim arrList() As String, strList As String, lngList As Long
    Dim lngLevel As Long, lngLevels As Long
    Dim objListTemplate As ListTemplate
    
    'Reads each line in the document.
    Set rngParas = ActiveDocument.StoryRanges(wdMainTextStory)
    arrParas = sctSaveParagraphsInAnArray(rngParas)
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = arrParas(lngPara)
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strPara, ", ")
        'Saves the first specification on each line, such as "Body Text style."
        strLabel = arrSpecs(0)
        strLabelLow = LCase(strLabel)
'Margins'
'-------'Sets the margins.
        If strLabelLow = "margins" Then
            For lngSpec = 1 To UBound(arrSpecs)
                strSpec = arrSpecs(lngSpec)
                strSpecLow = LCase(strSpec)
                dblSpec = Val(strSpec)
                With ActiveDocument.PageSetup
                    If InStr(strSpecLow, "left") <> 0 Then
                        .LeftMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "right") <> 0 Then
                        .RightMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "top") <> 0 Then
                        .TopMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "bottom") <> 0 Then
                        .BottomMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "header") <> 0 Then
                        .HeaderDistance = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "footer") <> 0 Then
                        .FooterDistance = InchesToPoints(dblSpec)
                    ElseIf strSpecLow = "mirror margins" Then
                        .MirrorMargins = True
                    ElseIf strSpecLow = "no mirror margins" Then
                        .MirrorMargins = False
                    End If
                End With
            Next lngSpec
'Styles '
'-------'Saves the style names in an array.
        ElseIf Right(strLabelLow, 5) = "style" _
            And Right(strLabelLow, 11) <> " list style" Then
            strStyle = Left(strLabel, InStr(strLabelLow, " style") - 1)
            lngStyles = lngStyles + 1
            If lngStyles = 1 Then
                ReDim arrStyles(1 To 1)
            Else
                ReDim Preserve arrStyles(1 To lngStyles)
            End If
            arrStyles(lngStyles) = strStyle
            
            'Adds a style if it doesn't exist.
            If Not sctStyleExists(strStyle, ActiveDocument) Then
                If InStr(strPara, ", character style") <> 0 Then
                    dblSpec = wdStyleTypeCharacter
                Else
                    dblSpec = wdStyleTypeParagraph
                End If
                ActiveDocument.Styles.Add strStyle, dblSpec
            End If
        End If
    Next lngPara
    
    'Reads each line in the document (again).
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = arrParas(lngPara)
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strPara, ", ")
        'Saves the first specification on each line, such as "Body Text style."
        strLabel = arrSpecs(0)
        strLabelLow = LCase(strLabel)
        
        'If the line begins "Defaults for all defined styles," then...
        If strLabelLow = "defaults for all defined styles" _
            Or strLabelLow = "defaults for defined styles" _
            Or strLabelLow = "default for all defined styles" _
            Or strLabelLow = "default for defined styles" Then
            'Applies the default specifications to all defined styles.
            For lngSpec = LBound(arrStyles) To UBound(arrStyles)
                strStyle = arrStyles(lngSpec)
                'Sends a style name and specs to the sctDefineStyle macro.
                sctDefineStyle strStyle, arrSpecs
            Next lngSpec
        
        'Or if the line begins with a style name, then...
        ElseIf Right(strLabelLow, 5) = "style" Then
            'Applies the specifications on the line to the style.
            strStyle = Left(strLabel, InStr(strLabelLow, "style") - 2)
            sctDefineStyle strStyle, arrSpecs
'Gallery'
'-------'Customizes the quick styles gallery.
        ElseIf strLabelLow = "styles gallery" Then
            'Removes the defaults.
            arrDefaultStyleGallery = Split(strDefaultStyleGallery, ", ")
            For lngSpec = LBound(arrDefaultStyleGallery) _
                To UBound(arrDefaultStyleGallery)
                strStyle = arrDefaultStyleGallery(lngSpec)
                ActiveDocument.Styles(strStyle).QuickStyle = False
            Next lngSpec
            'Adds styles to the Style gallery.
            For lngSpec = 1 To UBound(arrSpecs)
                strStyle = arrSpecs(lngSpec)
                With ActiveDocument.Styles(strStyle)
                    .QuickStyle = True ' True means include in the gallery.
                    .UnhideWhenUsed = False ' False means never hidden.
                    .Visibility = False ' False (sic) means always visible.
                    .Priority = lngSpec
                End With
            Next lngSpec
'Lists  '
'-------'Or if the line begins with a list template name, then...
        ElseIf Right(strLabelLow, 5) = " list" _
            Or Right(strLabelLow, 11) = " list style" _
            Or Right(strLabelLow, 12) = " list styles" Then
            
            'Saves the list name.
            If Right(strLabelLow, 4) = "list" _
                And Left(strLabel, InStr(strLabelLow, " multi") - 1) <> 0 Then
                strList = Left(strLabel, InStr(strLabelLow, " multi") - 1)
            Else
                strList = Left(strLabel, InStr(strLabelLow, " list") - 1)
            End If
            'Adds a list template if it doesn't exist.
            If Not sctStyleExists(strList, ActiveDocument) Then
                ActiveDocument.ListTemplates.Add True, CStr(strList)
            End If
            'Counts the styles in the list.
            lngLevels = UBound(arrSpecs)
            If lngLevels > 9 Then lngLevels = 9
            
            'Saves the style names in an array (_, 1).
            Erase arrList
            ReDim arrList(1 To lngLevels, 1 To sctSpecs)
            For lngLevel = 1 To lngLevels
                arrList(lngLevel, 1) = arrSpecs(lngLevel)
            Next lngLevel
            
            'Reads each line in the document, looking for specs for the list.
            For lngListPara = LBound(arrParas) To UBound(arrParas)
                strPara = arrParas(lngListPara)
                'Saves the specifications on each line (between commas).
                arrSpecs = Split(strPara, ", ")
                'Saves the first specification on each line.
                strLabel = arrSpecs(0)
                strLabelLow = LCase(strLabel)
                
                'If a line has defaults for the list, saves them in the array.
                If Left(strLabel, Len(strList)) = strList _
                    And Right(strLabelLow, 9) = " defaults" Then
                    For lngLevel = 1 To lngLevels
                        sctDefineList arrList, lngLevel, arrSpecs
                    Next lngLevel
                
                'If a line has specs for a style...
                ElseIf Right(strLabelLow, 5) = "style" Then
                    strStyle = Left(strLabel, InStr(strLabelLow, " style") - 1)
                    '...and if the style is in the list...
                    For lngLevel = 1 To lngLevels
                        If arrList(lngLevel, 1) = strStyle Then
                            '...saves the specs in the array.
                            sctDefineList arrList, lngLevel, arrSpecs
                        End If
                    Next lngLevel
                End If
            Next lngListPara
            
            'Adds a list template if it doesn't exist.
            If sctStyleExists(strList, ActiveDocument) Then
                Set objListTemplate = ActiveDocument.ListTemplates(strList)
            Else
                Set objListTemplate = _
                    ActiveDocument.ListTemplates.Add(True, CStr(strList))
            End If
            
            'Applies the list template specifications.
            For lngLevel = 1 To lngLevels
                With objListTemplate.ListLevels(lngLevel)
                    If arrList(lngLevel, 2) <> "" Then
                        .NumberFormat = arrList(lngLevel, 2)
                    End If
                    With .Font
                        If arrList(lngLevel, 11) <> "" Then
                            .Name = arrList(lngLevel, 11)
                        End If
                        If arrList(lngLevel, 12) <> "" Then
                            .Bold = arrList(lngLevel, 12)
                        End If
                        If arrList(lngLevel, 13) <> "" Then
                            .Italic = arrList(lngLevel, 13)
                        End If
                        If arrList(lngLevel, 14) <> "" Then
                            .Color = arrList(lngLevel, 14)
                        End If
                    End With
                    If arrList(lngLevel, 3) <> "" Then
                        .TrailingCharacter = arrList(lngLevel, 3)
                    End If
                    If arrList(lngLevel, 4) <> "" Then
                        .NumberStyle = arrList(lngLevel, 4)
                    End If
                    If arrList(lngLevel, 5) <> "" Then
                        .NumberPosition = arrList(lngLevel, 5)
                    End If
                    If arrList(lngLevel, 6) <> "" Then
                        .Alignment = arrList(lngLevel, 6)
                    End If
                    If arrList(lngLevel, 7) <> "" Then
                        .TextPosition = arrList(lngLevel, 7)
                    End If
'                    .TabPosition = wdUndefined
'                    .ResetOnHigher = (lngLevel - 1)
'                    .StartAt = 1
                    If arrList(lngLevel, 1) <> "" Then
                        .LinkedStyle = arrList(lngLevel, 1)
                    End If
                    'The linked style name must be set after the indents.
                End With
            Next lngLevel
            Set objListTemplate = Nothing
        End If
    Next lngPara
    strSpec = "Styles defined." & vbCrLf & vbCrLf & "Insert sample text?"
    dblSpec = MsgBox(strSpec, vbYesNo, "Macro complete")
    If dblSpec = vbYes Then
        rngParas.Select
        With Selection
            .Collapse wdCollapseEnd
            .TypeParagraph
            .ClearFormatting
        End With
        sctInsertSampleText arrStyles
    End If
End Sub

Private Function sctSaveParagraphsInAnArray(ByVal rngRange As Range) As String()
    Dim arrParas() As String, lngPara As Long, strPara As String
    'Saves paragraphs in an array.
    arrParas = Split(rngRange, vbCr)
    'Cleans up the text in the paragraphs.
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = CStr(arrParas(lngPara))
        'Removes spaces and a period at the end of a paragraph.
        strPara = Trim(strPara)
        If Right(strPara, 1) = "." Then
            strPara = Left(strPara, Len(strPara) - 1)
        End If
        'Replaces manual line breaks with a space.
        strPara = Replace(strPara, Chr(11), " ")
        'Removes extra spaces.
        Do While InStr(strPara, "  ") <> 0
            strPara = Replace(strPara, "  ", " ")
        Loop
        'Saves some text instead of an empty line.
        If strPara = "" Then strPara = "[empty line]"
        arrParas(lngPara) = strPara
    Next lngPara
    sctSaveParagraphsInAnArray = arrParas
End Function

Private Function sctStyleExists(ByVal strStyle As String, _
    ByVal objDocument As Document) As Boolean
    Dim objStyle As Style, objListTemplate As ListTemplate
    On Error Resume Next
    Set objStyle = objDocument.Styles(strStyle)
    sctStyleExists = Not objStyle Is Nothing
    If Not sctStyleExists Then
        Set objListTemplate = objDocument.ListTemplates(strStyle)
        sctStyleExists = Not objListTemplate Is Nothing
    End If
End Function

Private Sub sctInsertSampleText(arrStyles() As String)
    Dim lngStyle As Long
    For lngStyle = LBound(arrStyles) To UBound(arrStyles)
        With Selection
            .InsertAfter arrStyles(lngStyle) & " sample" & vbCrLf
            .Paragraphs(1).Style = arrStyles(lngStyle)
            .Collapse wdCollapseEnd
        End With
     Next lngStyle
End Sub

Private Sub sctDefineStyle(ByVal strStyle As String, arrSpecs() As String)
    
    Dim lngType As Long, lngSpec As Long, strSpec As String, dblSpec As Double
    Dim strSpecLow As String, dblSpec2 As Double
    
    lngType = ActiveDocument.Styles(strStyle).Type
    
    'Looks at each specification in the array.
    For lngSpec = 1 To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
        
        With ActiveDocument.Styles(strStyle)
            If Left(strSpecLow, 8) = "based on" Then '----------- based on style
                strSpec = Right(strSpec, Len(strSpec) - 9)
                strSpecLow = LCase(strSpec)
                If strSpecLow = "no style" Then
                    .BaseStyle = ""
                ElseIf strStyle <> "Normal" _
                    And strStyle <> "Default Paragraph Font" Then
                    .BaseStyle = strSpec
                End If
            ElseIf Left(strSpecLow, 11) = "followed by" Then '-- following style
                strSpec = Right(strSpec, Len(strSpec) - 12)
                strSpecLow = LCase(strSpec)
                If Right(strSpecLow, 6) = " style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 6)
                End If
                .NextParagraphStyle = strSpec
            ElseIf strSpecLow = "space between" _
                Or strSpecLow = "add space between" _
                Or strSpecLow = "space between paragraphs of the same style" _
                Or strSpecLow = "add space between paragraphs of the same" _
                & " style" Then '--------------------------------- space between
                .NoSpaceBetweenParagraphsOfSameStyle = False
            ElseIf strSpecLow = "no space between" _
                Or strSpecLow = "no space between paragraphs of the same style" _
                Or strSpecLow = "don't add space between paragraphs" _
                Or strSpecLow = "don't add space between paragraphs of the" _
                & " same style" Then
                .NoSpaceBetweenParagraphsOfSameStyle = True
            End If
        End With
        
        With ActiveDocument.Styles(strStyle).Font
            If Right(strSpecLow, 4) = "font" Then '------------------------ font
                strSpec = Left(strSpec, Len(strSpec) - 5)
                strSpecLow = LCase(strSpec)
                If strSpecLow = "body" Then
                    strSpec = "+Body"
                ElseIf strSpecLow = "headings" _
                    Or strSpecLow = "heading" Then
                    strSpec = "+Headings"
                ElseIf strSpecLow = "default" Then
                    strSpec = ""
                End If
                .Name = strSpec
            ElseIf Right(strSpecLow, 4) = "size" Then '-------------------- size
                .Size = Val(strSpec)
            ElseIf strSpecLow = "bold" Then '------------------------------ bold
                .Bold = True
            ElseIf strSpecLow = "not bold" Or strSpecLow = "no bold" Then
                .Bold = False
            ElseIf strSpecLow = "italic" Then '-------------------------- italic
                .Italic = True
            ElseIf strSpecLow = "not italic" Or strSpecLow = "no italic" Then
                .Italic = False
            ElseIf strSpecLow = "small caps" Then '------------------ small caps
                .SmallCaps = False
            ElseIf strSpecLow = "uppercase" Or strSpecLow = "all caps" _
                Then '----------------------------------------------------- caps
                .AllCaps = False
            ElseIf Right(strSpecLow, 5) = "color" Then '------------------ color
                strSpec = Split(strSpec, " ")(0)
                strSpecLow = LCase(strSpec)
                If Left(strSpec, 1) = "#" Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                    strSpec = Right(strSpec, 2) & Mid(strSpec, 3, 2) _
                        & Left(strSpec, 2)
                    dblSpec = Val("&H" & strSpec)
                    .Color = dblSpec
                ElseIf strSpecLow = "automatic" Or strSpecLow = "auto" _
                    Or strSpecLow = "no" Then
                    dblSpec = wdColorAutomatic
                    .Color = dblSpec
                ElseIf strSpecLow = "black" Then
                    dblSpec = wdColorBlack
                    .Color = dblSpec
                End If
            ElseIf strSpecLow = "normal character spacing" Then ' letter spacing
                .Spacing = 0
            ElseIf strSpecLow = "kerning" Then '------------------------ kerning
                .Kerning = 8
            ElseIf strSpecLow = "no kerning" Then
                .Kerning = 0
            End If
        End With
        
        If lngType = wdStyleTypeParagraph Then
            With ActiveDocument.Styles(strStyle).ParagraphFormat
                If Right(strSpecLow, 11) = "left indent" Then '--------- indents
                    .LeftIndent = InchesToPoints(dblSpec)
                ElseIf Right(strSpecLow, 12) = "right indent" Then
                    .RightIndent = InchesToPoints(dblSpec)
                ElseIf Right(strSpecLow, 6) = "before" _
                    And strSpecLow <> "page break before" _
                    And strSpecLow <> "no page break before" Then ' space before
                    .SpaceBefore = dblSpec
                ElseIf Right(strSpecLow, 5) = "after" Then '-------- space after
                    .SpaceAfter = dblSpec
                ElseIf Right(strSpecLow, 12) = "line spacing" Then 'line spacing
                    If Split(strSpecLow, " ")(1) = "pt" _
                        Or Split(strSpecLow, " ")(1) = "pt." Then
                        .LineSpacingRule = wdLineSpaceExactly
                        .LineSpacing = dblSpec
                    ElseIf Split(strSpecLow, " ")(0) = "exact" _
                        Or Split(strSpecLow, " ")(0) = "exactly" Then
                        dblSpec = Val(Split(strSpec, " ")(1))
                        .LineSpacingRule = wdLineSpaceExactly
                        .LineSpacing = dblSpec
                    ElseIf Split(strSpecLow, " ")(1) = "least" Then
                        dblSpec = Val(Split(strSpec, " ")(2))
                        .LineSpacingRule = wdLineSpaceAtLeast
                        .LineSpacing = dblSpec
                    ElseIf Split(strSpecLow, " ")(0) = "single" Then
                        .LineSpacingRule = wdLineSpaceSingle
                    Else
                        .LineSpacingRule = wdLineSpaceMultiple
                        .LineSpacing = LinesToPoints(dblSpec)
                    End If
                ElseIf strSpecLow = "left aligned" _
                    Or strSpecLow = "right aligned" _
                    Or strSpecLow = "centered" Or strSpecLow = "center" _
                    Or strSpecLow = "center align" _
                    Or strSpecLow = "justified" Or strSpecLow = "justify" _
                    Then '---------------------------------- --------- alignment
                    dblSpec = wdAlignParagraphLeft
                    If strSpecLow = "right aligned" Then
                        dblSpec = wdAlignParagraphRight
                    ElseIf strSpecLow = "centered" Or strSpecLow = "center" _
                        Or strSpecLow = "center align" Then
                        dblSpec = wdAlignParagraphCenter
                    ElseIf strSpecLow = "justified" Or strSpecLow = "justify" _
                        Then
                        dblSpec = wdAlignParagraphJustify
                    End If
                    .Alignment = dblSpec
                ElseIf strSpecLow = "widow/orphan control" _
                    Or strSpecLow = "orphan/widow control" _
                    Or strSpecLow = "widow and orphan control" _
                    Or strSpecLow = "orphan and widow control" _
                    Or strSpecLow = "widow control" _
                    Or strSpecLow = "orphan control" Then '------- widow control
                    .WidowControl = True
                ElseIf strSpecLow = "no widow/orphan control" _
                    Or strSpecLow = "no orphan/widow control" _
                    Or strSpecLow = "no widow and orphan control" _
                    Or strSpecLow = "no orphan and widow control" _
                    Or strSpecLow = "no widow or orphan control" _
                    Or strSpecLow = "no orphan or widow control" _
                    Or strSpecLow = "no widow control" _
                    Or strSpecLow = "no orphan control" Then
                    .WidowControl = False
                ElseIf strSpecLow = "keep with next" Then '------ keep with next
                    .KeepWithNext = True
                ElseIf strSpecLow = "don't keep with next" _
                    Or strSpecLow = "do not keep with next" Then
                    .KeepWithNext = False
                ElseIf strSpecLow = "keep lines together" _
                    Or strSpecLow = "keep together" Then '-- keep lines together
                    .KeepTogether = True
                ElseIf strSpecLow = "don't keep together" _
                    Or strSpecLow = "don't keep lines together" _
                    Or strSpecLow = "do not keep together" _
                    Or strSpecLow = "do not keep lines together" Then
                    .KeepTogether = False
                ElseIf strSpecLow = "page break before" Then ' page break before
                    .PageBreakBefore = True
                ElseIf strSpecLow = "no page break before" Then
                    .PageBreakBefore = False
                ElseIf Right(strSpecLow, 6) = "border" Then '----------- borders
                    If InStr(strSpecLow, "top") <> 0 Then
                        dblSpec2 = wdBorderTop
                    ElseIf InStr(strSpecLow, "bottom") <> 0 Then
                        dblSpec2 = wdBorderBottom
                    ElseIf InStr(strSpecLow, "left") <> 0 Then
                        dblSpec2 = wdBorderLeft
                    ElseIf InStr(strSpecLow, "right") <> 0 Then
                        dblSpec2 = wdBorderRight
                    End If
                    If dblSpec2 <> 0 Then
                        With .Borders(dblSpec2)
                            .LineStyle = wdLineStyleSingle
                            Select Case dblSpec
                                Case 0
                                    .LineWidth = wdUndefined
                                Case Is <= 0.25
                                    .LineWidth = wdLineWidth025pt
                                Case Is <= 0.5
                                    .LineWidth = wdLineWidth050pt
                                Case Is <= 0.75
                                    .LineWidth = wdLineWidth075pt
                                Case Is <= 1
                                    .LineWidth = wdLineWidth100pt
                                Case Is <= 1.5
                                    .LineWidth = wdLineWidth150pt
                                Case Is <= 2.25
                                    .LineWidth = wdLineWidth225pt
                                Case Is <= 3
                                    .LineWidth = wdLineWidth300pt
                                Case Is <= 4.5
                                    .LineWidth = wdLineWidth450pt
                                Case Is > 4.5
                                    .LineWidth = wdLineWidth600pt
                            End Select
                        End With
                    End If
                ElseIf strSpecLow = "no tabs" _
                    Or strSpecLow = "clear tabs" Then '-------------------- tabs
                    .TabStops.ClearAll
                ElseIf strSpecLow = "center tab" _
                    Or strSpecLow = "centered tab" Then
                    dblSpec = ActiveDocument.PageSetup.PageWidth _
                        - ActiveDocument.PageSetup.LeftMargin _
                        - ActiveDocument.PageSetup.RightMargin _
                        - .LeftIndent - .RightIndent
                    .TabStops.Add Position:=(dblSpec / 2), _
                        Alignment:=wdAlignTabCenter, _
                        Leader:=wdTabLeaderSpaces
                ElseIf strSpecLow = "right tab" Then
                    dblSpec = ActiveDocument.PageSetup.PageWidth _
                        - ActiveDocument.PageSetup.LeftMargin _
                        - ActiveDocument.PageSetup.RightMargin _
                        - .LeftIndent - .RightIndent
                    .TabStops.Add Position:=dblSpec, _
                        Alignment:=wdAlignTabRight, _
                        Leader:=wdTabLeaderSpaces
                End If
            End With
        End If
    Next lngSpec
End Sub

Private Sub sctDefineList(ByRef arrList() As String, ByVal lngLevel As Long, _
    arrSpecs() As String)
    
    Dim lngSpec As Long, strSpec As String, strSpecLow As String
    Dim dblSpec As Double
    
    'Looks at each specification on the line.
    For lngSpec = LBound(arrSpecs) To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
    
        'Saves whether no bullet or number is specified.
        If Right(strSpecLow, 9) = "no number" _
            Or Right(strSpecLow, 9) = "no bullet" _
            Or Right(strSpecLow, 10) = "no numbers" _
            Or Right(strSpecLow, 10) = "no bullets" _
            Or strSpecLow = "no numbers and bullets" _
            Or strSpecLow = "no bullets and numbers" _
            Or strSpecLow = "no numbers or bullets" _
            Or strSpecLow = "no bullets or numbers" Then
            arrList(lngLevel, 2) = ""
            arrList(lngLevel, 4) = wdListNumberStyleNone
        
        'Saves whether a tab or space follows (spec 3).
        ElseIf Right(strSpecLow, 12) = "after bullet" _
            Or Right(strSpecLow, 14) = "follows bullet" _
            Or Right(strSpecLow, 12) = "after number" _
            Or Right(strSpecLow, 14) = "follows number" Then
            If Split(strSpecLow, " ")(0) = "one" _
                Or Split(strSpecLow, " ")(0) = "a" _
                Or Split(strSpecLow, " ")(0) = "only" Then
                strSpecLow = Split(strSpecLow, " ")(1)
            Else
                strSpecLow = Split(strSpecLow, " ")(0)
            End If
            dblSpec = wdTrailingSpace
            If strSpecLow = "tab" _
                Or strSpecLow = "tabs" Then
                dblSpec = wdTrailingTab
            ElseIf strSpecLow = "nothing" _
                Or strSpecLow = "no" Then
                dblSpec = wdTrailingNone
            End If
            arrList(lngLevel, 3) = dblSpec
        
        'Saves the bullet or number font name (spec 11).
        ElseIf Right(strSpecLow, 11) = "bullet font" _
            Or Right(strSpecLow, 11) = "number font" Then
            strSpec = Left(strSpec, Len(strSpec) - 12)
            strSpecLow = LCase(strSpec)
            If strSpecLow = "body" Then
                strSpec = "+Body"
            ElseIf strSpecLow = "headings" _
                Or strSpecLow = "heading" Then
                strSpec = "+Headings"
            ElseIf strSpecLow = "default" Then
                strSpec = ""
            End If
            arrList(lngLevel, 11) = strSpec
        
        'Saves the number bold spec (spec 12).
        ElseIf strSpecLow = "bold bullet" _
            Or strSpecLow = "bold bullets" _
            Or strSpecLow = "bold number" _
            Or strSpecLow = "bold numbers" Then
            arrList(lngLevel, 12) = True
        
        'Saves the number italic spec (spec 13).
        ElseIf strSpecLow = "italic number" _
            Or strSpecLow = "italic numbers" Then
            arrList(lngLevel, 13) = True
        ElseIf strSpecLow = "bold italic number" _
            Or strSpecLow = "bold italic numbers" _
            Or strSpecLow = "italic bold number" _
            Or strSpecLow = "italic bold numbers" _
            Or strSpecLow = "bold and italic number" _
            Or strSpecLow = "bold and italic numbers" _
            Or strSpecLow = "italic and bold number" _
            Or strSpecLow = "italic and bold numbers" _
            Then
            arrList(lngLevel, 12) = True
            arrList(lngLevel, 13) = True
        
        'Saves the bullet or number color (spec 14).
        ElseIf Right(strSpecLow, 12) = "bullet color" _
            Or Right(strSpecLow, 12) = "number color" Then
            strSpec = Split(strSpec, " ")(0)
            strSpecLow = LCase(strSpec)
            If Left(strSpec, 1) = "#" Then
                strSpec = Right(strSpec, Len(strSpec) - 1)
                strSpec _
                    = Right(strSpec, 2) _
                    & Mid(strSpec, 3, 2) _
                    & Left(strSpec, 2)
                dblSpec = Val("&H" & strSpec)
                arrList(lngLevel, 14) = dblSpec
            ElseIf strSpecLow = "black" Then
                dblSpec = wdColorBlack
                arrList(lngLevel, 14) = dblSpec
            End If
        
        'Saves the bullet or number indent (_, 5).
        ElseIf Right(strSpecLow, 13) = "bullet indent" _
            Or Right(strSpecLow, 13) = "number indent" Then
            arrList(lngLevel, 5) = InchesToPoints(dblSpec)
        'Saves the text indent (_, 7).
        ElseIf Right(strSpecLow, 11) = "text indent" Then
            arrList(lngLevel, 7) = InchesToPoints(dblSpec)
        
        'If bullets, saves bullets (_, 2) and style (_, 4).
        ElseIf Right(strSpecLow, 6) = "bullet" _
            And Left(strSpecLow, 11) <> "followed by" Then
            arrList(lngLevel, 2) = Left(strSpec, 1)
            arrList(lngLevel, 4) = wdListNumberStyleBullet
        
        'If numbers, saves the number specs.
        ElseIf Right(strSpecLow, 6) = "number" Then
            'Saves the number format (_, 2).
            strSpec = Split(strSpec, " ")(0)
                'Removes quotation marks.
                If Left(strSpec, 1) = """" Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                End If
                If Right(strSpec, 1) = """" Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                End If
            arrList(lngLevel, 2) = strSpec
            'Saves the number style (_, 4).
            dblSpec = wdListNumberStyleArabic
            If InStr(strSpecLow, "uppercase roman") <> 0 Then
                dblSpec = wdListNumberStyleUppercaseRoman
            ElseIf InStr(strSpecLow, "lowercase roman") <> 0 Then
                dblSpec = wdListNumberStyleLowercaseRoman
            ElseIf InStr(strSpecLow, "uppercase letter") <> 0 Then
                dblSpec = wdListNumberStyleUppercaseLetter
            ElseIf InStr(strSpecLow, "lowercase letter") <> 0 Then
                dblSpec = wdListNumberStyleLowercaseLetter
            ElseIf InStr(strSpecLow, "legal") <> 0 Then
                dblSpec = wdListNumberStyleLegal
            End If
            arrList(lngLevel, 4) = dblSpec
        
        End If
    Next
End Sub
```
