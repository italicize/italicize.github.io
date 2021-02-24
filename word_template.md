# Define styles for a new Word&nbsp;template: first example

> Microsoft Word users (and who among us isn't) are invited to try a macro I wrote, hoping to make new templates simple. 
>
> A new template means many styles to define, which means many clicks and drop-downs. Instead, try a macro that reads a description, such as "Body Text style, 6 pt after, 0.5&quot; left indent." 
>
> I'm interested to hear whether my shortcut seems short to others or only me. I'm interested to hear what other shortcuts people know, which might be shorter still. I'm interested in every suggestion and comment you have. 

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

####  Open a document without styles

1. Type **winword /a /w** in the Windows taskbar and press **Enter**. \
<span style='font-size:small; color:darkgray;'>&#128712; The /a switch opens Word without opening your Normal template, which might have custom styles. The /w switch opens a new blank document. For more info see [Command-line switches for Microsoft Office products](https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6).</span>

#### Set the style defaults

1. In Word, click the **Design** menu and click **Fonts**.
1. Select the theme fonts. For these sample styles, click **Arial**.
1. In the search box in the menu bar, type **styles** and press **Enter**.
1. Click **Manage Styles**, the third button in the Styles pane.
1. Click the **Set Defaults** tab in the Manage Styles window. 
    1. Select a font size. For these sample styles, select **10**. 
    1. For the paragraph spacing after, select **0 pt**. 
    1. Select the line spacing. For these sample styles, leave **Multiple** and type **1.04**.
    1. Click **OK**. 

#### Add the style descriptions and macro

1. Copy the style descriptions (see above).
1. Right-click the Word document and select the paste option **Keep Text Only**.
1. Copy the macro (see below).
1. In the search box in the menu bar, type **visual basic** and press **Enter**.
    1. In the Visual Basic window, click **Insert** and **Module**.
    1. Click **Edit** and **Paste**.
    1. Click **File** and **Close and Return to Microsoft Word**.
1. In Word, click the **View** menu and click **Macros**.
1. Select the macro **sctApplySpecs** and click **Run**.

#### Save the file

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

```vb
Const sctSpecs As Long = 20
Const sctDefaultStyleGallery As String = "Normal, No Spacing, Heading 1, " _
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
    Dim arrList() As Variant, strList As String, lngList As Long
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
        If strLabelLow = "margins" Or strLabelLow = "margin" Then
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
        
        'If the line begins "Style defaults," then...
        If strLabelLow = "style defaults" Or strLabelLow = "style default" _
            Or strLabelLow = "defaults for all defined styles" _
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
        ElseIf strLabelLow = "styles gallery" _
            Or strLabelLow = "style gallery" Then
            'Removes the defaults.
            arrDefaultStyleGallery = Split(sctDefaultStyleGallery, ", ")
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
                And InStr(strLabelLow, " multi") <> 0 Then
                strList = Left(strLabel, InStr(strLabelLow, " multi") - 1)
            Else
                strList = Left(strLabel, InStr(strLabelLow, " list") - 1)
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
                
                'If a line has defaults for the list...
                If Left(strLabel, Len(strList)) = strList _
                    And (Right(strLabelLow, 9) = " defaults" _
                    Or Right(strLabelLow, 8) = " default") Then
                    For lngLevel = 1 To lngLevels
                        '...saves the specs in the array...
                        sctDefineList arrList, lngLevel, arrSpecs
                        '...and applies any style specs.
                        sctDefineStyle arrList(lngLevel, 1), arrSpecs
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
            
            'Reapplies any style specs, in case they override the list defaults.
            For lngListPara = LBound(arrParas) To UBound(arrParas)
                strPara = arrParas(lngListPara)
                'Saves the specifications on each line (between commas).
                arrSpecs = Split(strPara, ", ")
                'Saves the first specification on each line.
                strLabel = arrSpecs(0)
                strLabelLow = LCase(strLabel)
                
                'If a line has specs for a style...
                If Right(strLabelLow, 5) = "style" Then
                    strStyle = Left(strLabel, InStr(strLabelLow, " style") - 1)
                    '...and if the style is in the list...
                    For lngLevel = 1 To lngLevels
                        If arrList(lngLevel, 1) = strStyle Then
                            '...applies any style specs again.
                            sctDefineStyle strStyle, arrSpecs
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
                    .NumberFormat = arrList(lngLevel, 2)
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
            sctInsertSampleText arrStyles
            .EndKey Unit:=wdStory
        End With
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

Private Sub sctDefineStyle(ByVal strStyle As String, arrSpecs() As String)
    
    Dim lngType As Long, lngSpec As Long, strSpec As String, dblSpec As Double
    Dim strSpecLow As String, dblSpec2 As Double
    Dim objStyle As Object, objFont As Object, objFormat As Object
    
    lngType = ActiveDocument.Styles(strStyle).Type
    
    'Looks at each specification in the array.
    For lngSpec = 1 To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
        
        Set objStyle = ActiveDocument.Styles(strStyle)
        Set objFont = objStyle.Font
        Set objFormat = objStyle.ParagraphFormat
        
        If Left(strSpecLow, 8) = "based on" Then '----------- based on style
            strSpec = Right(strSpec, Len(strSpec) - 9)
            strSpecLow = LCase(strSpec)
            If strSpecLow = "no style" Then
                objStyle.BaseStyle = ""
            ElseIf strStyle <> "Normal" _
                And strStyle <> "Default Paragraph Font" Then
                objStyle.BaseStyle = strSpec
            End If
        ElseIf Left(strSpecLow, 11) = "followed by" Then '-- following style
            strSpec = Right(strSpec, Len(strSpec) - 12)
            strSpecLow = LCase(strSpec)
            If Right(strSpecLow, 6) = " style" Then
                strSpec = Left(strSpec, Len(strSpec) - 6)
            End If
            objStyle.NextParagraphStyle = strSpec
        ElseIf Left(strSpecLow, 13) = "space between" _
            Or Left(strSpecLow, 17) = "add space between" _
            Then '-------------------------------------------- space between
            objStyle.NoSpaceBetweenParagraphsOfSameStyle = False
        ElseIf Left(strSpecLow, 16) = "no space between" _
            Or Left(strSpecLow, 23) = "don't add space between" _
            Or Left(strSpecLow, 23) = "don’t add space between" _
            Or Left(strSpecLow, 24) = "do not add space between" Then
            objStyle.NoSpaceBetweenParagraphsOfSameStyle = True
        
        ElseIf Right(strSpecLow, 4) = "font" _
            And Right(strSpecLow, 11) <> "bullet font" _
            And Right(strSpecLow, 11) <> "number font" _
            And Right(strSpecLow, 11) <> "letter font" Then '---------- font
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
            objFont.Name = strSpec
        ElseIf Right(strSpecLow, 4) = "size" Then '-------------------- size
            objFont.Size = Val(strSpec)
        ElseIf strSpecLow = "bold" Then '------------------------------ bold
            objFont.Bold = True
        ElseIf strSpecLow = "not bold" Or strSpecLow = "no bold" Then
            objFont.Bold = False
        ElseIf strSpecLow = "italic" Then '-------------------------- italic
            objFont.Italic = True
        ElseIf strSpecLow = "not italic" Or strSpecLow = "no italic" Then
            objFont.Italic = False
        ElseIf strSpecLow = "bold and italic" _
            Or strSpecLow = "italic and bold" Then
            objFont.Bold = True
            objFont.Italic = True
        ElseIf strSpecLow = "small caps" Then '------------------ small caps
            objFont.SmallCaps = True
        ElseIf strSpecLow = "uppercase" Or strSpecLow = "all caps" _
            Then '----------------------------------------------------- caps
            objFont.AllCaps = True
        ElseIf Right(strSpecLow, 5) = "color" _
            And Right(strSpecLow, 12) <> "bullet color" _
            And Right(strSpecLow, 12) <> "number color" _
            And Right(strSpecLow, 12) <> "letter color" Then '-------- color
            strSpec = Split(strSpec, " ")(0)
            strSpecLow = LCase(strSpec)
            If Left(strSpec, 1) = "#" Then
                strSpec = Right(strSpec, Len(strSpec) - 1)
                strSpec = Right(strSpec, 2) & Mid(strSpec, 3, 2) _
                    & Left(strSpec, 2)
                dblSpec = Val("&H" & strSpec)
                objFont.Color = dblSpec
            ElseIf strSpecLow = "automatic" Or strSpecLow = "auto" _
                Or strSpecLow = "no" Then
                dblSpec = wdColorAutomatic
                objFont.Color = dblSpec
            ElseIf strSpecLow = "black" Then
                dblSpec = wdColorBlack
                objFont.Color = dblSpec
            End If
        ElseIf strSpecLow = "normal character spacing" Then ' letter spacing
            objFont.Spacing = 0
        ElseIf Right(strSpecLow, 17) = "character spacing" Then
            objFont.Spacing = dblSpec
        ElseIf strSpecLow = "kerning" Then '------------------------ kerning
            objFont.Kerning = 8
        ElseIf strSpecLow = "no kerning" Then
            objFont.Kerning = 0
        
        ElseIf lngType = wdStyleTypeParagraph Then
            If Right(strSpecLow, 11) = "left indent" Then '--------- indents
                objFormat.LeftIndent = InchesToPoints(dblSpec)
            ElseIf Right(strSpecLow, 12) = "right indent" Then
                objFormat.RightIndent = InchesToPoints(dblSpec)
            ElseIf (Right(strSpecLow, 6) = "before" _
                And strSpecLow <> "page break before" _
                And strSpecLow <> "no page break before") _
                Or Right(strSpecLow, 5) = "above" Then '------- space before
                objFormat.SpaceBefore = dblSpec
            ElseIf Right(strSpecLow, 5) = "after" _
                Or Right(strSpecLow, 5) = "below" Then '-------- space after
                objFormat.SpaceAfter = dblSpec
            ElseIf Right(strSpecLow, 12) = "line spacing" Then 'line spacing
                If Split(strSpecLow, " ")(1) = "pt" _
                    Or Split(strSpecLow, " ")(1) = "pt." _
                    Or Split(strSpecLow, " ")(1) = "point" _
                    Or Split(strSpecLow, " ")(1) = "points" Then
                    objFormat.LineSpacingRule = wdLineSpaceExactly
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(0) = "exact" _
                    Or Split(strSpecLow, " ")(0) = "exactly" Then
                    dblSpec = Val(Split(strSpec, " ")(1))
                    objFormat.LineSpacingRule = wdLineSpaceExactly
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(1) = "least" Then
                    dblSpec = Val(Split(strSpec, " ")(2))
                    objFormat.LineSpacingRule = wdLineSpaceAtLeast
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(0) = "single" Then
                    objFormat.LineSpacingRule = wdLineSpaceSingle
                Else
                    objFormat.LineSpacingRule = wdLineSpaceMultiple
                    objFormat.LineSpacing = LinesToPoints(dblSpec)
                End If
            ElseIf strSpecLow = "left aligned" _
                Or strSpecLow = "left align" _
                Or strSpecLow = "aligned left" _
                Or strSpecLow = "align left" _
                Or strSpecLow = "right aligned" _
                Or strSpecLow = "right align" _
                Or strSpecLow = "aligned right" _
                Or strSpecLow = "align right" _
                Or strSpecLow = "centered" Or strSpecLow = "center" _
                Or strSpecLow = "center aligned" _
                Or strSpecLow = "aligned center" _
                Or strSpecLow = "center align" _
                Or strSpecLow = "align center" _
                Or strSpecLow = "justified" Or strSpecLow = "justify" _
                Then '-------------------------------------------- alignment
                dblSpec = wdAlignParagraphLeft
                If strSpecLow = "right aligned" _
                    Or strSpecLow = "right align" _
                    Or strSpecLow = "aligned right" _
                    Or strSpecLow = "align right" Then
                    dblSpec = wdAlignParagraphRight
                ElseIf strSpecLow = "centered" Or strSpecLow = "center" _
                    Or strSpecLow = "center aligned" _
                    Or strSpecLow = "aligned center" _
                    Or strSpecLow = "center align" _
                    Or strSpecLow = "align center" Then
                    dblSpec = wdAlignParagraphCenter
                ElseIf strSpecLow = "justified" Or strSpecLow = "justify" _
                    Then
                    dblSpec = wdAlignParagraphJustify
                End If
                objFormat.Alignment = dblSpec
            ElseIf strSpecLow = "widow/orphan control" _
                Or strSpecLow = "orphan/widow control" _
                Or strSpecLow = "widow and orphan control" _
                Or strSpecLow = "orphan and widow control" _
                Or strSpecLow = "widow control" _
                Or strSpecLow = "orphan control" Then '------- widow control
                objFormat.WidowControl = True
            ElseIf strSpecLow = "no widow/orphan control" _
                Or strSpecLow = "no orphan/widow control" _
                Or strSpecLow = "no widow and orphan control" _
                Or strSpecLow = "no orphan and widow control" _
                Or strSpecLow = "no widow or orphan control" _
                Or strSpecLow = "no orphan or widow control" _
                Or strSpecLow = "no widow control" _
                Or strSpecLow = "no orphan control" Then
                objFormat.WidowControl = False
            ElseIf Left(strSpecLow, 14) = "keep with next" _
                Or Left(strSpecLow, 24) = "keep paragraph with next" _
                Or Left(strSpecLow, 30) = "keep the paragraph with the ne" _
                Or Left(strSpecLow, 22) = "no page break after" _
                Or Left(strSpecLow, 22) = "no page break below" _
                Or Left(strSpecLow, 28) = "don't allow page break after" _
                Or Left(strSpecLow, 30) = "don't allow a page break after" _
                Or Left(strSpecLow, 28) = "don't allow page break below" _
                Or Left(strSpecLow, 30) = "don't allow a page break below" _
                Or Left(strSpecLow, 28) = "don’t allow page break after" _
                Or Left(strSpecLow, 30) = "don’t allow a page break after" _
                Or Left(strSpecLow, 28) = "don’t allow page break below" _
                Or Left(strSpecLow, 30) = "don’t allow a page break below" _
                Or Left(strSpecLow, 29) = "do not allow page break after" _
                Or Left(strSpecLow, 30) = "do not allow a page break afte" _
                Or Left(strSpecLow, 29) = "do not allow page break below" _
                Or Left(strSpecLow, 30) = "do not allow a page break belo" _
                Then '--------------------------------------- keep with next
                objFormat.KeepWithNext = True
            ElseIf Left(strSpecLow, 17) = "no keep with next" _
                Or Left(strSpecLow, 27) = "no keep paragraph with next" _
                Or Left(strSpecLow, 30) = "no keep the paragraph with nex" _
                Or Left(strSpecLow, 20) = "don't keep with next" _
                Or Left(strSpecLow, 30) = "don't keep paragraph with next" _
                Or Left(strSpecLow, 30) = "don't keep the paragraph with " _
                Or Left(strSpecLow, 20) = "don’t keep with next" _
                Or Left(strSpecLow, 30) = "don’t keep paragraph with next" _
                Or Left(strSpecLow, 30) = "don’t keep the paragraph with " _
                Or Left(strSpecLow, 21) = "do not keep with next" _
                Or Left(strSpecLow, 30) = "do not keep paragraph with nex" _
                Or Left(strSpecLow, 30) = "do not keep the paragraph with" _
                Or Left(strSpecLow, 22) = "allow page break after" _
                Or Left(strSpecLow, 24) = "allow a page break after" _
                Or Left(strSpecLow, 22) = "allow page break below" _
                Or Left(strSpecLow, 24) = "allow a page break below" _
                Then
                objFormat.KeepWithNext = False
            ElseIf Left(strSpecLow, 13) = "keep together" _
                Or Left(strSpecLow, 19) = "keep lines together" _
                Or Left(strSpecLow, 29) = "keep paragraph lines together" _
                Or Left(strSpecLow, 30) = "keep the paragraph lines toget" _
                Or Left(strSpecLow, 30) = "keep the paragraph lines on th" _
                Or Left(strSpecLow, 21) = "keep on the same page" _
                Or Left(strSpecLow, 27) = "keep lines on the same page" _
                Or Left(strSpecLow, 30) = "keep paragraph lines on the sa" _
                Then '---------------------------------- keep lines together
                objFormat.KeepTogether = True
            ElseIf strSpecLow = "no keep together" _
                Or strSpecLow = "no keep lines together" _
                Or strSpecLow = "no keep paragraph lines together" _
                Or strSpecLow = "don't keep together" _
                Or strSpecLow = "don't keep lines together" _
                Or strSpecLow = "don't keep paragraph lines together" _
                Or strSpecLow = "don’t keep together" _
                Or strSpecLow = "don’t keep lines together" _
                Or strSpecLow = "don’t keep paragraph lines together" _
                Or strSpecLow = "do not keep together" _
                Or strSpecLow = "do not keep lines together" _
                Or strSpecLow = "do not keep paragraph lines together" _
                Or Left(strSpecLow, 19) = "allow page break in" _
                Or Left(strSpecLow, 21) = "allow a page break in" _
                Or Left(strSpecLow, 23) = "allow page break within" _
                Or Left(strSpecLow, 25) = "allow a page break within" _
                Then
                objFormat.KeepTogether = False
            ElseIf Left(strSpecLow, 17) = "page break before" _
                Or Left(strSpecLow, 25) = "require page break before" _
                Or Left(strSpecLow, 27) = "require a page break before" _
                Or Left(strSpecLow, 16) = "page break above" _
                Or Left(strSpecLow, 24) = "require page break above" _
                Or Left(strSpecLow, 26) = "require a page break above" _
                Then '------------------------------------ page break before
                objFormat.PageBreakBefore = True
            ElseIf Left(strSpecLow, 20) = "no page break before" _
                Or Left(strSpecLow, 28) = "no require page break before" _
                Or Left(strSpecLow, 30) = "no require a page break before" _
                Or Left(strSpecLow, 30) = "don't require page break befor" _
                Or Left(strSpecLow, 30) = "don't require a page break bef" _
                Or Left(strSpecLow, 30) = "don’t require page break befor" _
                Or Left(strSpecLow, 30) = "don’t require a page break bef" _
                Or Left(strSpecLow, 30) = "do not require page break befo" _
                Or Left(strSpecLow, 30) = "do not require a page break be" _
                Or Left(strSpecLow, 19) = "no page break above" _
                Or Left(strSpecLow, 27) = "no require page break above" _
                Or Left(strSpecLow, 29) = "no require a page break above" _
                Or Left(strSpecLow, 30) = "don't require page break above" _
                Or Left(strSpecLow, 30) = "don't require a page break abo" _
                Or Left(strSpecLow, 30) = "don’t require page break above" _
                Or Left(strSpecLow, 30) = "don’t require a page break abo" _
                Or Left(strSpecLow, 30) = "do not require page break abov" _
                Or Left(strSpecLow, 30) = "do not require a page break ab" _
                Then
                objFormat.PageBreakBefore = False
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
                                objFormat.LineWidth = wdUndefined
                            Case Is <= 0.25
                                objFormat.LineWidth = wdLineWidth025pt
                            Case Is <= 0.5
                                objFormat.LineWidth = wdLineWidth050pt
                            Case Is <= 0.75
                                objFormat.LineWidth = wdLineWidth075pt
                            Case Is <= 1
                                objFormat.LineWidth = wdLineWidth100pt
                            Case Is <= 1.5
                                objFormat.LineWidth = wdLineWidth150pt
                            Case Is <= 2.25
                                objFormat.LineWidth = wdLineWidth225pt
                            Case Is <= 3
                                objFormat.LineWidth = wdLineWidth300pt
                            Case Is <= 4.5
                                objFormat.LineWidth = wdLineWidth450pt
                            Case Is > 4.5
                                objFormat.LineWidth = wdLineWidth600pt
                        End Select
                    End With
                End If
            ElseIf strSpecLow = "no tabs" _
                Or strSpecLow = "clear tabs" _
                Or strSpecLow = "clear all tabs" _
                Or strSpecLow = "remove tabs" _
                Or strSpecLow = "remove all tabs" Then '--------------- tabs
                objFormat.TabStops.ClearAll
            ElseIf strSpecLow = "center tab" _
                Or strSpecLow = "centered tab" Then
                dblSpec = ActiveDocument.PageSetup.PageWidth _
                    - ActiveDocument.PageSetup.LeftMargin _
                    - ActiveDocument.PageSetup.RightMargin _
                    - objFormat.LeftIndent - objFormat.RightIndent
                objFormat.TabStops.Add Position:=(dblSpec / 2), _
                    Alignment:=wdAlignTabCenter, _
                    Leader:=wdTabLeaderSpaces
            ElseIf strSpecLow = "right tab" Then
                dblSpec = ActiveDocument.PageSetup.PageWidth _
                    - ActiveDocument.PageSetup.LeftMargin _
                    - ActiveDocument.PageSetup.RightMargin _
                    - objFormat.LeftIndent - objFormat.RightIndent
                objFormat.TabStops.Add Position:=dblSpec, _
                    Alignment:=wdAlignTabRight, _
                    Leader:=wdTabLeaderSpaces
            ElseIf Right(strSpecLow, 3) = "tab" Then
                If InStr(strSpecLow, "left") <> 0 Then
                    dblSpec2 = wdAlignTabLeft
                ElseIf InStr(strSpecLow, "center") <> 0 Then
                    dblSpec2 = wdAlignTabCenter
                ElseIf InStr(strSpecLow, "right") <> 0 Then
                    dblSpec2 = wdAlignTabRight
                ElseIf InStr(strSpecLow, "decimal") <> 0 Then
                    dblSpec2 = wdAlignTabDecimal
                Else
                    dblSpec2 = 99
                End If
                If dblSpec2 <> 99 Then
                    objFormat.TabStops.Add Position:=(dblSpec), _
                        Alignment:=dblSpec2, _
                        Leader:=wdTabLeaderSpaces
                End If
            End If
        End If
    Next lngSpec
End Sub

Private Sub sctDefineList(ByRef arrList() As Variant, ByVal lngLevel As Long, _
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
            Or Right(strSpecLow, 9) = "no letter" _
            Or Right(strSpecLow, 10) = "no numbers" _
            Or Right(strSpecLow, 10) = "no bullets" _
            Or Right(strSpecLow, 10) = "no letters" Then
            arrList(lngLevel, 2) = ""
            arrList(lngLevel, 4) = wdListNumberStyleNone
        
        'Saves whether a tab or space follows (spec 3).
        ElseIf Right(strSpecLow, 12) = "after bullet" _
            Or Right(strSpecLow, 14) = "follows bullet" _
            Or Right(strSpecLow, 16) = "following bullet" _
            Or Right(strSpecLow, 12) = "after number" _
            Or Right(strSpecLow, 14) = "follows number" _
            Or Right(strSpecLow, 16) = "following letter" _
            Or Right(strSpecLow, 14) = "follows letter" _
            Or Right(strSpecLow, 16) = "following letter" Then
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
        
        'Saves the font for the bullet or number (spec 11).
        ElseIf Right(strSpecLow, 11) = "bullet font" _
            Or Right(strSpecLow, 11) = "number font" _
            Or Right(strSpecLow, 11) = "letter font" Then
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
            Or strSpecLow = "bold numbers" _
            Or strSpecLow = "bold letter" _
            Or strSpecLow = "bold letters" Then
            arrList(lngLevel, 12) = True
        
        'Saves the number italic spec (spec 13).
        ElseIf strSpecLow = "italic number" _
            Or strSpecLow = "italic numbers" _
            Or strSpecLow = "italic letter" _
            Or strSpecLow = "italic letters" Then
            arrList(lngLevel, 13) = True
        ElseIf strSpecLow = "bold italic number" _
            Or strSpecLow = "bold italic numbers" _
            Or strSpecLow = "italic bold number" _
            Or strSpecLow = "italic bold numbers" _
            Or strSpecLow = "bold and italic number" _
            Or strSpecLow = "bold and italic numbers" _
            Or strSpecLow = "italic and bold number" _
            Or strSpecLow = "italic and bold numbers" _
            Or strSpecLow = "bold italic letter" _
            Or strSpecLow = "bold italic letters" _
            Or strSpecLow = "italic bold letter" _
            Or strSpecLow = "italic bold letters" _
            Or strSpecLow = "bold and italic letter" _
            Or strSpecLow = "bold and italic letters" _
            Or strSpecLow = "italic and bold letter" _
            Or strSpecLow = "italic and bold letters" Then
            arrList(lngLevel, 12) = True
            arrList(lngLevel, 13) = True
        
        'Saves the bullet or number color (spec 14).
        ElseIf Right(strSpecLow, 12) = "bullet color" _
            Or Right(strSpecLow, 12) = "number color" _
            Or Right(strSpecLow, 12) = "letter color" Then
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
            Or Right(strSpecLow, 13) = "number indent" _
            Or Right(strSpecLow, 13) = "letter indent" Then
            arrList(lngLevel, 5) = InchesToPoints(dblSpec)
        'Saves the text indent (_, 7).
        ElseIf Right(strSpecLow, 11) = "text indent" Then
            arrList(lngLevel, 7) = InchesToPoints(dblSpec)
        
        'If bullets, saves bullets (_, 2) and style (_, 4).
        ElseIf Right(strSpecLow, 6) = "bullet" _
            And Left(strSpecLow, 8) <> "based on" _
            And Left(strSpecLow, 11) <> "followed by" Then
            arrList(lngLevel, 2) = Left(strSpec, 1)
            arrList(lngLevel, 4) = wdListNumberStyleBullet
        
        'If numbers, saves the number specs.
        ElseIf (Right(strSpecLow, 6) = "number" _
            Or Right(strSpecLow, 6) = "letter") _
            And Left(strSpecLow, 8) <> "based on" _
            And Left(strSpecLow, 11) <> "followed by" _
            Then
            'Saves the number format (_, 2).
            strSpec = Split(strSpec, " ")(0)
                'Removes quotation marks.
                If Left(strSpec, 1) = Chr(34) Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                ElseIf Left(strSpec, 1) = Chr(147) Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                End If
                If Right(strSpec, 1) = Chr(34) Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                ElseIf Right(strSpec, 1) = Chr(148) Then
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
    Next lngSpec
End Sub

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

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
```

---

### Legal

Copyright (C) 2020 Jay Martin. 

**Permission is granted** to copy, distribute and/or modify this document
under the terms of the [GNU Free Documentation License, Version 1.3](https://www.gnu.org/licenses/fdl-1.3.txt)
or any later version published by the Free Software Foundation; 
with no Invariant Sections, no Front-Cover Texts, and no Back-Cover Texts.
A copy of the license is included in the section entitled "[GNU Free Documentation License](fdl-1.3.md)."

Microsoft Windows is a trademark of Microsoft. All other trademarks are the property of their respective owners. 

<!--- --->
