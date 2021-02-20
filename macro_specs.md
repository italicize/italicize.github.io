# Style specifications

The macro reads specific words. The specifications can appear in any order.

#### `Margins, . . .` defines the margins.

For example, the Microsoft default margins are **Margins, 1" left, 1" right, 1" top, 1" bottom, 0.5" header, 0.5" footer.**

Specification|Examples|Note
:---|:---|:---
_ left,<br>_ right,<br>_ top,<br>_ bottom,<br>_ header,<br>_ footer|0.5&quot;&nbsp;left,<br>-.25&nbsp;inches&nbsp;left|The number at the beginning is read to set a specific margin, such as the left margin.<br>The number is treated as inches. Currently the macro doesn't look for a unit, such as cm or pt.
mirror&nbsp;margins,<br>no&nbsp;mirror&nbsp;margins||Swaps the left and right margins on odd and even pages<br>(or doesn't swap if "no" at the beginning). 

#### `Defaults for all defined styles, . . .` defines all styles.<br>

The available specifications are the same as for style definitions (see next).<br>
Style definitions don't need to repeat any defaults, which allows style definitions to be shorter.<br>
Style definitions can include on the specifications that supersede the defaults.

#### `_ style, . . .` defines a style.<br>
For example, a Microsoft default style is **Body Text style, body font, 11 pt, 1.08 line spacing, 6 pt after, widow/orphan control, no kerning, based on Normal, followed by Body Text.**

        
#### `Styles gallery, . . .` defines the Styles gallery in the Home menu.

For example, the Microsoft default Styles gallery is **Styles gallery, Normal, No Spacing, Heading 1, Heading 2, Heading 3, Heading 4, Heading 5, Heading 6, Heading 7, Heading 8, Heading 9, Title, Subtitle, Subtle Emphasis, Emphasis, Intense Emphasis, Strong, Quote, Intense Quote, Subtle Reference, Intense Reference, Book Title, List Paragraph, Caption, TOC Heading.**

#### `_ list, . . .` defines a list name and the styles in a list.

The beginning is read as list name. Several variations are accepted, such as "_ multilevel list styles."<br>
The macro adds a new list, if it doesn't exist. The specifications are read as style names, in order, up to nine.<br>
For example, **ListBullets list, List Bullet, List Bullet 2, List Bullet 3, List Bullet 4, List Bullet 5.**

#### `_ bullet defaults, . . .` or `_ number defaults, . . .` defines all bullets or numbers in a list.

The beginning is read as list name. The specifications . . . . . . . <br>
For example, **ListBullets bullet defaults, Body bullet font, tab after bullet.**

#### Paragraph style specifications
Specification|Examples|Note
:---|:---|:---
based&nbsp;on _|based&nbsp;on&nbsp;no&nbsp;style,<br>based&nbsp;on&nbsp;Heading&nbsp;1|The end is read as a style name, the style with specifications to copy. The Microsoft defaults are "based on Normal" for paragraph styles and "based on Default Paragraph Font" for font styles.
followed&nbsp;by _|followed&nbsp;by&nbsp;Body&nbsp;Text|The end is read as a style name, the style for the next paragraph, after pressing Enter.
space&nbsp;between,<br>no&nbsp;space&nbsp;between||Several variations are accepted, but as "add space between paragraphs of the same style" and "don't add space between paragraphs."

#### Font and paragraph style specifications

Specification|Examples|Note
:---|:---|:---
_ font|Palatino&nbsp;Linotype&nbsp;font,<br>body&nbsp;font,<br>headings&nbsp;font|The beginning is read as a font name. "Body font" and "headings font" use the defaults (defined through the Design menu). 

" Then '------------------------ font
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

Private Sub sctDefineList(ByRef arrList() As Variant, ByVal lngLevel As Long, _
    arrSpecs() As String)
    
    Dim lngSpec As Long, strSpec As String, strSpecLow As String
    Dim dblSpec As Double
    
    'Looks at each specification on the line.
    For lngSpec = LBound(arrSpecs) To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
        
        'Saves defaults for true/false and number values
        arrList(lngLevel, 3) = wdTrailingSpace
        arrList(lngLevel, 4) = wdListNumberStyleNone
        arrList(lngLevel, 12) = False 'not bold
        arrList(lngLevel, 13) = False 'not italic
        arrList(lngLevel, 14) = wdColorAutomatic
    
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

            
            'Applies the list template specifications.
            For lngLevel = 1 To lngLevels
                With objListTemplate.ListLevels(lngLevel)
                    arrList(lngLevel, 2) = arrList(lngLevel, 2)
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

