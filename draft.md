# Sample specifications

Note, Copy and paste these specifications into your document, then make changes.


## General settings

Margins, 0.5&quot; left, 0.5&quot; right, 0.5&quot; top, 0.5&quot; bottom, 0.25&quot; header, 0.25&quot; footer.

Styles gallery, Body Text, List Bullet, List Number, List Continue, Heading 1, Heading 2, Heading 3, Heading 4, Caption.

Defaults for all defined styles, Body font, 11 pt size, 1.15 line spacing, 0 pt before, 0 pt after, 0&quot; left indent, 0&quot; right indent, left aligned, based on no style, followed by Body Text, normal character spacing, kerning.

Normal style, defaults.

Body Text style, 6 pt after, 0.5&quot; left indent.


## Heading style settings

Heading 1 style, Headings font, bold, #2D415A (blue) color, 16 pt size, 12 pt before, 6 pt after, keep with next, page break before.

Heading 2 style, Body font, bold, #2D415A (blue) color, uppercase, 12 pt size, 12 pt before, 6 pt after, keep with next.

Heading 3 style, Body font, bold, #2D415A (blue) color, uppercase, 10 pt size, 12 pt before, 6 pt after, 0.5&quot; left indent, keep with next.

Heading 4 style, Body font, bold, #2D415A (blue) color, 0.5&quot; left indent, keep with next.

Heading 5 style, Body font, italic, 6 pt after, 0.5&quot; left indent, keep with next.

TOC Heading style, based on Heading 1.


## Other built-in style settings

Header style, Body font italic, #2D415A (blue) color, 10 pt size, no tabs, bottom border.

Footer style, Body font, #2D415A (blue) color, 10 pt size, centered tab, right tab, top border.

Caption style, Headings font, #2D415A (blue) color, 10 pt size, 6 pt after, 0.5&quot; left indent.

Title style, Headings font, bold, #F8F8F8 (gray) color, 42 pt size, 1.03 line spacing.

Subtitle style, Headings font, #F8F8F8 (gray) color, 24 pt size, 1.03 line spacing.


## Custom style settings

TBD


## Bullet list settings

ListBullets list styles, List Bullet, List Bullet 2, List Bullet 3, List Bullet 4, List Bullet 5.

ListBullets bullets, &#8226; bullet, &#9702; bullet, &#9642; bullet, &#9643; bullet, &#8729; bullet.

ListBullets bullet defaults, Segoe UI font, #2D415A (blue) color, tab after bullet.

List Bullet style, 0.5&quot; bullet indent, 0.75&quot; text indent, 6 pt after, space between, followed by List Bullet.

List Bullet 2 style, 0.75&quot; bullet indent, 1&quot; text indent, 6 pt after, space between, followed by List Bullet 2.

List Bullet 3 style, 1&quot; bullet indent, 1.25&quot; text indent, 6 pt after, no space between, followed by List Bullet 3.

List Bullet 4 style, 1.25&quot; bullet indent, 1.5&quot; text indent, 6 pt after, no space between, followed by List Bullet 4.

List Bullet 5 style, 1.5&quot; bullet indent, 1.75&quot; text indent, 6 pt after, no space between, followed by List Bullet 5.

End of ListBullets.


## List without bullets settings

ListContinues list styles, List Continue, List Continue 2, List Continue 3, List Continue 4, List Continue 5.

ListContinues defaults, tabs only, no bullets.

List Continue style, 0.5&quot; bullet indent, 0.75&quot; text indent, 6 pt after, space between, followed by List Bullet.

List Continue 2 style, 0.75&quot; bullet indent, 1&quot; text indent, 6 pt after, space between, followed by List Bullet 2.

List Continue 3 style, 1&quot; bullet indent, 1.25&quot; text indent, 6 pt after, no space between, followed by List Bullet 3.

List Continue 4 style, 1.25&quot; bullet indent, 1.5&quot; text indent, 6 pt after, no space between, followed by List Bullet 4.

List Continue 5 style, 1.5&quot; bullet indent, 1.75&quot; text indent, 6 pt after, no space between, followed by List Bullet 5.

End of ListContinues.


## Heading list settings

ListHeadingsToNumbers list styles, Heading 1, Heading 2, Heading 3, Heading 4, List Number, List Number 2, List Number 3, List Number 4, List Number 5.

ListHeadingsToNumbers numbers,  
&quot;%1. &quot; arabic number,  
&quot;%1.%2. &quot; arabic number,  
&quot;&quot; number,  
&quot;&quot; number,  
&quot;%5.&quot; arabic number,  
&quot;%6.&quot; lowercase letter,  
&quot;%7.&quot; lowercase roman numeral,  
&quot;%8.&quot; uppercase letter,  
&quot;%9.&quot; uppercase roman numeral.

ListHeadingsToNumbers number defaults, default font, default color, tab after number.

Heading 1 style, 0&quot; number indent, 0.5&quot; text indent.

Heading 2 style, 0&quot; number indent, 0.5&quot; text indent.

Heading 3 style, 0&quot; number indent, 0.5&quot; text indent.

Heading 4 style, 0&quot; number indent, 0.5&quot; text indent.

List Number style, 0.5&quot; number indent, 0.75&quot; text indent, 6 pt after, space between, followed by List Number.

List Number 2 style, 0.75&quot; number indent, 1&quot; text indent, 6 pt after, space between, followed by List Number 2.

List Number 3 style, 1&quot; number indent, 1.25&quot; text indent, 6 pt after, no space between, followed by List Number 3.

List Number 4 style, 1.25&quot; number indent, 1.5&quot; text indent, 6 pt after, no space between, followed by List Number 4.

List Number 5 style, 1.5&quot; number indent, 1.75&quot; text indent, 6 pt after, no space between, followed by List Number 5.

End of ListHeadingsToNumbers.


# Fallback Code

Use this if the code in "General settings.docm" has problems.

```vb
Option Explicit
'---5---10---15---20---25---30---35---40---45---50---55---60---65---70---75---80

Sub sctApplySpecs()
    Dim objParagraph As Paragraph, objListParagraph As Paragraph
    Dim arrSpecs As Variant, strSpec As String, dblSpec As Double
    Dim arrStyles As Variant, lngStyles As Long, strStyle As String
    Dim arrList As Variant, strList As String, objListTemplate As ListTemplate
    Dim rngSearch As Range, rngFound As Range
    Dim lngA As Long, lngB As Long
'    Dim dblBodyTextIndent As Double, dblBodyTextWidth As Double
    
    'Find and saves the style names.
    For Each objParagraph In ActiveDocument.Paragraphs
        arrSpecs = Split(objParagraph.Range.Text, ", ")
        strSpec = arrSpecs(0)
        If Right(strSpec, 5) = "style" _
            And Right(strSpec, 10) <> "list style" Then
            strStyle = Left(strSpec, InStr(strSpec, "style") - 2)
            If lngStyles = 0 Then
                ReDim arrStyles(lngStyles)
            Else
                ReDim Preserve arrStyles(lngStyles)
            End If
            arrStyles(lngStyles) = strStyle
            lngStyles = lngStyles + 1
            
            'Adds a style if it doesn't exist.
            If Not sctStyleExists(strStyle, ActiveDocument) Then
                dblSpec = wdStyleTypeParagraph
                For lngA = 0 To UBound(arrSpecs)
                    strSpec = arrSpecs(lngA)
                    If Right(strSpec, 15) = "character style" Then
                        dblSpec = wdStyleTypeCharacter
                    End If
                Next lngA
                ActiveDocument.Styles.Add strStyle, dblSpec
            End If
        End If
    Next objParagraph
    
    'Reads each line.
    For Each objParagraph In ActiveDocument.Paragraphs
        strSpec = objParagraph.Range.Text
        'Doesn't save a carriage return, spaces, or period at the end of a line.
        If Right(strSpec, 1) = vbCr Then
            strSpec = Left(strSpec, Len(strSpec) - 1)
        End If
        strSpec = Trim(strSpec)
        If Right(strSpec, 1) = "." Then
            strSpec = Left(strSpec, Len(strSpec) - 1)
        End If
        'Replaces manual line breaks with a space and removes extra spaces.
        strSpec = Replace(strSpec, Chr(11), " ")
        strSpec = Replace(strSpec, "   ", " ")
        strSpec = Replace(strSpec, "  ", " ")
        'Saves something instead of an empty line.
        If strSpec = "" Then strSpec = "[empty line]"
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strSpec, ", ")
'Styles-'
'-------'If the line begins "Defaults for all defined styles," then...
        If arrSpecs(0) = "Defaults for all defined styles" Then
            'Applies the default specifications to all defined styles.
            For lngA = LBound(arrStyles) To UBound(arrStyles)
                strStyle = arrStyles(lngA)
                'Sends a style name and specs to the sctDefineStyle macro.
                sctDefineStyle strStyle, arrSpecs
            Next lngA
        
        'Or if the line begins with a style name, then...
        ElseIf Right(arrSpecs(0), 5) = "style" Then
            'Applies the specifications on the line to the style.
            strStyle = Left(arrSpecs(0), InStr(arrSpecs(0), "style") - 2)
            sctDefineStyle strStyle, arrSpecs
'Lists--'
'-------'Or if the line begins with a list template name, then...
        ElseIf Right(arrSpecs(0), 11) = "list styles" _
            Or Right(arrSpecs(0), 13) = "list template" Then
            'Saves the list template name.
            strList = Left(arrSpecs(0), InStr(arrSpecs(0), " list ") - 1)
            'Saves the style names and defaults in an array (_, 1).
            ReDim Preserve arrSpecs(9)
            ReDim arrList(1 To 9, 1 To 20)
            For lngA = 1 To 9
                arrList(lngA, 1) = arrSpecs(lngA) '--------------- linked style
                arrList(lngA, 2) = "" '---------------- number format or bullet
                arrList(lngA, 3) = wdTrailingSpace '-------- trailing character
                arrList(lngA, 4) = wdListNumberStyleNone '-------- number style
                arrList(lngA, 5) = 0 '------------------------- number position
                arrList(lngA, 6) = wdListLevelAlignLeft '------------ alignment
                arrList(lngA, 7) = 0 '--------------------------- text position
                'arrList(lngA, 8) = '--------------------------- for future use
                'arrList(lngA, 9) =
                'arrList(lngA, 10) =
                arrList(lngA, 11) = "" '---------------------------------- font
                arrList(lngA, 12) = False '------------------------------- bold
                arrList(lngA, 13) = False '----------------------------- italic
                arrList(lngA, 14) = wdColorAutomatic '------------------- color
                'arrList(lngA, 15) =
                'arrList(lngA, 16) =
                'arrList(lngA, 17) =
                'arrList(lngA, 18) =
                'arrList(lngA, 19) =
                'arrList(lngA, 20) =
            Next lngA
            
            'Looks at the next 12 lines for more specifications.
            Set rngSearch = objParagraph.Range
            rngSearch.MoveEnd wdParagraph, 12
            For Each objListParagraph In rngSearch.Paragraphs
                strSpec = objListParagraph.Range.Text
                'Doesn't save a carriage return, spaces, or period at the end.
                If Right(strSpec, 1) = vbCr Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                End If
                strSpec = Trim(strSpec)
                If Right(strSpec, 1) = "." Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                End If
                'Replaces manual line breaks and removes extra spaces.
                strSpec = Replace(strSpec, Chr(11), " ")
                strSpec = Replace(strSpec, "   ", " ")
                strSpec = Replace(strSpec, "  ", " ")
                'Refills the array arrSpecs.
                arrSpecs = Split(strSpec, ", ")
                
                'If a line says "end of," stops looking at lines.
                If LCase(Left(arrSpecs(0), 4)) = "end." _
                    Or LCase(Left(arrSpecs(0), 4)) = "end," _
                    Or LCase(Left(arrSpecs(0), 4)) = "end " Then
                    Exit For
                
                'If a line has bullets, saves bullets (_, 2) and style (_, 4).
                ElseIf arrSpecs(0) = strList & " bullets" Then
                    ReDim Preserve arrSpecs(9)
                    For lngA = 1 To 9
                        strSpec = arrSpecs(lngA)
                        If strSpec = "no bullet" _
                            Or strSpec = "no bullets" _
                            Or strSpec = "" Then
                            arrList(lngA, 2) = ""
                            arrList(lngA, 4) = wdListNumberStyleNone
                        Else
                            strSpec = Left(strSpec, 1)
                            arrList(lngA, 2) = strSpec
                            arrList(lngA, 4) = wdListNumberStyleBullet
                        End If
                    Next lngA
                'If a line has numbers, saves the number specs (_, 2).
                ElseIf arrSpecs(0) = strList & " numbers" Then
                    ReDim Preserve arrSpecs(9)
                    For lngA = 1 To 9
                        strSpec = LCase(arrSpecs(lngA))
                        If strSpec = "no number" _
                            Or strSpec = "no numbers" _
                            Or strSpec = "" Then
                            arrList(lngA, 2) = ""
                            arrList(lngA, 4) = wdListNumberStyleNone
                        Else
                            'Saves the number style (_, 4).
                            dblSpec = wdListNumberStyleArabic
                            If InStr(strSpec, "uppercase roman") <> 0 Then
                                dblSpec = wdListNumberStyleUppercaseRoman
                            ElseIf InStr(strSpec, "lowercase roman") <> 0 Then
                                dblSpec = wdListNumberStyleLowercaseRoman
                            ElseIf InStr(strSpec, "uppercase letter") <> 0 Then
                                dblSpec = wdListNumberStyleUppercaseLetter
                            ElseIf InStr(strSpec, "lowercase letter") <> 0 Then
                                dblSpec = wdListNumberStyleLowercaseLetter
                            ElseIf InStr(strSpec, "legal") <> 0 Then
                                dblSpec = wdListNumberStyleLegal
                            End If
                            arrList(lngA, 4) = dblSpec
                            'Saves the number format (_, 2).
                            strSpec = Split(strSpec, " ")(0)
                                'Removes quotation marks.
                                If Left(strSpec, 1) = """" Then
                                    strSpec = Right(strSpec, Len(strSpec) - 1)
                                End If
                                If Right(strSpec, 1) = """" Then
                                    strSpec = Left(strSpec, Len(strSpec) - 1)
                                End If
                            arrList(lngA, 2) = strSpec
                        End If
                    Next lngA
                
                'If a line has defaults, saves the default specs in the array.
                ElseIf arrSpecs(0) = strList & " bullet defaults" _
                    Or arrSpecs(0) = strList & " number defaults" _
                    Or arrSpecs(0) = strList & " numbering defaults" _
                    Or arrSpecs(0) = strList & " defaults" _
                    Then
                    'Looks at each specification in the line.
                    For lngA = 1 To UBound(arrSpecs)
                        strSpec = arrSpecs(lngA)
                        'Saves whether no bullets or numbers is specified.
                        If strSpec = "no numbers" Or strSpec = "no bullets" _
                            Or strSpec = "no numbers and no bullets" _
                            Or strSpec = "no bullets and no numbers" _
                            Or strSpec = "no numbers and bullets" _
                            Or strSpec = "no bullets and numbers" _
                            Or strSpec = "no numbers or bullets" _
                            Or strSpec = "no bullets or numbers" Then
                            For lngB = 1 To 9
                                arrList(lngB, 2) = ""
                                arrList(lngB, 4) = wdListNumberStyleNone
                            Next lngB
                        'Saves whether a tab or space follows (_, 3).
                        ElseIf strSpec = "tabs" _
                            Or strSpec = "tabs only" Or strSpec = "only tabs" _
                            Or Right(strSpec, 12) = "after bullet" _
                            Or Right(strSpec, 14) = "follows bullet" _
                            Or Right(strSpec, 12) = "after number" _
                            Or Right(strSpec, 14) = "follows number" Then
                            If Split(strSpec, " ")(0) = "one" _
                                Or Split(strSpec, " ")(0) = "a" _
                                Or Split(strSpec, " ")(0) = "only" Then
                                strSpec = Split(strSpec, " ")(1)
                            Else
                                strSpec = Split(strSpec, " ")(0)
                            End If
                            If strSpec = "tab" Or strSpec = "tabs" Then
                                dblSpec = wdTrailingTab
                            ElseIf strSpec = "space" Then
                                dblSpec = wdTrailingSpace
                            ElseIf strSpec = "nothing" Or strSpec = "no" Then
                                dblSpec = wdTrailingNone
                            End If
                            For lngB = 1 To 9
                                arrList(lngB, 3) = dblSpec
                            Next lngB
                        'Saves the font name (_, 11).
                        ElseIf Right(strSpec, 4) = "font" Then
                            If Right(strSpec, 11) = "bullet font" _
                                Or Right(strSpec, 11) = "number font" Then
                                strSpec = Left(strSpec, Len(strSpec) - 12)
                            Else
                                strSpec = Left(strSpec, Len(strSpec) - 5)
                            End If
                            If strSpec = "Body" Or strSpec = "Headings" Then
                                strSpec = "+" & strSpec
                            End If
                            For lngB = 1 To 9
                                arrList(lngB, 11) = strSpec
                            Next lngB
                        'Saves the bold spec (_, 12).
                        ElseIf strSpec = "bold bullet" _
                            Or strSpec = "bold bullets" _
                            Or strSpec = "bold number" _
                            Or strSpec = "bold numbers" Then
                            For lngB = 1 To 9
                                arrList(lngB, 12) = True
                            Next lngB
                        'Saves the italic spec (_, 13).
                        ElseIf strSpec = "italic number" _
                            Or strSpec = "italic numbers" Then
                            For lngB = 1 To 9
                                arrList(lngB, 13) = True
                            Next lngB
                        ElseIf strSpec = "bold italic number" _
                            Or strSpec = "bold italic numbers" _
                            Or strSpec = "italic bold number" _
                            Or strSpec = "italic bold numbers" _
                            Or strSpec = "bold and italic number" _
                            Or strSpec = "bold and italic numbers" _
                            Or strSpec = "italic and bold number" _
                            Or strSpec = "italic and bold numbers" Then
                            For lngB = 1 To 9
                                arrList(lngB, 12) = True
                                arrList(lngB, 13) = True
                            Next lngB
                        'Saves the color (_, 14).
                        ElseIf Right(strSpec, 5) = "color" Then
                            strSpec = Split(strSpec, " ")(0)
                            If Left(strSpec, 1) = "#" Then
                                strSpec = Right(strSpec, Len(strSpec) - 1)
                            End If
                            strSpec = Right(strSpec, 2) & Mid(strSpec, 3, 2) _
                                & Left(strSpec, 2)
                            dblSpec = Val("&H" & strSpec)
                            For lngB = 1 To 9
                                arrList(lngB, 14) = dblSpec
                            Next lngB
                        End If
                    Next lngA
                
                'If a line has style specs, saves the indents.
                ElseIf Right(arrSpecs(0), 5) = "style" Then
                    strStyle = arrSpecs(0)
                    strStyle = Left(strStyle, InStr(strStyle, "style") - 2)
                    For lngA = 1 To UBound(arrSpecs)
                        strSpec = arrSpecs(lngA)
                        dblSpec = Val(strSpec)
                        'Saves the bullet or number indent (_, 5).
                        If Right(strSpec, 13) = "bullet indent" _
                            Or Right(strSpec, 13) = "number indent" Then
                            For lngB = 1 To 9
                                If arrList(lngB, 1) = strStyle Then
                                    arrList(lngB, 5) = dblSpec
                                    Exit For
                                End If
                            Next lngB
                        'Saves the text indent (_, 7).
                        ElseIf Right(strSpec, 11) = "text indent" Then
                            For lngB = 1 To 9
                                If arrList(lngB, 1) = strStyle Then
                                    arrList(lngB, 7) = dblSpec
                                    Exit For
                                End If
                            Next lngB
                        End If
                    Next lngA
                End If
            Next objListParagraph
'Stop
'lngA = 2: Print arrList(lngA, 1) & ", " & arrList(lngA, 2) & ", " & arrList(lngA, 3) & ", " & arrList(lngA, 4) & ", " & arrList(lngA, 5) & ", " & arrList(lngA, 6) & ", " & arrList(lngA, 7) & ", " & arrList(lngA, 8) & ", " & arrList(lngA, 9) & ", " & arrList(lngA, 10) & ", " & arrList(lngA, 11) & ", " & arrList(lngA, 12) & ", " & arrList(lngA, 13) & ", " & arrList(lngA, 14) & ", " & arrList(lngA, 15) & ", " & arrList(lngA, 16) & ", " & arrList(lngA, 17) & ", " & arrList(lngA, 18) & ", " & arrList(lngA, 19) & ", " & arrList(lngA, 20)
            
            'Adds the list template if it doesn't exist.
            If sctStyleExists(strList, ActiveDocument) Then
                Set objListTemplate = ActiveDocument.ListTemplates(strList)
            Else
                Set objListTemplate = _
                    ActiveDocument.ListTemplates.Add(True, CStr(strList))
            End If
            'Applies the list template specifications.
            With objListTemplate
                For lngA = 1 To 9
                    With .ListLevels(lngA)
                        .NumberFormat = arrList(lngA, 2)
                        With .Font
                            .Name = arrList(lngA, 11)
                            If arrList(lngA, 12) <> "" Then
                                .Bold = arrList(lngA, 12)
                            End If
                            If arrList(lngA, 13) <> "" Then
                                .Italic = arrList(lngA, 13)
                            End If
                            If arrList(lngA, 14) <> "" Then
                            .Color = arrList(lngA, 14)
                            End If
                        End With
                        .TrailingCharacter = arrList(lngA, 3)
                        .NumberStyle = arrList(lngA, 4)
                        .NumberPosition = InchesToPoints(arrList(lngA, 5))
                        .Alignment = arrList(lngA, 6)
                        .TextPosition = InchesToPoints(arrList(lngA, 7))
                        .TabPosition = wdUndefined
                        .ResetOnHigher = (lngA - 1)
                        .StartAt = 1
                        .LinkedStyle = arrList(lngA, 1)
                        'The linked style name must be set after the indents.
                    End With
                Next lngA
            End With
            Set objListTemplate = Nothing
'Margins'
'-------'Sets the margins.
        ElseIf arrSpecs(0) = "Margins" Then
            For lngA = 1 To UBound(arrSpecs)
                strSpec = arrSpecs(lngA)
                dblSpec = Val(arrSpecs(lngA))
                With ActiveDocument.PageSetup
                    If InStr(strSpec, "left") <> 0 Then
                        .LeftMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpec, "right") <> 0 Then
                        .RightMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpec, "top") <> 0 Then
                        .TopMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpec, "bottom") <> 0 Then
                        .BottomMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpec, "header") <> 0 Then
                        .HeaderDistance = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpec, "footer") <> 0 Then
                        .FooterDistance = InchesToPoints(dblSpec)
                    ElseIf strSpec = "mirror margins" Then
                        .MirrorMargins = True
                    ElseIf strSpec = "no mirror margins" Then
                        .MirrorMargins = False
                    End If
                End With
            Next lngA
'Gallery'
'-------'Customizes the quick styles gallery.
        ElseIf arrSpecs(0) = "Styles gallery" Then
            'Removes the defaults.
            arrStyles = Array("Normal", "No Spacing", "Heading 1", _
                "Heading 2", "Heading 3", "Heading 4", "Heading 5", _
                "Heading 6", "Heading 7", "Heading 8", "Heading 9", _
                "Title", "Subtitle", _
                "Subtle Emphasis", "Emphasis", "Intense Emphasis", _
                "Strong", "Quote", "Intense Quote", _
                "Subtle Reference", "Intense Reference", "Book Title", _
                "List Paragraph", "Caption", "TOC Heading")
            For lngA = LBound(arrStyles) To UBound(arrStyles)
                strStyle = arrStyles(lngA)
                ActiveDocument.Styles(strStyle).QuickStyle = False
            Next lngA
            'Adds styles to the Style gallery.
            For lngA = 1 To UBound(arrSpecs)
                strStyle = arrSpecs(lngA)
                With ActiveDocument.Styles(strStyle)
                    .QuickStyle = True ' True means include in the gallery.
                    .UnhideWhenUsed = False ' False means never hidden.
                    .Visibility = False ' False (sic) means always visible.
                    .Priority = lngA
                End With
            Next lngA
        End If
    Next objParagraph
    MsgBox "Done."
End Sub

Private Sub sctDefineStyle(ByVal strStyle As String, ByVal arrSpecs As Variant)
    Dim lngType As Long, lngA As Long, strSpec As String, dblSpec As Double
    lngType = ActiveDocument.Styles(strStyle).Type
    For lngA = 1 To UBound(arrSpecs)
        strSpec = arrSpecs(lngA)
        
        With ActiveDocument.Styles(strStyle)
            If Left(strSpec, 8) = "based on" Then '------------- based on style
                strSpec = Right(strSpec, Len(strSpec) - 9)
                If strSpec = "no style" Then
                    .BaseStyle = ""
                ElseIf strStyle <> "Normal" _
                    And strStyle <> "Default Paragraph Font" Then
                    .BaseStyle = strSpec
                End If
            ElseIf Left(strSpec, 11) = "followed by" Then '---- following style
                strSpec = Right(strSpec, Len(strSpec) - 12)
                If Right(strSpec, 6) = " style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 6)
                End If
                .NextParagraphStyle = strSpec
            ElseIf strSpec = "space between" _
                Or strSpec = "add space between" _
                Or strSpec = "space between paragraphs of the same style" _
                Or strSpec = "add space between paragraphs of the same style" _
                Then '------------------------------------------ space between
                .NoSpaceBetweenParagraphsOfSameStyle = False
            ElseIf strSpec = "no space between" _
                Or strSpec = "no space between paragraphs of the same style" _
                Or strSpec = "don't add space between paragraphs" _
                Or strSpec = "don't add space between paragraphs of the same" _
                & " style" Then
                .NoSpaceBetweenParagraphsOfSameStyle = True
            End If
        End With
        
        With ActiveDocument.Styles(strStyle).Font
            If Right(strSpec, 4) = "font" Then '-------------------------- font
                strSpec = Left(strSpec, Len(strSpec) - 5)
                If strSpec = "Body" Or strSpec = "Headings" Then
                    strSpec = "+" & strSpec
                End If
                .Name = strSpec
            ElseIf Right(strSpec, 4) = "size" Then '---------------------- size
                .Size = Val(strSpec)
            ElseIf strSpec = "bold" Then '-------------------------------- bold
                .Bold = True
            ElseIf strSpec = "not bold" Or strSpec = "no bold" Then
                .Bold = False
            ElseIf strSpec = "italic" Then '---------------------------- italic
                .Italic = True
            ElseIf strSpec = "not italic" Or strSpec = "no italic" Then
                .Italic = False
            ElseIf strSpec = "small caps" Then '-------------------- small caps
                .SmallCaps = False
            ElseIf strSpec = "uppercase" Or strSpec = "all caps" Then '--- caps
                .AllCaps = False
            ElseIf Right(strSpec, 5) = "color" Then '-------------------- color
                strSpec = Split(strSpec, " ")(0)
                If Left(strSpec, 1) = "#" Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                    strSpec = Right(strSpec, 2) & Mid(strSpec, 3, 2) _
                        & Left(strSpec, 2)
                    dblSpec = Val("&H" & strSpec)
                    .Color = dblSpec
                ElseIf strSpec = "automatic color" Or strSpec = "auto color" _
                    Or strSpec = "color automatic" Or strSpec = "no color" Then
                    dblSpec = wdColorAutomatic
                    .Color = dblSpec
                ElseIf strSpec = "black color" Then
                    dblSpec = wdColorBlack
                    .Color = dblSpec
                End If
            ElseIf strSpec = "normal character spacing" Then '-- letter spacing
                .Spacing = 0
            ElseIf strSpec = "kerning" Then '-------------------------- kerning
                .Kerning = 8
            ElseIf strSpec = "no kerning" Then
                .Kerning = 0
            End If
        End With
        
        If lngType = wdStyleTypeParagraph Then
            With ActiveDocument.Styles(strStyle).ParagraphFormat
                If Right(strSpec, 11) = "left indent" Then '------- left indent
                    dblSpec = Split(strSpec, """")(0)
                    .LeftIndent = InchesToPoints(dblSpec)
                ElseIf Right(strSpec, 12) = "right indent" Then '- right indent
                    dblSpec = Split(strSpec, """")(0)
                    .RightIndent = InchesToPoints(dblSpec)
                ElseIf Right(strSpec, 6) = "before" _
                    And strSpec <> "page break before" Then '----- space before
                    dblSpec = Split(strSpec, " ")(0)
                    .SpaceBefore = dblSpec
                ElseIf Right(strSpec, 5) = "after" Then '---------- space after
                    dblSpec = Split(strSpec, " ")(0)
                    .SpaceAfter = dblSpec
                ElseIf Right(strSpec, 12) = "line spacing" Then '- line spacing
                    If Split(strSpec, " ")(1) = "pt" _
                        Or Split(strSpec, " ")(1) = "pt." Then
                        dblSpec = Split(strSpec, " ")(0)
                        .LineSpacingRule = wdLineSpaceExactly
                        .LineSpacing = dblSpec
                    ElseIf Split(strSpec, " ")(1) = "least" Then
                        dblSpec = Split(strSpec, " ")(2)
                        .LineSpacingRule = wdLineSpaceAtLeast
                        .LineSpacing = dblSpec
                    Else
                        dblSpec = Split(strSpec, " ")(0)
                        .LineSpacingRule = wdLineSpaceMultiple
                        .LineSpacing = LinesToPoints(dblSpec)
                    End If
                ElseIf strSpec = "left aligned" Or strSpec = "right aligned" _
                    Or strSpec = "centered" Or strSpec = "center" _
                    Or strSpec = "center align" _
                    Or strSpec = "justified" Or strSpec = "justify" _
                    Then '------------------------------------------- alignment
                    If strSpec = "left aligned" Then
                        dblSpec = wdAlignParagraphLeft
                    ElseIf strSpec = "right aligned" Then
                        dblSpec = wdAlignParagraphRight
                    ElseIf strSpec = "centered" Or strSpec = "center" _
                        Or strSpec = "center align" Then
                        dblSpec = wdAlignParagraphCenter
                    ElseIf strSpec = "justified" Or strSpec = "justify" Then
                        dblSpec = wdAlignParagraphJustify
                    End If
                    .Alignment = dblSpec
                ElseIf strSpec = "widow/orphan control" _
                    Or strSpec = "orphan/widow control" _
                    Or strSpec = "widow and orphan control" _
                    Or strSpec = "orphan and widow control" _
                    Or strSpec = "widow control" _
                    Or strSpec = "orphan control" Then '--------- widow control
                    .WidowControl = True
                ElseIf strSpec = "no widow/orphan control" _
                    Or strSpec = "no orphan/widow control" _
                    Or strSpec = "no widow and orphan control" _
                    Or strSpec = "no orphan and widow control" _
                    Or strSpec = "no widow or orphan control" _
                    Or strSpec = "no orphan or widow control" _
                    Or strSpec = "no widow control" _
                    Or strSpec = "no orphan control" Then
                    .WidowControl = False
                ElseIf strSpec = "keep with next" Then '-------- keep with next
                    .KeepWithNext = True
                ElseIf strSpec = "don't keep with next" _
                    Or strSpec = "do not keep with next" Then
                    .KeepWithNext = False
                ElseIf strSpec = "keep lines together" _
                    Or strSpec = "keep together" Then '---- keep lines together
                    .KeepTogether = True
                ElseIf strSpec = "don't keep together" _
                    Or strSpec = "don't keep lines together" _
                    Or strSpec = "do not keep together" _
                    Or strSpec = "do not keep lines together" Then
                    .KeepTogether = False
                ElseIf strSpec = "page break before" Then '-- page break before
                    .PageBreakBefore = True
                ElseIf strSpec = "no page break before" Then
                    .PageBreakBefore = False
                End If
            End With
        End If
    Next lngA
End Sub

Private Function sctStyleExists(ByVal strStyle As String, _
    ByVal objDocument As Word.Document) As Boolean
'Source: http://www.vbaexpress.com/forum/showthread.php?15259-Solved-How-to-check-if-a-Word-Style-exists
'Source: https://roxtonlabs.blogspot.com/2015/09/vba-test-if-style-exists-in-word.html
    Dim objStyle As Word.Style, objListTemplate As Word.ListTemplate
    On Error Resume Next
    Set objStyle = objDocument.Styles(strStyle)
    sctStyleExists = Not objStyle Is Nothing
    If Not sctStyleExists Then
        Set objListTemplate = objDocument.ListTemplates(strStyle)
        sctStyleExists = Not objListTemplate Is Nothing
    End If
End Function
```
