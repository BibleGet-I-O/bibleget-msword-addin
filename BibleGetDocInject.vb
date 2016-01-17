Imports System.ComponentModel
Imports Newtonsoft.Json.Linq
Imports System.Drawing
Imports System.Text.RegularExpressions


Public Class BibleGetDocInject

    Private Application As Word.Application = Globals.ThisAddIn.Application
    Private worker As BackgroundWorker

    Public Sub New(ByRef worker As BackgroundWorker)
        Me.worker = worker
    End Sub

    Public Function InsertTextAtCurrentSelection(ByVal myString As String) As String
        Dim currentSelection As Word.Selection = Application.Selection

        'Dim rng As Word.Range = currentSelection.Range
        ' Store the user's current Overtype selection
        Dim userOvertype As Boolean = Application.Options.Overtype

        ' Make sure Overtype is turned off.
        If Application.Options.Overtype Then
            Application.Options.Overtype = False
        End If

        Dim paragraphFont As Word.Font = wdFontConverter(My.Settings.BookChapterFont)
        Dim bookChapterFont As Word.Font = wdFontConverter(My.Settings.BookChapterFont)
        Dim verseNumberFont As Word.Font = wdFontConverter(My.Settings.VerseNumberFont)
        Dim verseTextFont As Word.Font = wdFontConverter(My.Settings.VerseTextFont)
        Dim noVersionFormatting As Boolean = My.Settings.NOVERSIONFORMATTING

        Dim jsObj As JToken = JObject.Parse(myString)
        Dim jRRArray As JArray = jsObj.SelectToken("results")

        Dim prevversion As String = String.Empty
        Dim newversion As Boolean
        Dim prevbook As String = String.Empty
        Dim newbook As Boolean
        Dim prevchapter As Integer = -1
        Dim newchapter As Boolean
        Dim prevverse As String = String.Empty
        Dim newverse As Boolean

        Dim currentversion As String
        Dim currentbook As String
        Dim currentchapter As Integer
        Dim currentverse As String

        Dim firstversion As Boolean = True
        Dim firstchapter As Boolean = True

        Dim firstVerse As Boolean = False
        Dim normalText As Boolean = False

        Dim workerProgress As Integer = 20
        For Each currentJson As JToken In jRRArray

            worker.ReportProgress(workerProgress)
            currentbook = currentJson.SelectToken("book").Value(Of String)()
            currentchapter = currentJson.SelectToken("chapter").Value(Of Integer)()
            currentverse = currentJson.SelectToken("verse").Value(Of String)()
            currentversion = currentJson.SelectToken("version").Value(Of String)()

            If Not currentverse = prevverse Then
                newverse = True
                prevverse = currentverse
            Else
                newverse = False
            End If

            If Not currentchapter = prevchapter Then
                newchapter = True
                newverse = True
                prevchapter = currentchapter
            Else
                newchapter = False
            End If

            If Not currentbook = prevbook Then
                newbook = True
                newchapter = True
                newverse = True
                prevbook = currentbook
            Else
                newbook = False
            End If

            If Not currentversion = prevversion Then
                newversion = True
                newbook = True
                newchapter = True
                newverse = True
                prevversion = currentversion
            Else
                newversion = False
            End If

            setParagraphStyles(currentSelection, paragraphFont)

            If newversion Then
                firstVerse = True

                firstchapter = True
                If firstversion Then
                    firstversion = False
                Else
                    TypeText(currentSelection, "NEWLINE")
                End If

                setVersionStyles(currentSelection, paragraphFont)

                TypeText(currentSelection, currentversion)
                TypeText(currentSelection, "NEWLINE")
            End If

            If newbook Or newchapter Then
                '//System.out.println(currentbook+" "+currentchapter);
                firstVerse = True

                If firstchapter Then
                    firstchapter = False
                Else
                    TypeText(currentSelection, "NEWLINE")
                End If

                setBookChapterStyles(currentSelection, bookChapterFont)
                TypeText(currentSelection, currentbook + " " + currentchapter.ToString)
                TypeText(currentSelection, "NEWLINE")
                setParagraphAlignment(currentSelection)
            End If

            If newverse Then
                normalText = False
                '//System.out.print("\n"+currentverse);
                setVerseNumberStyles(currentSelection, verseNumberFont)
                TypeText(currentSelection, " " + currentverse)
            End If

            setVerseTextStyles(currentSelection, verseTextFont)
            Dim currentText As String = currentJson.SelectToken("text").Value(Of String)()
            currentText = currentText.Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
            Dim remainingText As String = currentText

            If Regex.IsMatch(currentText, ".*<[/]{0,1}(?:speaker|sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>.*", RegexOptions.Singleline) Then '//[/]{0,1}(?:sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>
                'Diagnostics.Debug.WriteLine("We have detected a text string with special formatting: " + currentText)
                Dim pattern1 As String = "(.*?)<((speaker|sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1})>(.*?)</\2>"
                'Matcher matcher1 = pattern1.matcher(currentText);
                Dim iteration As Integer = 0
                Dim currentSpaceAfter As Single = currentSelection.Range.ParagraphFormat.SpaceAfter
                For Each match As Match In Regex.Matches(currentText, pattern1, RegexOptions.Singleline)
                    '//                    System.out.print("Iteration ");
                    '//                    System.out.println(++iteration);
                    '//                    System.out.println("group1:"+matcher1.group(1));
                    '//                    System.out.println("group2:"+matcher1.group(2));
                    '//                    System.out.println("group3:"+matcher1.group(3));
                    '//                    System.out.println("group4:"+matcher1.group(4));
                    If match.Groups(1).Value IsNot Nothing And match.Groups(1).Value IsNot String.Empty Then
                        'Diagnostics.Debug.WriteLine("We seem to have some normal text preceding our special formatting text: " + match.Groups(1).Value)
                        normalText = True
                        TypeText(currentSelection, match.Groups(1).Value)
                        Dim regex As Regex = New Regex(match.Groups(1).Value)
                        remainingText = regex.Replace(remainingText, String.Empty, 1)
                        'remainingText.replaceFirst(match.Groups(1).Value, "")
                    End If

                    If match.Groups(4).Value IsNot Nothing And match.Groups(4).Value IsNot String.Empty Then

                        Dim matchedTag As String = match.Groups(2).Value
                        Dim formattingTagContents As String = match.Groups(4).Value

                        '//check for nested speaker tags!
                        Dim nestedTag As Boolean = False
                        Dim speakerTagBefore As String = ""
                        Dim speakerTagContents As String = ""
                        Dim speakerTagAfter As String = ""

                        If Regex.IsMatch(formattingTagContents, ".*<[/]{0,1}speaker>.*", RegexOptions.Singleline) Then
                            nestedTag = True
                            'Diagnostics.Debug.WriteLine("We have a nested tag in this special formatting text: " + formattingTagContents)
                            Dim remainingText2 As String = formattingTagContents

                            'Matcher matcher2 = pattern1.matcher(formattingTagContents);
                            Dim iteration2 As Integer = 0
                            For Each matcher2 As Match In Regex.Matches(formattingTagContents, pattern1)
                                If matcher2.Groups(2).Value IsNot Nothing And matcher2.Groups(2).Value IsNot String.Empty And matcher2.Groups(2).Value = "speaker" Then
                                    If matcher2.Groups(1).Value IsNot Nothing And matcher2.Groups(1).Value IsNot String.Empty Then
                                        speakerTagBefore = matcher2.Groups(1).Value
                                        Dim reggaeton As Regex = New Regex(matcher2.Groups(1).Value)
                                        remainingText2 = reggaeton.Replace(remainingText2, String.Empty, 1)
                                    End If
                                    speakerTagContents = matcher2.Groups(4).Value
                                    Dim reggae As Regex = New Regex("<" + matcher2.Groups(2).Value + ">" + matcher2.Groups(4).Value + "</" + matcher2.Groups(2).Value + ">")
                                    speakerTagAfter = reggae.Replace(remainingText2, String.Empty, 1)
                                End If
                            Next
                        End If

                        If noVersionFormatting Then formattingTagContents = " " + formattingTagContents + " "

                        Select Case matchedTag
                            Case "pof"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                normalText = False
                            Case "pos"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                normalText = False
                            Case "poif"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                normalText = False
                            Case "po"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                normalText = False
                            Case "poi"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 1)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                End If
                                normalText = False
                            Case "pol"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 1)
                                End If
                                normalText = False
                            Case "poil"
                                If Not noVersionFormatting Then
                                    If firstVerse = False And normalText = True Then TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent + 1)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not noVersionFormatting Then
                                    TypeText(currentSelection, "NEWLINE")
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setVerseTextStyles(currentSelection, verseTextFont)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent)
                                End If
                                normalText = False
                            Case "sm"
                                Dim smallCaps As String = match.Groups(4).Value.ToLower
                                '//System.out.println("SMALLCAPSIZE THIS TEXT: "+smallCaps);
                                currentSelection.Font.SmallCaps = True
                                TypeText(currentSelection, smallCaps)
                                currentSelection.Font.SmallCaps = False
                            Case "speaker"
                                '//                                System.out.println("We have found a speaker tag");
                                currentSelection.Font.Bold = True
                                '//xPropertySet.setPropertyValue("CharBackTransparent", false);
                                'xPropertySet.setPropertyValue("CharBackColor", Color.LIGHT_GRAY.getRGB() & ~0xFF000000);
                                currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30
                                TypeText(currentSelection, match.Groups(4).Value)
                                If Not My.Settings.VerseTextFont.Bold Then
                                    currentSelection.Font.Bold = False
                                End If
                                'xPropertySet.setPropertyValue("CharBackColor", bgColorVerseText.getRGB() & ~0xFF000000);
                                currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Microsoft.Office.Interop.Word.WdColor)
                        End Select

                        Dim nonmereggaepiu As Regex = New Regex("<" + match.Groups(2).Value + ">" + match.Groups(4).Value + "</" + match.Groups(2).Value + ">")
                        remainingText = nonmereggaepiu.Replace(remainingText, String.Empty, 1)
                    End If
                Next
                currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                '//                System.out.println("We have a match for special formatting: "+currentText);
                '//                System.out.println("And after elaborating our matches, this is what we have left: "+remainingText);
                '//                System.out.println();
                If remainingText IsNot String.Empty Then
                    'Diagnostics.Debug.WriteLine("We have a fragment of text left over after all this: " + remainingText)
                    TypeText(currentSelection, remainingText)
                End If
                '/*
                'Pattern pattern2 = Pattern.compile("([.^>]+)$",Pattern.UNICODE_CHARACTER_CLASS | Pattern.DOTALL);
                'Matcher matcher2 = pattern2.matcher(currentText);
                '//String lastPiece = currentText.
                '                                                                                                                                                        If (matcher2.find()) Then
                '{
                '    if(matcher2.group(1) != null && !"".equals(matcher2.group(1)))
                '    { 
                '        m_xText.insertString(xTextRange, matcher2.group(1), false);
                '    }
                '}
                '*/

            Else

                normalText = True
                '//System.out.println("No match for special case formatting here: "+currentText);
                '// set properties of text change based on user preferences
                '//setVerseTextStyles(xPropertySet);
                TypeText(currentSelection, currentText)
            End If

            If firstVerse Then firstVerse = False

            workerProgress += 1
        Next


        ' Restore the user's Overtype selection
        Application.Options.Overtype = userOvertype

        Return "Ok! All done!"
    End Function

    Private Sub TypeText(ByVal currentSelection As Microsoft.Office.Interop.Word.Selection, ByVal myString As String)
        With currentSelection

            ' Test to see if selection is an insertion point.
            If .Type = Word.WdSelectionType.wdSelectionIP Then
                If myString = "NEWLINE" Then
                    .TypeParagraph()
                Else
                    .TypeText(myString)
                End If

            ElseIf .Type = Word.WdSelectionType.wdSelectionNormal Then
                ' Move to start of selection.
                If Application.Options.ReplaceSelection Then
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
                End If
                If myString = "NEWLINE" Then
                    .TypeParagraph()
                Else
                    .TypeText(myString)
                End If

            Else
                ' Do nothing.
            End If
        End With
    End Sub

    Private Function wdFontConverter(ByVal myFont As Drawing.Font) As Word.Font
        Dim returnFont As Word.Font = New Word.Font
        If myFont IsNot Nothing Then
            returnFont.Name = myFont.Name
            returnFont.Bold = myFont.Bold
            returnFont.Italic = myFont.Italic
            If myFont.Underline Then
                returnFont.Underline = Word.WdUnderline.wdUnderlineSingle
            Else
                returnFont.Underline = Word.WdUnderline.wdUnderlineNone
            End If            
            returnFont.StrikeThrough = myFont.Strikeout
            returnFont.Size = myFont.Size
        Else
            returnFont.Name = "Times New Roman"
            returnFont.Bold = False
            returnFont.Italic = False
            returnFont.Underline = Word.WdUnderline.wdUnderlineNone
            returnFont.StrikeThrough = False
            returnFont.Size = 12
        End If

        Return returnFont
    End Function

    Private Sub setParagraphStyles(ByRef currentSelection As Word.Selection, ByVal paragraphFont As Word.Font)
        'currentSelection.Font.Name = paragraphFont.Name
        'currentSelection.Font.Size = paragraphFont.Size
        'currentSelection.Font.Bold = paragraphFont.Bold
        'currentSelection.Font.Italic = paragraphFont.Italic
        'currentSelection.Font.Underline = paragraphFont.Underline
        'currentSelection.Font.StrikeThrough = paragraphFont.StrikeThrough

        currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.Indent)
        'currentSelection.Range.ParagraphFormat.FirstLineIndent = My.Settings.Indent
        'Diagnostics.Debug.WriteLine("current linespacing = " + My.Settings.Linespacing.ToString)
        Select Case My.Settings.Linespacing
            Case 1.0
                'Diagnostics.Debug.WriteLine("current linespacing has been detected as Single")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
            Case 1.5
                'Diagnostics.Debug.WriteLine("current linespacing has been detected as one and a half")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
            Case 2.0
                'Diagnostics.Debug.WriteLine("current linespacing has been detected as Double")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble
        End Select
    End Sub

    Private Sub setVersionStyles(ByRef currentSelection As Word.Selection, ByVal paragraphFont As Word.Font)
        currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        currentSelection.Font.Name = paragraphFont.Name
        currentSelection.Font.Size = paragraphFont.Size
        currentSelection.Font.Bold = paragraphFont.Bold
        currentSelection.Font.Italic = paragraphFont.Italic
        currentSelection.Font.Underline = paragraphFont.Underline
        currentSelection.Font.StrikeThrough = paragraphFont.StrikeThrough
        'currentSelection.Font.Color = CType(ColorTranslator.ToOle(My.Settings.BookChapterForeColor), Microsoft.Office.Interop.Word.WdColor)
        'currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.BookChapterBackColor), Microsoft.Office.Interop.Word.WdColor)
        Select Case My.Settings.BookChapterVAlign
            Case "sub"
                currentSelection.Font.Subscript = True
            Case "super"
                currentSelection.Font.Superscript = True
            Case "baseline"
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
        End Select
    End Sub

    Private Sub setBookChapterStyles(ByRef currentSelection As Word.Selection, ByVal bookChapterFont As Word.Font)
        currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

        currentSelection.Font.Name = bookChapterFont.Name
        currentSelection.Font.Size = bookChapterFont.Size
        currentSelection.Font.Bold = bookChapterFont.Bold
        currentSelection.Font.Italic = bookChapterFont.Italic
        currentSelection.Font.Underline = bookChapterFont.Underline
        currentSelection.Font.StrikeThrough = bookChapterFont.StrikeThrough
        currentSelection.Font.Color = CType(ColorTranslator.ToOle(My.Settings.BookChapterForeColor), Microsoft.Office.Interop.Word.WdColor)
        If My.Settings.BookChapterBackColor = Nothing Then
            currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic
        Else
            currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.BookChapterBackColor), Microsoft.Office.Interop.Word.WdColor)
        End If
        Select Case My.Settings.BookChapterVAlign
            Case "sub"
                currentSelection.Font.Subscript = True
            Case "super"
                currentSelection.Font.Superscript = True
            Case "baseline"
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
        End Select

    End Sub

    Private Sub setParagraphAlignment(ByRef currentSelection)
        Select Case My.Settings.ParagraphAlignment
            Case "left"
                currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            Case "right"
                currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            Case "center"
                currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            Case "justify"
                currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
        End Select
    End Sub

    Private Sub setVerseNumberStyles(ByRef currentSelection As Word.Selection, ByVal verseNumberFont As Word.Font)
        currentSelection.Font.Name = verseNumberFont.Name
        currentSelection.Font.Size = verseNumberFont.Size
        currentSelection.Font.Bold = verseNumberFont.Bold
        currentSelection.Font.Italic = verseNumberFont.Italic
        currentSelection.Font.Underline = verseNumberFont.Underline
        currentSelection.Font.StrikeThrough = verseNumberFont.StrikeThrough
        currentSelection.Font.Color = CType(ColorTranslator.ToOle(My.Settings.VerseNumberForeColor), Microsoft.Office.Interop.Word.WdColor)
        If My.Settings.VerseNumberBackColor = Nothing Then
            currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic
        Else
            currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.VerseNumberBackColor), Microsoft.Office.Interop.Word.WdColor)
        End If
        Select Case My.Settings.VerseNumberVAlign
            Case "sub"
                currentSelection.Font.Subscript = True
            Case "super"
                currentSelection.Font.Superscript = True
            Case "baseline"
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
        End Select
    End Sub

    Private Sub setVerseTextStyles(ByRef currentSelection As Word.Selection, ByVal verseTextFont As Word.Font)
        currentSelection.Font.Name = verseTextFont.Name
        currentSelection.Font.Size = verseTextFont.Size
        currentSelection.Font.Bold = verseTextFont.Bold
        currentSelection.Font.Italic = verseTextFont.Italic
        currentSelection.Font.Underline = verseTextFont.Underline
        currentSelection.Font.StrikeThrough = verseTextFont.StrikeThrough
        currentSelection.Font.Color = CType(ColorTranslator.ToOle(My.Settings.VerseTextForeColor), Microsoft.Office.Interop.Word.WdColor)
        If My.Settings.VerseTextBackColor = Nothing Then
            currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic
        Else
            currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Microsoft.Office.Interop.Word.WdColor)
        End If
        Select Case My.Settings.VerseTextVAlign
            Case "sub"
                currentSelection.Font.Subscript = True
            Case "super"
                currentSelection.Font.Superscript = True
            Case "baseline"
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
        End Select
    End Sub

    Private Sub insertNestedSpeakerTag(ByVal speakerTagBefore As String, ByVal speakerTagContents As String, ByVal speakerTagAfter As String, ByRef currentSelection As Word.Selection)


        '//        System.out.println("We are now working with a nested Speaker Tag."); //Using BG="+grayBG.getRGB()+"=R("+r+"),B("+b+"),G("+g+")
        '//        System.out.println("speakerTagBefore=<"+speakerTagBefore+">,speakerTagContents=<"+speakerTagContents+">,speakerTagAfter=<"+speakerTagAfter+">");
        TypeText(currentSelection, speakerTagBefore)
        currentSelection.Font.Bold = True

        'xPropertySet.setPropertyValue("CharBackColor", Color.LIGHT_GRAY.getRGB() & ~0xFF000000);
        currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30

        TypeText(currentSelection, " " + speakerTagContents + " ")

        If Not My.Settings.VerseTextFont.Bold Then
            currentSelection.Font.Bold = False
        End If
        'xPropertySet.setPropertyValue("CharBackColor", bgColorVerseText.getRGB() & ~0xFF000000);
        currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Microsoft.Office.Interop.Word.WdColor)
        TypeText(currentSelection, speakerTagAfter)
    End Sub


End Class
