﻿Imports System.ComponentModel
Imports Newtonsoft.Json.Linq
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Collections

Public Class BibleGetDocInject

    Private DEBUG_MODE As Boolean
    Private Application As Word.Application = Globals.BibleGetAddIn.Application
    Private worker As BackgroundWorker
    Private e As DoWorkEventArgs

    Public Sub New(ByRef worker As BackgroundWorker, ByRef eventArgs As System.ComponentModel.DoWorkEventArgs)
        Me.worker = worker
        e = eventArgs
        DEBUG_MODE = My.Settings.DEBUG_MODE
    End Sub

    'Compared to Google Docs, Microsoft Word does not return a reference to a paragraph when creating it
    'Instead it moves the cursor into the new paragraph and you simply set styles on the new position in the document
    Public Function InsertTextAtCurrentSelection(ByVal myStr As String) As String
        Dim currentSelection As Word.Selection = Application.Selection
        Dim BibleVersionStack = New ArrayList
        Dim BookChapterStack = New ArrayList
        Dim L10NBookNames = New LocalizedBibleBooks

        'Backup current styles, will restore after
        Dim currentStyle As Word.Font = New Word.Font
        currentStyle.Bold = currentSelection.Font.Bold
        currentStyle.Italic = currentSelection.Font.Italic
        currentStyle.Name = currentSelection.Font.Name
        currentStyle.Underline = currentSelection.Font.Underline
        currentStyle.StrikeThrough = currentSelection.Font.StrikeThrough
        currentStyle.Size = currentSelection.Font.Size
        currentStyle.Color = currentSelection.Font.Color
        currentStyle.Shading.BackgroundPatternColor = currentSelection.Font.Shading.BackgroundPatternColor

        'Dim rng As Word.Range = currentSelection.Range
        ' Store the user's current Overtype selection
        Dim userOvertype As Boolean = Application.Options.Overtype

        ' Make sure Overtype is turned off.
        If userOvertype Then
            Application.Options.Overtype = False
        End If

        'Dim noVersionFormatting As Boolean = My.Settings.NOVERSIONFORMATTING

        Dim jsObj As JToken = JObject.Parse(myStr)
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
        Dim currentbookabbrev As String
        Dim currentbookUnivIdx As Integer
        Dim currentchapter As Integer
        Dim currentverse As String
        Dim originalquery As String

        Dim firstversion As Boolean = True
        Dim firstChapter As Boolean = True

        Dim firstVerse As Boolean = False
        Dim normalText As Boolean = False

        Dim workerProgress As Integer = 20
        For Each currentJson As JToken In jRRArray
            If worker.CancellationPending Then
                e.Cancel = True
                Return "Work was cancelled"
            End If
            worker.ReportProgress(workerProgress)
            currentbook = currentJson.SelectToken("book").Value(Of String)()
            currentbookabbrev = currentJson.SelectToken("bookabbrev").Value(Of String)()
            currentbookUnivIdx = currentJson.SelectToken("univbooknum").Value(Of Integer)()
            currentchapter = currentJson.SelectToken("chapter").Value(Of Integer)()
            currentverse = currentJson.SelectToken("verse").Value(Of String)()
            currentversion = currentJson.SelectToken("version").Value(Of String)()
            originalquery = currentJson.SelectToken("originalquery").Value(Of String)()

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

            If newversion Then
                firstVerse = True
                'firstChapter = True

                Select Case My.Settings.BibleVersionWrap
                    Case WRAP.PARENTHESES
                        currentversion = "(" & currentversion & ")"
                    Case WRAP.BRACKETS
                        currentversion = "[" & currentversion & "]"
                End Select


                'so here's a case for you: Bible Version on top, Book/Chapter on the bottom
                'Book/Chapter however was winding up below the next Version instead of before
                'so when Version is top and Book/Chapter is bottom we have to output Book/Chapter first

                If My.Settings.BibleVersionVisibility = VISIBILITY.SHOW Then
                    If BookChapterStack.Count > 0 Then
                        Select Case My.Settings.BookChapterPosition
                            Case POS.BOTTOM
                                CreateNewPar(currentSelection)                   'create a new paragraph
                                setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            Case POS.TOP
                                CreateNewPar(currentSelection)                   'create a new paragraph
                                setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            Case POS.BOTTOMINLINE
                                TypeText(currentSelection, " ")
                        End Select
                        setTextStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                        TypeText(currentSelection, BookChapterStack.Item(0))   'insert the Bible version into the paragraph
                        BookChapterStack.RemoveAt(0)
                    End If
                    Select Case My.Settings.BibleVersionPosition
                        Case POS.BOTTOM
                            BibleVersionStack.Add(currentversion)
                            If BibleVersionStack.Count > 1 Then
                                CreateNewPar(currentSelection)                   'create a new paragraph
                                setParagraphStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                                setTextStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                                TypeText(currentSelection, BibleVersionStack.Item(0))   'insert the Bible version into the paragraph
                                BibleVersionStack.RemoveAt(0)
                            End If
                        Case POS.TOP
                            If firstversion = False Then
                                CreateNewPar(currentSelection)
                            Else
                                firstversion = False
                            End If
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                            setTextStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                            TypeText(currentSelection, currentversion)
                    End Select
                End If
            End If

            If newbook Or newchapter Then
                '//System.out.println(currentbook+" "+currentchapter);
                Dim bkChStr As String = ""
                Select Case My.Settings.BookChapterFormat
                    Case FORMAT.BIBLELANG
                        bkChStr = currentbook + " " + currentchapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.BIBLELANGABBREV
                        bkChStr = currentbookabbrev + " " + currentchapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.USERLANG
                        bkChStr = L10NBookNames.GetBookByIndex(currentbookUnivIdx - 1).Fullname + " " + currentchapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.USERLANGABBREV
                        bkChStr = L10NBookNames.GetBookByIndex(currentbookUnivIdx - 1).Abbrev + " " + currentchapter.ToString(CultureInfo.InvariantCulture)
                End Select
                If My.Settings.BookChapterFullReference Then
                    'retrieve the original query from originalquery property in the json response received
                    If Regex.IsMatch(originalquery, "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}", RegexOptions.Singleline) Then
                        bkChStr &= Regex.Replace(originalquery, "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}", "")
                    Else
                        bkChStr &= Regex.Replace(originalquery, "^[1-9][0-9]{0,2}", "")
                    End If
                End If

                Select Case My.Settings.BookChapterWrap
                    Case WRAP.PARENTHESES
                        bkChStr = "(" & bkChStr & ")"
                    Case WRAP.BRACKETS
                        bkChStr = "[" & bkChStr & "]"
                End Select

                firstVerse = True

                Select Case My.Settings.BookChapterPosition
                    Case POS.BOTTOM
                        BookChapterStack.Add(bkChStr)
                        If BookChapterStack.Count > 1 Then
                            CreateNewPar(currentSelection)
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            setTextStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            TypeText(currentSelection, BookChapterStack.Item(0))
                            BookChapterStack.RemoveAt(0)
                        End If
                    Case POS.TOP
                        If firstChapter = False Then
                            CreateNewPar(currentSelection)
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                        ElseIf firstChapter = True And (My.Settings.BibleVersionVisibility = VISIBILITY.HIDE Or My.Settings.BibleVersionPosition = POS.BOTTOM) Then
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            firstChapter = False
                        Else
                            CreateNewPar(currentSelection)
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            firstChapter = False
                        End If
                        setTextStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                        TypeText(currentSelection, bkChStr)
                    Case POS.BOTTOMINLINE
                        BookChapterStack.Add(bkChStr)
                        If BookChapterStack.Count > 1 Then
                            TypeText(currentSelection, " ")
                            setTextStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
                            TypeText(currentSelection, BookChapterStack.Item(0))
                            BookChapterStack.RemoveAt(0)
                        End If
                End Select
            End If

            If newverse Then
                normalText = False
                '//System.out.print("\n"+currentverse);
                If firstVerse Then
                    CreateNewPar(currentSelection)
                    setParagraphStyles(currentSelection, PARAGRAPHTYPE.VERSES)
                    firstVerse = False
                End If
                setTextStyles(currentSelection, PARAGRAPHTYPE.VERSENUMBER)
                TypeText(currentSelection, " " + currentverse)
            End If

            setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
            Dim currentText As String = currentJson.SelectToken("text").Value(Of String)()
            currentText = currentText.Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
            Dim remainingText As String = currentText

            If Regex.IsMatch(currentText, ".*<[/]{0,1}(?:speaker|sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>.*", RegexOptions.Singleline) Then '//[/]{0,1}(?:sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have detected a text string with special formatting: " + currentText)
                Dim pattern1 As String = "(.*?)<((speaker|sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1})>(.*?)</\2>"
                'Matcher matcher1 = pattern1.matcher(currentText);
                'Dim iteration As Integer = 0
                Dim currentSpaceAfter As Single = currentSelection.Range.ParagraphFormat.SpaceAfter
                For Each match As Match In Regex.Matches(currentText, pattern1, RegexOptions.Singleline)
                    '//                    System.out.print("Iteration ");
                    '//                    System.out.println(++iteration);
                    '//                    System.out.println("group1:"+matcher1.group(1));
                    '//                    System.out.println("group2:"+matcher1.group(2));
                    '//                    System.out.println("group3:"+matcher1.group(3));
                    '//                    System.out.println("group4:"+matcher1.group(4));
                    If match.Groups(1).Value IsNot Nothing And match.Groups(1).Value IsNot String.Empty Then
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We seem to have some normal text preceding our special formatting text: " + match.Groups(1).Value)
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
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have a nested tag in this special formatting text: " + formattingTagContents)
                            Dim remainingText2 As String = formattingTagContents

                            'Matcher matcher2 = pattern1.matcher(formattingTagContents);
                            'Dim iteration2 As Integer = 0
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

                        If My.Settings.NOVERSIONFORMATTING Then formattingTagContents = " " + formattingTagContents + " "

                        Select Case matchedTag
                            Case "pof"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                normalText = False
                            Case "pos"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                normalText = False
                            Case "poif"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                normalText = False
                            Case "po"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                normalText = False
                            Case "poi"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 1)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                normalText = False
                            Case "pol"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+400);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 1)
                                End If
                                normalText = False
                            Case "poil"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 1)
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedSpeakerTag(speakerTagBefore, speakerTagContents, speakerTagAfter, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent)
                                End If
                                normalText = False
                            Case "sm"
                                Dim smallCaps As String = match.Groups(4).Value.ToLower(CultureInfo.CurrentUICulture)
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
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have a fragment of text left over after all this: " + remainingText)
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

        If BookChapterStack.Count > 0 Then
            If My.Settings.BookChapterPosition = POS.BOTTOM Then
                CreateNewPar(currentSelection)
                setParagraphStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
            Else
                TypeText(currentSelection, " ")
            End If
            setTextStyles(currentSelection, PARAGRAPHTYPE.BOOKCHAPTER)
            TypeText(currentSelection, BookChapterStack.Item(0))
            BookChapterStack.RemoveAt(0)
        End If

        If BibleVersionStack.Count > 0 Then
            CreateNewPar(currentSelection)
            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
            setTextStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
            TypeText(currentSelection, BibleVersionStack.Item(0))
            BibleVersionStack.RemoveAt(0)
        End If

        CreateNewPar(currentSelection) 'one last paragraph
        'Restore original text styles
        currentSelection.Font.Bold = currentStyle.Bold
        currentSelection.Font.Italic = currentStyle.Italic
        currentSelection.Font.Name = currentStyle.Name
        currentSelection.Font.Underline = currentStyle.Underline
        currentSelection.Font.StrikeThrough = currentStyle.StrikeThrough
        currentSelection.Font.Size = currentStyle.Size
        currentSelection.Font.Color = currentStyle.Color
        currentSelection.Font.Shading.BackgroundPatternColor = currentStyle.Shading.BackgroundPatternColor

        ' Restore the user's Overtype selection
        If userOvertype = True Then Application.Options.Overtype = True

        Return "Ok! All done!"
    End Function

    Private Sub CreateNewPar(ByVal currentSelection As Word.Selection)
        With currentSelection
            If .Type = Word.WdSelectionType.wdSelectionIP Then
                .TypeParagraph()
            ElseIf .Type = Word.WdSelectionType.wdSelectionNormal Then
                ' Move to start of selection.
                If Application.Options.ReplaceSelection Then
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
                End If
                .TypeParagraph()
            Else
                ' Do nothing.
            End If
        End With
    End Sub

    Private Sub TypeText(ByVal currentSelection As Word.Selection, ByVal myString As String)
        With currentSelection
            ' Test to see if selection is an insertion point.
            If .Type = Word.WdSelectionType.wdSelectionIP Then
                .TypeText(myString)
            ElseIf .Type = Word.WdSelectionType.wdSelectionNormal Then
                ' Move to start of selection.
                If Application.Options.ReplaceSelection Then
                    .Collapse(Direction:=Word.WdCollapseDirection.wdCollapseStart)
                End If
                .TypeText(myString)
            Else
                ' Do nothing.
            End If
        End With
    End Sub

    Private Shared Function wdFontConverter(ByVal myFont As Drawing.Font) As Word.Font
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
            'returnFont.StrikeThrough = myFont.Strikeout
            returnFont.Size = myFont.Size
        Else
            returnFont.Name = "Times New Roman"
            returnFont.Bold = False
            returnFont.Italic = False
            returnFont.Underline = Word.WdUnderline.wdUnderlineNone
            'returnFont.StrikeThrough = False
            returnFont.Size = 12
        End If

        Return returnFont
    End Function

    Private Sub setParagraphStyles(ByRef currentSelection As Word.Selection, ByVal PARTYPE As PARAGRAPHTYPE)
        Select Case Application.Options.MeasurementUnit
            Case Word.WdMeasurementUnits.wdInches
                currentSelection.Range.ParagraphFormat.LeftIndent = Application.InchesToPoints(My.Settings.LeftIndent)
                currentSelection.Range.ParagraphFormat.RightIndent = Application.InchesToPoints(My.Settings.RightIndent)
            Case Word.WdMeasurementUnits.wdCentimeters
                currentSelection.Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent)
                currentSelection.Range.ParagraphFormat.RightIndent = Application.CentimetersToPoints(My.Settings.RightIndent)
            Case Word.WdMeasurementUnits.wdMillimeters
                currentSelection.Range.ParagraphFormat.LeftIndent = Application.MillimetersToPoints(My.Settings.LeftIndent)
                currentSelection.Range.ParagraphFormat.RightIndent = Application.MillimetersToPoints(My.Settings.RightIndent)
            Case Word.WdMeasurementUnits.wdPoints
                currentSelection.Range.ParagraphFormat.LeftIndent = My.Settings.LeftIndent
                currentSelection.Range.ParagraphFormat.RightIndent = My.Settings.RightIndent
            Case Word.WdMeasurementUnits.wdPicas
                currentSelection.Range.ParagraphFormat.LeftIndent = Application.PicasToPoints(My.Settings.LeftIndent)
                currentSelection.Range.ParagraphFormat.RightIndent = Application.PicasToPoints(My.Settings.RightIndent)
        End Select

        'currentSelection.Range.ParagraphFormat.FirstLineIndent = My.Settings.Indent
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "current linespacing = " + My.Settings.Linespacing.ToString)
        Select Case My.Settings.Linespacing
            Case 1.0F
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "current linespacing has been detected as Single")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
            Case 1.5F
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "current linespacing has been detected as one and a half")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
            Case 2.0F
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "current linespacing has been detected as Double")
                currentSelection.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble
        End Select

        Dim myAlignment As ALIGN
        Select Case PARTYPE
            Case PARAGRAPHTYPE.BIBLEVERSION
                myAlignment = My.Settings.BibleVersionAlign
            Case PARAGRAPHTYPE.BOOKCHAPTER
                myAlignment = My.Settings.BookChapterAlign
            Case PARAGRAPHTYPE.VERSES
                myAlignment = My.Settings.ParagraphAlignment
        End Select
        If PARTYPE = PARAGRAPHTYPE.BOOKCHAPTER And My.Settings.BookChapterPosition = POS.BOTTOMINLINE Then
            'Do nothing
        Else
            Select Case myAlignment
                Case ALIGN.LEFT
                    currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                Case ALIGN.CENTER
                    currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Case ALIGN.RIGHT
                    currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                Case ALIGN.JUSTIFY
                    currentSelection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
            End Select
        End If
    End Sub

    Private Sub setTextStyles(ByRef currentSelection As Word.Selection, ByVal PARTYPE As PARAGRAPHTYPE)
        Dim paragraphFont As Word.Font
        Dim foreColor As Word.WdColor
        Dim backColor As Word.WdColor

        Select Case PARTYPE
            Case PARAGRAPHTYPE.BIBLEVERSION
                paragraphFont = wdFontConverter(My.Settings.BibleVersionFont)
                foreColor = CType(ColorTranslator.ToOle(My.Settings.BibleVersionForeColor), Word.WdColor)
                If My.Settings.BibleVersionBackColor = Nothing Then
                    backColor = Word.WdColor.wdColorAutomatic
                Else
                    backColor = CType(ColorTranslator.ToOle(My.Settings.BibleVersionBackColor), Word.WdColor)
                End If
            Case PARAGRAPHTYPE.BOOKCHAPTER
                paragraphFont = wdFontConverter(My.Settings.BookChapterFont)
                foreColor = CType(ColorTranslator.ToOle(My.Settings.BookChapterForeColor), Word.WdColor)
                If My.Settings.BookChapterBackColor = Nothing Then
                    backColor = Word.WdColor.wdColorAutomatic
                Else
                    backColor = CType(ColorTranslator.ToOle(My.Settings.BookChapterBackColor), Word.WdColor)
                End If
            Case PARAGRAPHTYPE.VERSENUMBER
                paragraphFont = wdFontConverter(My.Settings.VerseNumberFont)
                foreColor = CType(ColorTranslator.ToOle(My.Settings.VerseNumberForeColor), Word.WdColor)
                If My.Settings.VerseNumberBackColor = Nothing Then
                    backColor = Word.WdColor.wdColorAutomatic
                Else
                    backColor = CType(ColorTranslator.ToOle(My.Settings.VerseNumberBackColor), Word.WdColor)
                End If
                Select Case My.Settings.VerseNumberVAlign
                    Case VALIGN.SUPERSCRIPT
                        currentSelection.Font.Superscript = True
                    Case VALIGN.SUBSCRIPT
                        currentSelection.Font.Subscript = True
                    Case VALIGN.NORMAL
                        currentSelection.Font.Superscript = False
                        currentSelection.Font.Subscript = False
                End Select
            Case PARAGRAPHTYPE.VERSETEXT
                paragraphFont = wdFontConverter(My.Settings.VerseTextFont)
                foreColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextForeColor), Word.WdColor)
                If My.Settings.VerseTextBackColor = Nothing Then
                    backColor = Word.WdColor.wdColorAutomatic
                Else
                    backColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Word.WdColor)
                End If
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
            Case Else 'this will never be the case, but it's just to say that the variables will certainly be set
                paragraphFont = wdFontConverter(My.Settings.VerseTextFont)
                foreColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextForeColor), Word.WdColor)
                If My.Settings.VerseTextBackColor = Nothing Then
                    backColor = Word.WdColor.wdColorAutomatic
                Else
                    backColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Word.WdColor)
                End If
        End Select
        currentSelection.Font.Name = paragraphFont.Name
        currentSelection.Font.Size = paragraphFont.Size
        currentSelection.Font.Bold = paragraphFont.Bold
        currentSelection.Font.Italic = paragraphFont.Italic
        currentSelection.Font.Underline = paragraphFont.Underline
        currentSelection.Font.Color = foreColor
        currentSelection.Font.Shading.BackgroundPatternColor = backColor
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
