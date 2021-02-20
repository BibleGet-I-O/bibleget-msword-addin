Imports System.ComponentModel
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
    Private pofIndent As Single
    Private poiIndent As Single
    Private leftIndent As Single
    Private rightIndent As Single


    Public Sub New(ByRef worker As BackgroundWorker, ByRef eventArgs As DoWorkEventArgs)
        Me.worker = worker
        e = eventArgs
        DEBUG_MODE = My.Settings.DEBUG_MODE

        Select Case Application.Options.MeasurementUnit
            Case Word.WdMeasurementUnits.wdInches
                leftIndent = Application.InchesToPoints(My.Settings.LeftIndent)
                rightIndent = Application.InchesToPoints(My.Settings.RightIndent)
                pofIndent = Application.InchesToPoints(My.Settings.LeftIndent + 0.2F)
                poiIndent = Application.InchesToPoints(My.Settings.LeftIndent + 0.4F)
            Case Word.WdMeasurementUnits.wdCentimeters
                leftIndent = Application.CentimetersToPoints(My.Settings.LeftIndent)
                rightIndent = Application.CentimetersToPoints(My.Settings.RightIndent)
                pofIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 0.5F)
                poiIndent = Application.CentimetersToPoints(My.Settings.LeftIndent + 1.0F)
            Case Word.WdMeasurementUnits.wdMillimeters
                leftIndent = Application.MillimetersToPoints(My.Settings.LeftIndent)
                rightIndent = Application.MillimetersToPoints(My.Settings.RightIndent)
                pofIndent = Application.MillimetersToPoints(My.Settings.LeftIndent + 5.0F)
                poiIndent = Application.MillimetersToPoints(My.Settings.LeftIndent + 10.0F)
            Case Word.WdMeasurementUnits.wdPoints
                leftIndent = My.Settings.LeftIndent
                rightIndent = My.Settings.RightIndent
                pofIndent = My.Settings.LeftIndent + 14.0F
                poiIndent = My.Settings.LeftIndent + 28.0F
            Case Word.WdMeasurementUnits.wdPicas
                leftIndent = Application.PicasToPoints(My.Settings.LeftIndent)
                rightIndent = Application.PicasToPoints(My.Settings.RightIndent)
                pofIndent = Application.PicasToPoints(My.Settings.LeftIndent + 1.2F)
                poiIndent = Application.PicasToPoints(My.Settings.LeftIndent + 2.4F)
        End Select

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

        Dim prevVersion As String = String.Empty
        Dim newVersion As Boolean
        Dim prevBook As String = String.Empty
        Dim newBook As Boolean
        Dim prevChapter As Integer = -1
        Dim newChapter As Boolean
        Dim prevVerse As String = String.Empty
        Dim newVerse As Boolean

        Dim currentVersion As String
        Dim currentBook As String
        Dim currentBookAbbrev As String
        Dim currentBookUnivIdx As Integer
        Dim currentChapter As Integer
        Dim currentVerse As String
        Dim originalQuery As String

        Dim firstVersion As Boolean = True
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
            currentBook = currentJson.SelectToken("book").Value(Of String)()
            currentBookAbbrev = currentJson.SelectToken("bookabbrev").Value(Of String)()
            currentBookUnivIdx = currentJson.SelectToken("univbooknum").Value(Of Integer)()
            currentChapter = currentJson.SelectToken("chapter").Value(Of Integer)()
            currentVerse = currentJson.SelectToken("verse").Value(Of String)()
            currentVersion = currentJson.SelectToken("version").Value(Of String)()
            originalQuery = currentJson.SelectToken("originalquery").Value(Of String)()

            If Not currentVerse = prevVerse Then
                newVerse = True
                prevVerse = currentVerse
            Else
                newVerse = False
            End If

            If Not currentChapter = prevChapter Then
                newChapter = True
                newVerse = True
                prevChapter = currentChapter
            Else
                newChapter = False
            End If

            If Not currentBook = prevBook Then
                newBook = True
                newChapter = True
                newVerse = True
                prevBook = currentBook
            Else
                newBook = False
            End If

            If Not currentVersion = prevVersion Then
                newVersion = True
                newBook = True
                newChapter = True
                newVerse = True
                prevVersion = currentVersion
            Else
                newVersion = False
            End If

            If newVersion Then
                firstVerse = True
                'firstChapter = True

                Select Case My.Settings.BibleVersionWrap
                    Case WRAP.PARENTHESES
                        currentVersion = "(" & currentVersion & ")"
                    Case WRAP.BRACKETS
                        currentVersion = "[" & currentVersion & "]"
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
                            BibleVersionStack.Add(currentVersion)
                            If BibleVersionStack.Count > 1 Then
                                CreateNewPar(currentSelection)                   'create a new paragraph
                                setParagraphStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                                setTextStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                                TypeText(currentSelection, BibleVersionStack.Item(0))   'insert the Bible version into the paragraph
                                BibleVersionStack.RemoveAt(0)
                            End If
                        Case POS.TOP
                            If firstVersion = False Then
                                CreateNewPar(currentSelection)
                            Else
                                firstVersion = False
                            End If
                            setParagraphStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                            setTextStyles(currentSelection, PARAGRAPHTYPE.BIBLEVERSION)
                            TypeText(currentSelection, currentVersion)
                    End Select
                End If
            End If

            If newBook Or newChapter Then
                '//System.out.println(currentbook+" "+currentchapter);
                Dim bkChStr As String = ""
                Select Case My.Settings.BookChapterFormat
                    Case FORMAT.BIBLELANG
                        bkChStr = currentBook + " " + currentChapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.BIBLELANGABBREV
                        bkChStr = currentBookAbbrev + " " + currentChapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.USERLANG
                        bkChStr = L10NBookNames.GetBookByIndex(currentBookUnivIdx - 1).Fullname + " " + currentChapter.ToString(CultureInfo.InvariantCulture)
                    Case FORMAT.USERLANGABBREV
                        bkChStr = L10NBookNames.GetBookByIndex(currentBookUnivIdx - 1).Abbrev + " " + currentChapter.ToString(CultureInfo.InvariantCulture)
                End Select
                If My.Settings.BookChapterFullReference Then
                    'retrieve the original query from originalquery property in the json response received
                    If Regex.IsMatch(originalQuery, "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}", RegexOptions.Singleline) Then
                        bkChStr &= Regex.Replace(originalQuery, "^[1-3]{0,1}((\p{L}\p{M}*)+)[1-9][0-9]{0,2}", "")
                    Else
                        bkChStr &= Regex.Replace(originalQuery, "^[1-9][0-9]{0,2}", "")
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

            If newVerse Then
                normalText = False
                '//System.out.print("\n"+currentverse);
                If firstVerse Then
                    CreateNewPar(currentSelection)
                    setParagraphStyles(currentSelection, PARAGRAPHTYPE.VERSES)
                    firstVerse = False
                End If
                If My.Settings.VerseNumberVisibility = VISIBILITY.SHOW Then
                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSENUMBER)
                    TypeText(currentSelection, " " + currentVerse)
                Else
                    TypeText(currentSelection, " ")
                End If
            End If
            setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
            Dim currentText As String = currentJson.SelectToken("text").Value(Of String)()
            currentText = currentText.Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
            Dim remainingText As String = currentText
            Dim pattern1 As String = "(.*?)<((speaker|sm|i|pr|po)[f|l|s|i|3]{0,1}[f|l]{0,1})>(.*?)</\2>"

            If Regex.IsMatch(currentText, ".*<[/]{0,1}(?:speaker|sm|i|pr|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>.*", RegexOptions.Singleline) Then '//[/]{0,1}(?:sm|po)[f|l|s|i|3]{0,1}[f|l]{0,1}>
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have detected a text string with special formatting: " + currentText)
                Dim currentSpaceAfter As Single = currentSelection.Range.ParagraphFormat.SpaceAfter

                'HandleFormattingTags(currentText)
                'Matcher matcher1 = pattern1.matcher(currentText);
                'Dim iteration As Integer = 0
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
                        'Insert into the document normal text before the detected special formatting tag
                        TypeText(currentSelection, match.Groups(1).Value)
                        Dim normalTextStr As String = Regex.Escape(match.Groups(1).Value)
                        Dim regx As Regex = New Regex(normalTextStr)
                        remainingText = regx.Replace(remainingText, String.Empty, 1)
                        'remainingText.replaceFirst(match.Groups(1).Value, "")
                    End If

                    If match.Groups(4).Value IsNot Nothing And match.Groups(4).Value IsNot String.Empty Then

                        Dim matchedTag As String = match.Groups(2).Value
                        Dim formattingTagContents As String = match.Groups(4).Value

                        '//check for nested speaker tags!
                        Dim nestedTag As Boolean = False
                        Dim nestedTagObj As NestedTagObj = Nothing
                        'Dim speakerTagBefore As String = ""
                        'Dim speakerTagContents As String = ""
                        'Dim speakerTagAfter As String = ""

                        If Regex.IsMatch(formattingTagContents, ".*<[/]{0,1}(?:speaker|sm|i|pr|po)[f|l|s|i]{0,1}[f|l]{0,1}>.*", RegexOptions.Singleline) Then
                            nestedTag = True
                            'Debug.Print("nestedTag was detected in string {" & formattingTagContents & "}")
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have a nested tag in this special formatting text: " + formattingTagContents)
                            nestedTagObj = New NestedTagObj(formattingTagContents)

                            'Matcher matcher2 = pattern1.matcher(formattingTagContents);
                            'Dim iteration2 As Integer = 0
                            'Dim remainingText2 As String = formattingTagContents
                            'For Each matcher2 As Match In Regex.Matches(formattingTagContents, pattern1)
                            '    If matcher2.Groups(2).Value IsNot Nothing And matcher2.Groups(2).Value IsNot String.Empty And matcher2.Groups(2).Value = "speaker" Then
                            '        If matcher2.Groups(1).Value IsNot Nothing And matcher2.Groups(1).Value IsNot String.Empty Then
                            '            speakerTagBefore = matcher2.Groups(1).Value
                            '            Dim reggaeton As Regex = New Regex(matcher2.Groups(1).Value)
                            '            remainingText2 = reggaeton.Replace(remainingText2, String.Empty, 1)
                            '        End If
                            '        speakerTagContents = matcher2.Groups(4).Value
                            '        Dim reggae As Regex = New Regex("<" + matcher2.Groups(2).Value + ">" + matcher2.Groups(4).Value + "</" + matcher2.Groups(2).Value + ">")
                            '        speakerTagAfter = reggae.Replace(remainingText2, String.Empty, 1)
                            '    End If
                            'Next
                        End If

                        If My.Settings.NOVERSIONFORMATTING Then formattingTagContents = " " + formattingTagContents + " "

                        Select Case matchedTag
                            Case "pof"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
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
                                    currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
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
                                    currentSelection.Range.ParagraphFormat.LeftIndent = poiIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
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
                                    currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
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
                                    currentSelection.Range.ParagraphFormat.LeftIndent = poiIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
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
                                    currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                                End If
                                normalText = False
                            Case "poil"
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    If firstVerse = False And normalText = True Then CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200)+600);
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = poiIndent
                                    currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                                End If
                                If nestedTag Then
                                    insertNestedTag(nestedTagObj, currentSelection)
                                Else
                                    TypeText(currentSelection, formattingTagContents)
                                End If
                                If Not My.Settings.NOVERSIONFORMATTING Then
                                    CreateNewPar(currentSelection)
                                    'xPropertySet.setPropertyValue("ParaLeftMargin", (paragraphLeftIndent*200));
                                    setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                                    currentSelection.Range.ParagraphFormat.LeftIndent = leftIndent
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
                                TypeText(currentSelection, " " & match.Groups(4).Value & " ")
                                If Not My.Settings.VerseTextFont.Bold Then
                                    currentSelection.Font.Bold = False
                                End If
                                'xPropertySet.setPropertyValue("CharBackColor", bgColorVerseText.getRGB() & ~0xFF000000);
                                currentSelection.Font.Shading.BackgroundPatternColor = CType(ColorTranslator.ToOle(My.Settings.VerseTextBackColor), Word.WdColor)
                        End Select

                        'Debug.Print("currentVerse: " & currentverse & " :: remainingText before regex.replace = {" & remainingText & "}")
                        Dim remainingPattern As String = Regex.Escape("<" + match.Groups(2).Value + ">" + match.Groups(4).Value + "</" + match.Groups(2).Value + ">")
                        Dim nonmereggaepiu As Regex = New Regex(remainingPattern)
                        remainingText = nonmereggaepiu.Replace(remainingText, String.Empty, 1)
                        'Debug.Print("currentVerse: " & currentverse & " :: remainingText after regex.replace = {" & remainingText & "}")

                    End If
                Next
                currentSelection.Range.ParagraphFormat.SpaceAfter = currentSpaceAfter
                '//                System.out.println("We have a match for special formatting: "+currentText);
                '//                System.out.println("And after elaborating our matches, this is what we have left: "+remainingText);
                '//                System.out.println();
                If remainingText IsNot String.Empty Then
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "We have a fragment of text left over after all this: " + remainingText)
                    TypeText(currentSelection, remainingText)
                    'Debug.Print("remainingText = {" & remainingText & "}")
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

        If My.Settings.BibleVersionVisibility = VISIBILITY.SHOW And BibleVersionStack.Count > 0 Then
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
        currentSelection.Range.ParagraphFormat.LeftIndent = leftIndent
        currentSelection.Range.ParagraphFormat.RightIndent = rightIndent

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

    Private Sub insertNestedTag(ByVal nestedTagObj As NestedTagObj, ByVal currentSelection As Word.Selection)
        If nestedTagObj.Before IsNot String.Empty Then
            setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
            TypeText(currentSelection, nestedTagObj.Before)
        End If

        Select Case nestedTagObj.Tag
            Case "speaker"
                currentSelection.Font.Bold = True
                currentSelection.Font.Italic = False
                currentSelection.Font.Underline = False
                currentSelection.Font.Size = My.Settings.VerseTextFont.Size
                currentSelection.Font.Color = Word.WdColor.wdColorBlack
                currentSelection.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30
                currentSelection.Font.Superscript = False
                currentSelection.Font.Subscript = False
                TypeText(currentSelection, nestedTagObj.Contents)
            Case "sm"
                currentSelection.Font.SmallCaps = True
                TypeText(currentSelection, nestedTagObj.Contents)
                currentSelection.Font.SmallCaps = False
            Case "poi"
                'If firstVerse = False And normalText = True Then CreateNewPar(currentSelection) ???
                CreateNewPar(currentSelection)
                currentSelection.Range.ParagraphFormat.LeftIndent = poiIndent
                currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                TypeText(currentSelection, nestedTagObj.Contents)
                CreateNewPar(currentSelection)
                setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
            Case "po"
                'If firstVerse = False And normalText = True Then CreateNewPar(currentSelection) ???
                CreateNewPar(currentSelection)
                currentSelection.Range.ParagraphFormat.LeftIndent = pofIndent
                currentSelection.Range.ParagraphFormat.SpaceAfter = 0F
                setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
                TypeText(currentSelection, nestedTagObj.Contents)
                CreateNewPar(currentSelection)
                setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
        End Select

        If nestedTagObj.After IsNot String.Empty Then
            setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)
            TypeText(currentSelection, nestedTagObj.After)
        End If

        setTextStyles(currentSelection, PARAGRAPHTYPE.VERSETEXT)

    End Sub

    'Private Sub HandleFormattingTags(ByVal currentText As String)
    '    Dim pattern1 As String = "(.*?)<((speaker|sm|i|pr|po)[f|l|s|i|3]{0,1}[f|l]{0,1})>(.*?)</\2>"

    'End Sub

End Class
