Imports System.Drawing
Imports System.Globalization

Public Class Preferences
    Private initializing As Boolean = True

    'turn off the annoying clicking sound when the preview window refreshes (WebBrowser control)
    Const DS As Integer = 21
    Const SP As Integer = &H2
    Private DEBUG_MODE As Boolean
    Private Application As Word.Application = Globals.BibleGetAddIn.Application
    Private InterfaceInCM As Boolean

    Private Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = BibleGetAddIn.RM.GetString(myStr, BibleGetAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub setStyleLable(ByVal myFont As Drawing.Font, ByVal myCase As String)
        Dim fntSize As Integer
        Dim styleStr As String
        fntSize = Math.Round(myFont.SizeInPoints)
        styleStr = fntSize.ToString(CultureInfo.InvariantCulture) & "pt   "
        styleStr &= If(myFont.Style And Drawing.FontStyle.Bold, " Bold", "")
        styleStr &= If(myFont.Style And Drawing.FontStyle.Italic, " Italic", "")
        styleStr &= If(myFont.Style And Drawing.FontStyle.Underline, " Underscore", "")
        styleStr &= If(((myFont.Style And Drawing.FontStyle.Bold) = False) And ((myFont.Style And Drawing.FontStyle.Italic) = False) And ((myFont.Style And Drawing.FontStyle.Underline) = False), " Normal", "")
        'styleStr &= If(My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Regular, " Normal", "")

        Select Case myCase
            Case "BibleVersion"
                BibleVersionStyeLbl.Text = styleStr
            Case "BookChapter"
                BookChapterStyleLbl.Text = styleStr
            Case "VerseNumber"
                styleStr &= If(VerseNumberSuperscriptBtn.Checked, " Superscript", "")
                styleStr &= If(VerseNumberSubscriptBtn.Checked, " Subscript", "")
                VerseNumberStyleLbl.Text = styleStr
            Case "VerseText"
                VerseTextStyleLbl.Text = styleStr
        End Select
    End Sub

    Private Sub setFontBtn(ByVal myCase As String)
        Select Case myCase
            Case "BibleVersion"
                BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                BibleVersionFontBtnn.Text = My.Settings.BibleVersionFont.Name
                BibleVersionFontBtnn.ForeColor = My.Settings.BibleVersionForeColor
                BibleVersionFontBtnn.BackColor = My.Settings.BibleVersionBackColor
            Case "BookChapter"
                BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                BookChapterFontBtnn.Text = My.Settings.BookChapterFont.Name
                BookChapterFontBtnn.ForeColor = My.Settings.BookChapterForeColor
                BookChapterFontBtnn.BackColor = My.Settings.BookChapterBackColor
            Case "VerseNumber"
                VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                VerseNumberFontBtnn.Text = My.Settings.VerseNumberFont.Name
                VerseNumberFontBtnn.ForeColor = My.Settings.VerseNumberForeColor
                VerseNumberFontBtnn.BackColor = My.Settings.VerseNumberBackColor
            Case "VerseText"
                VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                VerseTextFontBtnn.Text = My.Settings.VerseTextFont.Name
                VerseTextFontBtnn.ForeColor = My.Settings.VerseTextForeColor
                VerseTextFontBtnn.BackColor = My.Settings.VerseTextBackColor
        End Select

    End Sub

    Private Sub setCheckBtns(ByVal myCase As String)
        Select Case myCase
            Case "BibleVersion"
                BibleVersionBoldBtn.Checked = (My.Settings.BibleVersionFont.Style And Drawing.FontStyle.Bold)
                BibleVersionItalicBtn.Checked = (My.Settings.BibleVersionFont.Style And Drawing.FontStyle.Italic)
                BibleVersionUnderlineBtn.Checked = (My.Settings.BibleVersionFont.Style And Drawing.FontStyle.Underline)
            Case "BookChapter"
                BookChapterBoldBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Bold)
                BookChapterItalicBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Italic)
                BookChapterUnderlineBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Underline)
            Case "VerseNumber"
                VerseNumberBoldBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Bold)
                VerseNumberItalicBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Italic)
                VerseNumberUnderlineBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Underline)
                VerseNumberSuperscriptBtn.Checked = (My.Settings.VerseNumberVAlign = VALIGN.SUPERSCRIPT)
                VerseNumberSubscriptBtn.Checked = (My.Settings.VerseNumberVAlign = VALIGN.SUBSCRIPT)
            Case "VerseText"
                VerseTextBoldBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Bold)
                VerseTextItalicBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Italic)
                VerseTextUnderlineBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Underline)
        End Select
    End Sub

    Private Sub checkBoxChanged(ByVal myCase As String, ByVal myBtn As String)
        Select Case myBtn
            Case "Bold"
                Select Case myCase
                    Case "BibleVersion"
                        If BibleVersionBoldBtn.Checked Then
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style Or Drawing.FontStyle.Bold)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style And Not Drawing.FontStyle.Bold)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "BookChapter"
                        If BookChapterBoldBtn.Checked Then
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style Or Drawing.FontStyle.Bold)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style And Not Drawing.FontStyle.Bold)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberBoldBtn.Checked Then
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style Or Drawing.FontStyle.Bold)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Bold)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextBoldBtn.Checked Then
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style Or Drawing.FontStyle.Bold)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style And Not Drawing.FontStyle.Bold)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
            Case "Italic"
                Select Case myCase
                    Case "BibleVersion"
                        If BibleVersionItalicBtn.Checked Then
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style Or Drawing.FontStyle.Italic)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style And Not Drawing.FontStyle.Italic)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "BookChapter"
                        If BookChapterItalicBtn.Checked Then
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style Or Drawing.FontStyle.Italic)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style And Not Drawing.FontStyle.Italic)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberItalicBtn.Checked Then
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style Or Drawing.FontStyle.Italic)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Italic)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextItalicBtn.Checked Then
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style Or Drawing.FontStyle.Italic)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style And Not Drawing.FontStyle.Italic)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
            Case "Underline"
                Select Case myCase
                    Case "BibleVersion"
                        If BibleVersionUnderlineBtn.Checked Then
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style Or Drawing.FontStyle.Underline)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BibleVersionFont = New Drawing.Font(My.Settings.BibleVersionFont.Name, My.Settings.BibleVersionFont.Size, My.Settings.BibleVersionFont.Style And Not Drawing.FontStyle.Underline)
                            BibleVersionFontBtnn.Font = New Drawing.Font(My.Settings.BibleVersionFont.Name, 12, My.Settings.BibleVersionFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "BookChapter"
                        If BookChapterUnderlineBtn.Checked Then
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style Or Drawing.FontStyle.Underline)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.BookChapterFont = New Drawing.Font(My.Settings.BookChapterFont.Name, My.Settings.BookChapterFont.Size, My.Settings.BookChapterFont.Style And Not Drawing.FontStyle.Underline)
                            BookChapterFontBtnn.Font = New Drawing.Font(My.Settings.BookChapterFont.Name, 12, My.Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberUnderlineBtn.Checked Then
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style Or Drawing.FontStyle.Underline)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseNumberFont = New Drawing.Font(My.Settings.VerseNumberFont.Name, My.Settings.VerseNumberFont.Size, My.Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Underline)
                            VerseNumberFontBtnn.Font = New Drawing.Font(My.Settings.VerseNumberFont.Name, 12, My.Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextUnderlineBtn.Checked Then
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style Or Drawing.FontStyle.Underline)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            My.Settings.VerseTextFont = New Drawing.Font(My.Settings.VerseTextFont.Name, My.Settings.VerseTextFont.Size, My.Settings.VerseTextFont.Style And Not Drawing.FontStyle.Underline)
                            VerseTextFontBtnn.Font = New Drawing.Font(My.Settings.VerseTextFont.Name, 12, My.Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
        End Select
        My.Settings.Save()
    End Sub

    Private Sub setPreviewDocument()
        Dim previewDocument As String
        Dim stylesheet As String
        Dim script As String

        Dim textColorBibleVersion As String = If(Not My.Settings.BibleVersionForeColor.IsEmpty, "#" & My.Settings.BibleVersionForeColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim bgColorBibleVersion As String = If(Not My.Settings.BibleVersionBackColor.IsEmpty, "#" & My.Settings.BibleVersionBackColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim textColorBookChapter As String = If(Not My.Settings.BookChapterForeColor.IsEmpty, "#" & My.Settings.BookChapterForeColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim bgColorBookChapter As String = If(Not My.Settings.BookChapterBackColor.IsEmpty, "#" & My.Settings.BookChapterBackColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim textColorVerseNumber As String = If(Not My.Settings.VerseNumberForeColor.IsEmpty, "#" & My.Settings.VerseNumberForeColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim bgColorVerseNumber As String = If(Not My.Settings.VerseNumberBackColor.IsEmpty, "#" & My.Settings.VerseNumberBackColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim textColorVerseText As String = If(Not My.Settings.VerseTextForeColor.IsEmpty, "#" & My.Settings.VerseTextForeColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        Dim bgColorVerseText As String = If(Not My.Settings.VerseTextBackColor.IsEmpty, "#" & My.Settings.VerseTextBackColor.ToArgb().ToString("X", CultureInfo.InvariantCulture).Substring(2), "transparent")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & bgColorVerseNumber)

        Dim vnPosition As String = If(My.Settings.VerseNumberVAlign = VALIGN.NORMAL, "position: static;", "position: relative;")
        Dim vnTop As String = ""
        Select Case My.Settings.VerseNumberVAlign
            Case VALIGN.SUPERSCRIPT
                vnTop = " top: -0.6em;"
            Case VALIGN.SUBSCRIPT
                vnTop = " top: 0.6em;"
            Case VALIGN.NORMAL
                vnTop = ""
        End Select

        Dim bibleVersionWrapBefore As String = ""
        Dim bibleVersionWrapAfter As String = ""
        If My.Settings.BibleVersionWrap = WRAP.PARENTHESES Then
            bibleVersionWrapBefore = "("
            bibleVersionWrapAfter = ")"
        ElseIf My.Settings.BibleVersionWrap = WRAP.BRACKETS Then
            bibleVersionWrapBefore = "["
            bibleVersionWrapAfter = "]"
        End If

        Dim bookChapterWrapBefore As String = ""
        Dim bookChapterWrapAfter As String = ""
        If My.Settings.BookChapterWrap = WRAP.PARENTHESES Then
            bookChapterWrapBefore = "("
            bookChapterWrapAfter = ")"
        ElseIf My.Settings.BookChapterWrap = WRAP.BRACKETS Then
            bookChapterWrapBefore = "["
            bookChapterWrapAfter = "]"
        End If



        previewDocument = "<!DOCTYPE html>"
        previewDocument &= "<head>"
        previewDocument &= "<meta http-equiv=""X-UA-Compatible"" content=""IE=Edge"" >"
        previewDocument &= "<meta charset=""UTF-8"">"

        stylesheet = "<style type=""text/css"">"
        stylesheet &= "html,body { padding: 0px; margin: 0px; background-color: #FFFFFF; }"
        stylesheet &= "p { padding: 0px; margin: 0px; }"
        stylesheet &= ".previewRuler { margin: 0px auto; }"
        stylesheet &= "div.results {  
  box-sizing: border-box;
  margin: 0px auto;
  padding-left: 35px;
  padding-right:35px;
}"
        stylesheet &= "div.results .bibleVersion { font-family: " & My.Settings.BookChapterFont.Name & "; }"
        stylesheet &= "div.results .bibleVersion { font-size: " & Math.Round(My.Settings.BookChapterFont.SizeInPoints).ToString(CultureInfo.InvariantCulture) & "pt; }"
        stylesheet &= "div.results .bibleVersion { font-weight: " & If(BibleVersionBoldBtn.Checked, "bold", "normal") & "; }"
        stylesheet &= "div.results .bibleVersion { font-style: " & If(BibleVersionItalicBtn.Checked, "italic", "normal") & "; }"
        If BibleVersionUnderlineBtn.Checked = True Then
            stylesheet &= "div.results .bibleVersion { text-decoration: underline; }"
        End If
        stylesheet &= "div.results .bibleVersion { color: " & textColorBibleVersion & "; }"
        stylesheet &= "div.results .bibleVersion { background-color: " & bgColorBibleVersion & "; }"
        stylesheet &= "div.results .bibleVersion { text-align: " & CSSRULE.ALIGN(My.Settings.BibleVersionAlign) & "; }"
        stylesheet &= "div.results .bookChapter { font-family: " & My.Settings.BookChapterFont.Name & "; }"
        stylesheet &= "div.results .bookChapter { font-size: " & Math.Round(My.Settings.BookChapterFont.SizeInPoints).ToString(CultureInfo.InvariantCulture) & "pt; }"
        stylesheet &= "div.results .bookChapter { font-weight: " & If(BookChapterBoldBtn.Checked, "bold", "normal") & "; }"
        stylesheet &= "div.results .bookChapter { font-style: " & If(BookChapterItalicBtn.Checked, "italic", "normal") & "; }"
        If BookChapterUnderlineBtn.Checked = True Then
            stylesheet &= "div.results .bookChapter { text-decoration: underline; }"
        End If
        stylesheet &= "div.results .bookChapter { color: " & textColorBookChapter & "; }"
        stylesheet &= "div.results .bookChapter { background-color: " & bgColorBookChapter & "; }"
        stylesheet &= "div.results .bookChapter { text-align: " & CSSRULE.ALIGN(My.Settings.BookChapterAlign) & "; }"
        stylesheet &= "div.results span.bookChapter { display: inline-block; margin-left: 6px; }"
        stylesheet &= "div.results .versesParagraph { text-align: " & CSSRULE.ALIGN(My.Settings.ParagraphAlignment) & "; }"
        stylesheet &= "div.results .versesParagraph { line-height: " & My.Settings.Linespacing.ToString("F1", CultureInfo.InvariantCulture) & "em; }"
        stylesheet &= "div.results .versesParagraph .verseNum { font-family: " & My.Settings.VerseNumberFont.Name & "; }"
        stylesheet &= "div.results .versesParagraph .verseNum { font-size:" & Math.Round(My.Settings.VerseNumberFont.SizeInPoints).ToString(CultureInfo.InvariantCulture) & "pt; }"
        stylesheet &= "div.results .versesParagraph .verseNum { font-weight: " & If(VerseNumberBoldBtn.Checked, "bold", "normal") & "; }"
        stylesheet &= "div.results .versesParagraph .verseNum { font-style: " & If(VerseNumberItalicBtn.Checked, "italic", "normal") & "; }"
        If VerseNumberUnderlineBtn.Checked = True Then
            stylesheet &= "div.results .versesParagraph .verseNum { text-decoration: underline; }"
        End If
        stylesheet &= "div.results .versesParagraph .verseNum { color: " & textColorVerseNumber & "; }"
        stylesheet &= "div.results .versesParagraph .verseNum { background-color: " & bgColorVerseNumber & "; }"
        stylesheet &= "div.results .versesParagraph .verseNum { vertical-align: baseline; " & vnPosition & vnTop & " }"
        stylesheet &= "div.results .versesParagraph .verseNum { padding-left: 3px; }"
        stylesheet &= "div.results .versesParagraph .verseText { font-family: " & My.Settings.VerseTextFont.Name & "; }"
        stylesheet &= "div.results .versesParagraph .verseText { font-size:" & Math.Round(My.Settings.VerseTextFont.SizeInPoints).ToString(CultureInfo.InvariantCulture) & "pt; }"
        stylesheet &= "div.results .versesParagraph .verseText { font-weight: " & If(VerseTextBoldBtn.Checked, "bold", "normal") & "; }"
        stylesheet &= "div.results .versesParagraph .verseText { font-style: " & If(VerseTextItalicBtn.Checked, "italic", "normal") & "; }"
        If VerseTextUnderlineBtn.Checked = True Then
            stylesheet &= "div.results .versesParagraph .verseText { text-decoration: underline; }"
        End If
        stylesheet &= "div.results .versesParagraph .verseText { color: " & textColorVerseText & "; }"
        stylesheet &= "div.results .versesParagraph .verseText { background-color: " & bgColorVerseText & "; }"
        stylesheet &= "</style>"

        script = "<script src=""https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js""></script>"
        script &= "<script type=""text/javascript"">"
        script &= "var getPixelRatioVals = function(rulerLength,convertToCM){
  let inchesToCM = 2.54,
    dpr = window.devicePixelRatio,
    //ppi = ((96 * dpr) / 100),
    //dpi = (96 * ppi),
    dpi = 96 * dpr,
    drawInterval = 0.125;
  if(convertToCM){
    //ppi /= inchesToCM;
    dpi /= inchesToCM;
    rulerLength *= inchesToCM;
    drawInterval = 0.25;
  }
  return {
    inchesToCM: inchesToCM,
    dpr: dpr,      
    //ppi: ppi,
    dpi: dpi,
    rulerLength: rulerLength,
    drawInterval: drawInterval
  };
},
triangleAt = function (x,context,pixelRatioVals,initialPadding,canvasWidth){
  let xPos = x*pixelRatioVals.dpi;//x.map(0,pixelRatioVals.rulerLength,0,(canvasWidth-(initialPadding*2)));
  window.xVal = x;
  window.xPos = xPos;
  context.lineWidth = 0.5;
  context.fillStyle = ""#4285F4"";
  context.beginPath();
  context.moveTo(initialPadding+xPos-6,11);
  context.lineTo(initialPadding+xPos+6,11);
  context.lineTo(initialPadding+xPos,18);
  context.closePath();
  context.stroke();
  context.fill();
},
/**
 * FUNCTION drawRuler
 * @ rulerLen (float) in Inches (e.g. 7)
 * @ cvtToCM (boolean) whether to convert values from inches to centimeters
 * @ lftindnt (float) in Inches, to set xPos of triangle for left indent
 * @ rgtindnt (float) in Inches, to set xPos of triangle for right indent
*/
drawRuler = function(rulerLen, cvtToCM, lftindnt, rgtindnt){
  var pixelRatioVals = getPixelRatioVals(rulerLen,cvtToCM),
      initialPadding = 35,
      $canvas = jQuery('.previewRuler');
  $canvas.each(function(){
    let canvas = this,
        context = canvas.getContext('2d'),
        canvasWidth = (rulerLen * 96 * pixelRatioVals.dpr) + (initialPadding*2);
    canvas.style.width = canvasWidth + 'px';
    canvas.style.height = '20px';
    canvas.width = Math.round(canvasWidth * pixelRatioVals.dpr);
    canvas.height = Math.round(20 * pixelRatioVals.dpr);
    canvas.style.width = Math.round(canvas.width / pixelRatioVals.dpr) + 'px';
    canvas.style.height = Math.round(canvas.height / pixelRatioVals.dpr) + 'px';

    context.scale(pixelRatioVals.dpr,pixelRatioVals.dpr);
    context.translate(pixelRatioVals.dpr, 0);
      
    context.lineWidth = 0.5;
    context.strokeStyle = '#000';
    context.font = 'bold 10px Arial';

    context.beginPath();
    context.moveTo(initialPadding, 1);
    context.lineTo(initialPadding + (pixelRatioVals.rulerLength * pixelRatioVals.dpi),1);
    context.stroke();

    let currentWholeNumber = 0;
    let offset = 2;

    for(let interval = 0; interval <= pixelRatioVals.rulerLength; interval += pixelRatioVals.drawInterval){
      let xPosA = Math.round(interval*pixelRatioVals.dpi)+0.5;
    
      if(interval == Math.floor(interval) && interval > 0){
        if(currentWholeNumber+1 == 10){ offset+=4; }
        context.fillText(++currentWholeNumber,initialPadding+xPosA-offset,14);
      }
      else if(interval == Math.floor(interval)+0.5){
        context.beginPath();
        context.moveTo(initialPadding+xPosA,15);
        context.lineTo(initialPadding+xPosA,5); 
        context.closePath();
        context.stroke();   
      }
      else{
        context.beginPath();
        context.moveTo(initialPadding+xPosA,10);
        context.lineTo(initialPadding+xPosA,5);
        context.closePath();
        context.stroke();
      }
    }
  
    triangleAt(lftindnt,context,pixelRatioVals,initialPadding,canvasWidth);
    triangleAt(pixelRatioVals.rulerLength-rgtindnt,context,pixelRatioVals,initialPadding,canvasWidth);
    context.translate(-pixelRatioVals.dpr, -0);

  });
  
};
jQuery(document).ready(function(){
    let pixelRatioVals = getPixelRatioVals(7," & InterfaceInCM.ToString.ToLower & ");
    let leftindent = " & My.Settings.LeftIndent.ToString("F1", CultureInfo.InvariantCulture) & " * pixelRatioVals.dpi + 35;
    let rightindent = " & My.Settings.RightIndent.ToString("F1", CultureInfo.InvariantCulture) & " * pixelRatioVals.dpi + 35;
    let bestWidth = 7 * 96 * window.devicePixelRatio + (35*2);
    $('.bibleQuote').css({""width"":bestWidth+""px"",""padding-left"":leftindent+""px"",""padding-right"":rightindent+""px""}); 
    drawRuler(7," & InterfaceInCM.ToString.ToLower & "," & My.Settings.LeftIndent.ToString("F1", CultureInfo.InvariantCulture) & "," & My.Settings.RightIndent.ToString("F1", CultureInfo.InvariantCulture) & ");
});
"

        script &= "</script>"

        previewDocument &= stylesheet
        previewDocument &= script
        previewDocument &= "</head>"
        previewDocument &= "<body>"
        previewDocument &= "<div style=""text-align: center;""><canvas class=""previewRuler""></canvas></div>"
        previewDocument &= "<div class=""results bibleQuote"">"
        If My.Settings.BibleVersionPosition = POS.TOP And My.Settings.BibleVersionVisibility = VISIBILITY.SHOW Then
            previewDocument &= "<p class=""bibleVersion"">" & bibleVersionWrapBefore & "NVBSE" & bibleVersionWrapAfter & "</p>"
        End If
        If My.Settings.BookChapterPosition = POS.TOP Then
            previewDocument &= "<p class=""bookChapter"">" & bookChapterWrapBefore & "Genesis&nbsp;1" & bookChapterWrapAfter & "</p>"
        End If
        previewDocument &= "<p class=""versesParagraph"" style=""margin-top:0px;"">"
        If My.Settings.VerseNumberVisibility = VISIBILITY.SHOW Then
            previewDocument &= "<span class=""verseNum"">1</span>"
        End If
        previewDocument &= "<span class=""verseText"">In principio creavit Deus caelum et terram.</span>"
        If My.Settings.VerseNumberVisibility = VISIBILITY.SHOW Then
            previewDocument &= "<span class=""verseNum"">2</span>"
        End If
        previewDocument &= "<span class=""verseText"">Terra autem erat inanis et vacua, et tenebrae super faciem abyssi, et spiritus Dei ferebatur super aquas.</span>"
        If My.Settings.VerseNumberVisibility = VISIBILITY.SHOW Then
            previewDocument &= "<span class=""verseNum"">3</span>"
        End If
        previewDocument &= "<span class=""verseText"">Dixitque Deus: ""Fiat lux"". Et facta est lux.</span>"
        If My.Settings.BookChapterPosition = POS.BOTTOMINLINE Then
            previewDocument &= "<span class=""bookChapter"">" & bookChapterWrapBefore & "Genesis&nbsp;1" & bookChapterWrapAfter & "</span>"
        End If
        previewDocument &= "</p>"
        If My.Settings.BookChapterPosition = POS.BOTTOM Then
            previewDocument &= "<p class=""bookChapter"">" & bookChapterWrapBefore & "Genesis&nbsp;1" & bookChapterWrapAfter & "</p>"
        End If
        If My.Settings.BibleVersionPosition = POS.BOTTOM And My.Settings.BibleVersionVisibility = VISIBILITY.SHOW Then
            previewDocument &= "<p class=""bibleVersion"">" & bibleVersionWrapBefore & "NVBSE" & bibleVersionWrapAfter & "</p>"
        End If
        previewDocument &= "</div>"
        previewDocument &= "</body>"


        If WebBrowser1.Document Is Nothing Then
            WebBrowser1.DocumentText = previewDocument
        Else
            WebBrowser1.Document.Write(String.Empty)
            WebBrowser1.Document.Write(previewDocument)
            WebBrowser1.Refresh()
        End If
    End Sub

    Private Sub Preferences_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DEBUG_MODE = My.Settings.DEBUG_MODE
        Text = __("User Preferences")

        setFontBtn("BibleVersion")
        setStyleLable(My.Settings.BookChapterFont, "BibleVersion")
        setCheckBtns("BibleVersion")

        setFontBtn("BookChapter")
        setStyleLable(My.Settings.BookChapterFont, "BookChapter")
        setCheckBtns("BookChapter")

        setFontBtn("VerseNumber")
        setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
        setCheckBtns("VerseNumber")

        setFontBtn("VerseText")
        setStyleLable(My.Settings.VerseTextFont, "VerseText")
        setCheckBtns("VerseText")

        Select Case My.Settings.Linespacing
            Case 1.0F
                ComboBox1.SelectedIndex = 0
            Case 1.5F
                ComboBox1.SelectedIndex = 1
            Case 2.0F
                ComboBox1.SelectedIndex = 2
        End Select
        GroupBox5.Text = __("Paragraph")
        GroupBox11.Text = __("Bible version")
        GroupBox1.Text = __("Book / Chapter")
        GroupBox2.Text = __("Verse Number")
        GroupBox13.Text = __("Layout preferences for Bible version")
        GroupBox16.Text = __("Layout preferences for Book / Chapter")
        GroupBox14.Text = __("Layout preferences for Verse number")
        GroupBox6.Text = __("Alignment")
        GroupBox7.Text = __("Left Indent")
        GroupBox10.Text = __("Right Indent")
        GroupBox8.Text = __("Line-spacing")
        GroupBox9.Text = __("Override Bible Version Formatting")
        GroupBox3.Text = __("Verse Text")
        GroupBox4.Text = __("Preview")
        ToolTip1.SetToolTip(Label1, __("Some Bible versions have their own formatting. This is left by default to keep the text as close as possible to the original.<br> If however you need to have consistent formatting in your document, you may override the Bible version's own formatting."))
        GroupBox12.Text = __("Position")
        GroupBox15.Text = __("Position")
        GroupBox19.Text = __("Wrap")
        GroupBox20.Text = __("Wrap")
        GroupBox21.Text = __("Display options")
        Label4.Text = __("Hidden")
        Label5.Text = __("Hidden")
        Label3.Text = __("Visible")
        Label6.Text = __("Visible")
        RadioButton11.Text = __("NONE")
        RadioButton21.Text = __("NONE")
        RadioButton22.Text = __("Bible Lang Abbrev")
        RadioButton23.Text = __("Bible Lang Full")
        RadioButton24.Text = __("System Lang Abbrev")
        RadioButton25.Text = __("System Lang Full")
        GroupBox22.Text = __("Current displayed units of Measurement in the Microsoft Word interface:")

        Select Case My.Settings.ParagraphAlignment
            Case BibleGetIO.ALIGN.LEFT
                RadioButton1.Checked = True
            Case BibleGetIO.ALIGN.CENTER
                RadioButton2.Checked = True
            Case BibleGetIO.ALIGN.RIGHT
                RadioButton3.Checked = True
            Case BibleGetIO.ALIGN.JUSTIFY
                RadioButton4.Checked = True
        End Select

        Dim rgtIndIncBitM As Bitmap = My.Resources.increase_indent
        Dim rgtIndDecBitM As Bitmap = My.Resources.decrease_indent
        rgtIndIncBitM.RotateFlip(RotateFlipType.RotateNoneFlipX)
        rgtIndDecBitM.RotateFlip(RotateFlipType.RotateNoneFlipX)
        RightIndentBtn.Image = rgtIndIncBitM
        RightIndentBtn2.Image = rgtIndDecBitM

        CheckBox1.Checked = My.Settings.NOVERSIONFORMATTING
        CheckBox2.Checked = (My.Settings.BibleVersionVisibility = VISIBILITY.SHOW)
        Label3.Visible = (My.Settings.BibleVersionVisibility = VISIBILITY.SHOW)
        Label4.Visible = (My.Settings.BibleVersionVisibility = VISIBILITY.HIDE)
        CheckBox3.Checked = (My.Settings.VerseNumberVisibility = VISIBILITY.SHOW)
        Label6.Visible = (My.Settings.VerseNumberVisibility = VISIBILITY.SHOW)
        Label5.Visible = (My.Settings.VerseNumberVisibility = VISIBILITY.HIDE)

        Select Case My.Settings.BibleVersionAlign
            Case ALIGN.LEFT
                RadioButton7.Checked = True
            Case ALIGN.CENTER
                RadioButton6.Checked = True
            Case ALIGN.RIGHT
                RadioButton5.Checked = True
        End Select

        Select Case My.Settings.BibleVersionPosition
            Case POS.TOP
                RadioButton13.Checked = True
            Case POS.BOTTOM
                RadioButton12.Checked = True
        End Select

        Select Case My.Settings.BibleVersionWrap
            Case WRAP.NONE
                RadioButton11.Checked = True
            Case WRAP.PARENTHESES
                RadioButton17.Checked = True
            Case WRAP.BRACKETS
                RadioButton18.Checked = True
        End Select

        Select Case My.Settings.BookChapterAlign
            Case ALIGN.LEFT
                RadioButton10.Checked = True
            Case ALIGN.CENTER
                RadioButton9.Checked = True
            Case ALIGN.RIGHT
                RadioButton8.Checked = True
        End Select

        Select Case My.Settings.BookChapterPosition
            Case POS.TOP
                RadioButton16.Checked = True
            Case POS.BOTTOM
                RadioButton15.Checked = True
            Case POS.BOTTOMINLINE
                RadioButton14.Checked = True
        End Select

        Select Case My.Settings.BookChapterWrap
            Case WRAP.NONE
                RadioButton21.Checked = True
            Case WRAP.PARENTHESES
                RadioButton20.Checked = True
            Case WRAP.BRACKETS
                RadioButton19.Checked = True
        End Select

        Select Case My.Settings.BookChapterFormat
            Case FORMAT.BIBLELANGABBREV
                RadioButton22.Checked = True
            Case FORMAT.BIBLELANG
                RadioButton23.Checked = True
            Case FORMAT.USERLANGABBREV
                RadioButton24.Checked = True
            Case FORMAT.USERLANG
                RadioButton25.Checked = True
        End Select

        'we need to detect if display measurement units option has changed
        'if it has changed we need to adapt our left and right indents accordingly
        If Application.Options.MeasurementUnit <> My.Settings.CurrentDisplayUnit Then
            Select Case Application.Options.MeasurementUnit
                Case Word.WdMeasurementUnits.wdCentimeters
                    If My.Settings.CurrentDisplayUnit = Word.WdMeasurementUnits.wdInches Then
                        My.Settings.LeftIndent = My.Settings.LeftIndent * 2.0F
                        My.Settings.RightIndent = My.Settings.RightIndent * 2.0F
                    End If
                Case Word.WdMeasurementUnits.wdInches
                    If My.Settings.CurrentDisplayUnit = Word.WdMeasurementUnits.wdCentimeters Then
                        My.Settings.LeftIndent = My.Settings.LeftIndent / 2.0F
                        My.Settings.RightIndent = My.Settings.RightIndent / 2.0F
                    End If
            End Select
            My.Settings.CurrentDisplayUnit = Application.Options.MeasurementUnit
            My.Settings.Save()
        End If


        Select Case Application.Options.MeasurementUnit
            Case Word.WdMeasurementUnits.wdCentimeters
                InterfaceInCM = True
                Label2.Text = My.Settings.LeftIndent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
                Label7.Text = My.Settings.RightIndent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
            Case Else
                InterfaceInCM = False
                Label2.Text = My.Settings.LeftIndent.ToString("F1", CultureInfo.InvariantCulture) & "in"
                Label7.Text = My.Settings.RightIndent.ToString("F1", CultureInfo.InvariantCulture) & "in"
        End Select

        Select Case InterfaceInCM
            Case True
                RadioButton27.Checked = True
            Case False
                RadioButton26.Checked = True
        End Select

        NativeMethods.CoInternetSetFeatureEnabled(DS, SP, True)

        setPreviewDocument()

        initializing = False
    End Sub

    Private Sub BookChapterFontBtn_Click(sender As Object, e As EventArgs) Handles BookChapterFontBtnn.Click
        FontDlg.Font = My.Settings.BookChapterFont
        FontDlg.ShowDialog()
        My.Settings.BookChapterFont = FontDlg.Font
        My.Settings.Save()

        setFontBtn("BookChapter")

        setStyleLable(My.Settings.BookChapterFont, "BookChapter")

        setCheckBtns("BookChapter")
        setPreviewDocument()
    End Sub

    Private Sub CheckBoxBold_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterBoldBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Bold")
        setStyleLable(My.Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub CheckBoxItalic_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterItalicBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Italic")
        setStyleLable(My.Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub CheckBoxUnderline_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterUnderlineBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Underline")
        setStyleLable(My.Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub BookChapterColorBtn_Click(sender As Object, e As EventArgs) Handles BookChapterColorBtn.Click
        ColorDlg.Color = My.Settings.BookChapterForeColor
        ColorDlg.ShowDialog()
        My.Settings.BookChapterForeColor = ColorDlg.Color
        My.Settings.Save()
        BookChapterFontBtnn.ForeColor = My.Settings.BookChapterForeColor
        setPreviewDocument()
    End Sub

    Private Sub BookChapterBGColorBtn_Click(sender As Object, e As EventArgs) Handles BookChapterBGColorBtn.Click
        ColorDlg.Color = My.Settings.BookChapterBackColor
        ColorDlg.ShowDialog()
        My.Settings.BookChapterBackColor = ColorDlg.Color
        My.Settings.Save()
        BookChapterFontBtnn.BackColor = My.Settings.BookChapterBackColor
        setPreviewDocument()
    End Sub


    Private Sub VerseNumberFontBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberFontBtnn.Click
        FontDlg.Font = My.Settings.VerseNumberFont
        FontDlg.ShowDialog()
        My.Settings.VerseNumberFont = FontDlg.Font
        My.Settings.Save()

        setFontBtn("VerseNumber")

        setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")

        setCheckBtns("VerseNumber")

        setPreviewDocument()
    End Sub

    Private Sub VerseNumberColorBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberColorBtn.Click
        ColorDlg.Color = My.Settings.VerseNumberForeColor
        ColorDlg.ShowDialog()
        My.Settings.VerseNumberForeColor = ColorDlg.Color
        My.Settings.Save()
        VerseNumberFontBtnn.ForeColor = My.Settings.VerseNumberForeColor
        setPreviewDocument()
    End Sub

    Private Sub VerseNumberBGColorBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberBGColorBtn.Click
        ColorDlg.Color = My.Settings.VerseNumberBackColor
        ColorDlg.ShowDialog()
        My.Settings.VerseNumberBackColor = ColorDlg.Color
        My.Settings.Save()
        VerseNumberFontBtnn.BackColor = My.Settings.VerseNumberBackColor
        setPreviewDocument()
    End Sub

    Private Sub VerseNumberBoldBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberBoldBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Bold")
        setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberItalicBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberItalicBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Italic")
        setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberUnderlineBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberUnderlineBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Underline")
        setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberSuperscriptBtn.CheckedChanged
        If VerseNumberSuperscriptBtn.Checked Then
            VerseNumberSubscriptBtn.Checked = False
            My.Settings.VerseNumberVAlign = VALIGN.SUPERSCRIPT
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSuperscriptBtn.Checked = False And VerseNumberSubscriptBtn.Checked = False Then
            My.Settings.VerseNumberVAlign = VALIGN.NORMAL
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseNumberSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberSubscriptBtn.CheckedChanged
        If VerseNumberSubscriptBtn.Checked Then
            VerseNumberSuperscriptBtn.Checked = False
            My.Settings.VerseNumberVAlign = VALIGN.SUBSCRIPT
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSubscriptBtn.Checked = False And VerseNumberSuperscriptBtn.Checked = False Then
            My.Settings.VerseNumberVAlign = VALIGN.NORMAL
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseTextFontBtn_Click(sender As Object, e As EventArgs) Handles VerseTextFontBtnn.Click
        FontDlg.Font = My.Settings.VerseTextFont
        FontDlg.ShowDialog()
        My.Settings.VerseTextFont = FontDlg.Font
        My.Settings.Save()

        setFontBtn("VerseText")

        setStyleLable(My.Settings.VerseTextFont, "VerseText")

        setCheckBtns("VerseText")
        setPreviewDocument()
    End Sub

    Private Sub VerseTextColorBtn_Click(sender As Object, e As EventArgs) Handles VerseTextColorBtn.Click
        ColorDlg.Color = My.Settings.VerseTextForeColor
        ColorDlg.ShowDialog()
        My.Settings.VerseTextForeColor = ColorDlg.Color
        My.Settings.Save()
        VerseTextFontBtnn.ForeColor = My.Settings.VerseTextForeColor
        setPreviewDocument()
    End Sub

    Private Sub VerseTextBGColorBtn_Click(sender As Object, e As EventArgs) Handles VerseTextBGColorBtn.Click
        ColorDlg.Color = My.Settings.VerseTextBackColor
        ColorDlg.ShowDialog()
        My.Settings.VerseTextBackColor = ColorDlg.Color
        My.Settings.Save()
        VerseTextFontBtnn.BackColor = My.Settings.VerseTextBackColor
        setPreviewDocument()
    End Sub

    Private Sub VerseTextBoldBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextBoldBtn.CheckedChanged
        checkBoxChanged("VerseText", "Bold")
        setStyleLable(My.Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseTextItalicBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextItalicBtn.CheckedChanged
        checkBoxChanged("VerseText", "Italic")
        setStyleLable(My.Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseTextUnderlineBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextUnderlineBtn.CheckedChanged
        checkBoxChanged("VerseText", "Underline")
        setStyleLable(My.Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox1.Image = My.Resources.toggle_button_state_on
            My.Settings.NOVERSIONFORMATTING = True
        Else
            CheckBox1.Image = My.Resources.toggle_button_state_off
            My.Settings.NOVERSIONFORMATTING = False
        End If
        My.Settings.Save()
    End Sub

    Private Sub CheckBox1_MouseEnter(sender As Object, e As EventArgs) Handles CheckBox1.MouseEnter
        If CheckBox1.Checked Then
            CheckBox1.Image = My.Resources.toggle_button_state_on_hover
        Else
            CheckBox1.Image = My.Resources.toggle_button_state_off_hover
        End If
    End Sub

    Private Sub CheckBox1_MouseLeave(sender As Object, e As EventArgs) Handles CheckBox1.MouseLeave
        If CheckBox1.Checked Then
            CheckBox1.Image = My.Resources.toggle_button_state_on
        Else
            CheckBox1.Image = My.Resources.toggle_button_state_off
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            My.Settings.ParagraphAlignment = ALIGN.LEFT
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            My.Settings.ParagraphAlignment = ALIGN.CENTER
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            My.Settings.ParagraphAlignment = ALIGN.RIGHT
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            My.Settings.ParagraphAlignment = ALIGN.JUSTIFY
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub LeftIndentBtn_Click(sender As Object, e As EventArgs) Handles LeftIndentBtn.Click
        Dim indent As Single = My.Settings.LeftIndent
        If Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdCentimeters Then
            indent += 1.0F
            If indent > 8.0F Then indent = 8.0F
            Label2.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
        ElseIf Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdInches Then
            indent += 0.5F
            If indent > 3.0F Then indent = 3.0F
            Label2.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "in"
        End If
        My.Settings.LeftIndent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub LeftIndentBtn2_Click(sender As Object, e As EventArgs) Handles LeftIndentBtn2.Click
        Dim indent As Single = My.Settings.LeftIndent
        If Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdCentimeters Then
            indent -= 1.0F
            Label2.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
        ElseIf Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdInches Then
            indent -= 0.5F
            Label2.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "in"
        End If
        If indent < 0F Then indent = 0F
        My.Settings.LeftIndent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RightIndentBtn_Click(sender As Object, e As EventArgs) Handles RightIndentBtn.Click
        Dim indent As Single = My.Settings.RightIndent
        If Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdCentimeters Then
            indent += 1.0F
            If indent > 8.0F Then indent = 8.0F
            Label7.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
        ElseIf Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdInches Then
            indent += 0.5F
            If indent > 3.0F Then indent = 3.0F
            Label7.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "in"
        End If
        My.Settings.RightIndent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RightIndentBtn2_Click(sender As Object, e As EventArgs) Handles RightIndentBtn2.Click
        Dim indent As Single = My.Settings.RightIndent
        If Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdCentimeters Then
            indent -= 1.0F
            Label7.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "cm"
        ElseIf Application.Options.MeasurementUnit = Microsoft.Office.Interop.Word.WdMeasurementUnits.wdInches Then
            indent -= 0.5F
            Label7.Text = indent.ToString("F1", CultureInfo.InvariantCulture) & "in"
        End If
        If indent < 0F Then indent = 0F
        My.Settings.RightIndent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not initializing Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    My.Settings.Linespacing = 1.0F
                Case 1
                    My.Settings.Linespacing = 1.5F
                Case 2
                    My.Settings.Linespacing = 2.0F
            End Select
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "linespacing has been set to " + My.Settings.Linespacing.ToString)
            My.Settings.Save()
            setPreviewDocument()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            CheckBox2.Image = My.Resources.toggle_button_state_on
            My.Settings.BibleVersionVisibility = VISIBILITY.SHOW
        Else
            CheckBox2.Image = My.Resources.toggle_button_state_off
            My.Settings.BibleVersionVisibility = VISIBILITY.HIDE
        End If
        My.Settings.Save()
        If Not initializing Then setPreviewDocument()
        Label3.Visible = (My.Settings.BibleVersionVisibility = VISIBILITY.SHOW)
        Label4.Visible = (My.Settings.BibleVersionVisibility = VISIBILITY.HIDE)
    End Sub

    Private Sub CheckBox2_MouseEnter(sender As Object, e As EventArgs) Handles CheckBox2.MouseEnter
        If CheckBox2.Checked Then
            CheckBox2.Image = My.Resources.toggle_button_state_on_hover
        Else
            CheckBox2.Image = My.Resources.toggle_button_state_off_hover
        End If
    End Sub

    Private Sub CheckBox2_MouseLeave(sender As Object, e As EventArgs) Handles CheckBox2.MouseLeave
        If CheckBox2.Checked Then
            CheckBox2.Image = My.Resources.toggle_button_state_on
        Else
            CheckBox2.Image = My.Resources.toggle_button_state_off
        End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            CheckBox3.Image = My.Resources.toggle_button_state_on
            My.Settings.VerseNumberVisibility = VISIBILITY.SHOW
        Else
            CheckBox3.Image = My.Resources.toggle_button_state_off
            My.Settings.VerseNumberVisibility = VISIBILITY.HIDE
        End If
        My.Settings.Save()
        If Not initializing Then setPreviewDocument()
        Label6.Visible = (My.Settings.VerseNumberVisibility = VISIBILITY.SHOW)
        Label5.Visible = (My.Settings.VerseNumberVisibility = VISIBILITY.HIDE)
    End Sub

    Private Sub CheckBox3_MouseEnter(sender As Object, e As EventArgs) Handles CheckBox3.MouseEnter
        If CheckBox3.Checked Then
            CheckBox3.Image = My.Resources.toggle_button_state_on_hover
        Else
            CheckBox3.Image = My.Resources.toggle_button_state_off_hover
        End If
    End Sub

    Private Sub CheckBox3_MouseLeave(sender As Object, e As EventArgs) Handles CheckBox3.MouseLeave
        If CheckBox3.Checked Then
            CheckBox3.Image = My.Resources.toggle_button_state_on
        Else
            CheckBox3.Image = My.Resources.toggle_button_state_off
        End If
    End Sub

    Private Sub BibleVersionFontBtnn_Click(sender As Object, e As EventArgs) Handles BibleVersionFontBtnn.Click
        FontDlg.Font = My.Settings.BibleVersionFont
        FontDlg.ShowDialog()
        My.Settings.BibleVersionFont = FontDlg.Font
        My.Settings.Save()

        setFontBtn("BibleVersion")

        setStyleLable(My.Settings.BibleVersionFont, "BibleVersion")

        setCheckBtns("BibleVersion")
        setPreviewDocument()

    End Sub

    Private Sub RadioButton7_Click(sender As Object, e As EventArgs) Handles RadioButton7.Click
        My.Settings.BibleVersionAlign = ALIGN.LEFT
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles RadioButton6.Click
        My.Settings.BibleVersionAlign = ALIGN.CENTER
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton5_Click(sender As Object, e As EventArgs) Handles RadioButton5.Click
        My.Settings.BibleVersionAlign = ALIGN.RIGHT
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton13_Click(sender As Object, e As EventArgs) Handles RadioButton13.Click
        My.Settings.BibleVersionPosition = POS.TOP
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton12_Click(sender As Object, e As EventArgs) Handles RadioButton12.Click
        My.Settings.BibleVersionPosition = POS.BOTTOM
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton10_Click(sender As Object, e As EventArgs) Handles RadioButton10.Click
        My.Settings.BookChapterAlign = ALIGN.LEFT
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton9_Click(sender As Object, e As EventArgs) Handles RadioButton9.Click
        My.Settings.BookChapterAlign = ALIGN.CENTER
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton8_Click(sender As Object, e As EventArgs) Handles RadioButton8.Click
        My.Settings.BookChapterAlign = ALIGN.RIGHT
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton16_Click(sender As Object, e As EventArgs) Handles RadioButton16.Click
        My.Settings.BookChapterPosition = POS.TOP
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton15_Click(sender As Object, e As EventArgs) Handles RadioButton15.Click
        My.Settings.BookChapterPosition = POS.BOTTOM
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton14_Click(sender As Object, e As EventArgs) Handles RadioButton14.Click
        My.Settings.BookChapterPosition = POS.BOTTOMINLINE
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub BibleVersionBoldBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BibleVersionBoldBtn.CheckedChanged
        checkBoxChanged("BibleVersion", "Bold")
        setStyleLable(My.Settings.BibleVersionFont, "BibleVersion")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub BibleVersionItalicBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BibleVersionItalicBtn.CheckedChanged
        checkBoxChanged("BibleVersion", "Italic")
        setStyleLable(My.Settings.BibleVersionFont, "BibleVersion")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub BibleVersionUnderlineBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BibleVersionUnderlineBtn.CheckedChanged
        checkBoxChanged("BibleVersion", "Underline")
        setStyleLable(My.Settings.BibleVersionFont, "BibleVersion")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub RadioButton11_Click(sender As Object, e As EventArgs) Handles RadioButton11.Click
        My.Settings.BibleVersionWrap = WRAP.NONE
        setPreviewDocument()
    End Sub

    Private Sub RadioButton17_Click(sender As Object, e As EventArgs) Handles RadioButton17.Click
        My.Settings.BibleVersionWrap = WRAP.PARENTHESES
        setPreviewDocument()
    End Sub

    Private Sub RadioButton18_Click(sender As Object, e As EventArgs) Handles RadioButton18.Click
        My.Settings.BibleVersionWrap = WRAP.BRACKETS
        setPreviewDocument()
    End Sub

    Private Sub RadioButton21_Click(sender As Object, e As EventArgs) Handles RadioButton21.Click
        My.Settings.BookChapterWrap = WRAP.NONE
        setPreviewDocument()
    End Sub

    Private Sub RadioButton20_Click(sender As Object, e As EventArgs) Handles RadioButton20.Click
        My.Settings.BookChapterWrap = WRAP.PARENTHESES
        setPreviewDocument()
    End Sub

    Private Sub RadioButton19_Click(sender As Object, e As EventArgs) Handles RadioButton19.Click
        My.Settings.BookChapterWrap = WRAP.BRACKETS
        setPreviewDocument()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ColorDlg.Color = My.Settings.BibleVersionForeColor
        ColorDlg.ShowDialog()
        My.Settings.BibleVersionForeColor = ColorDlg.Color
        My.Settings.Save()
        BibleVersionFontBtnn.ForeColor = My.Settings.BibleVersionForeColor
        setPreviewDocument()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ColorDlg.Color = My.Settings.BibleVersionBackColor
        ColorDlg.ShowDialog()
        My.Settings.BibleVersionBackColor = ColorDlg.Color
        My.Settings.Save()
        BibleVersionFontBtnn.BackColor = My.Settings.BibleVersionBackColor
        setPreviewDocument()
    End Sub

    Private Sub RadioButton22_Click(sender As Object, e As EventArgs) Handles RadioButton22.Click
        My.Settings.BookChapterFormat = FORMAT.BIBLELANGABBREV
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton23_Click(sender As Object, e As EventArgs) Handles RadioButton23.Click
        My.Settings.BookChapterFormat = FORMAT.BIBLELANG
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton24_Click(sender As Object, e As EventArgs) Handles RadioButton24.Click
        My.Settings.BookChapterFormat = FORMAT.USERLANGABBREV
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub RadioButton25_Click(sender As Object, e As EventArgs) Handles RadioButton25.Click
        My.Settings.BookChapterFormat = FORMAT.USERLANG
        My.Settings.Save()
        setPreviewDocument()
    End Sub

End Class