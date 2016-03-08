'Imports mshtml
Imports System.Globalization
Imports System.Runtime.InteropServices

Public Class Preferences
    Private initializing As Boolean = True

    'turn off the annoying clicking sound when the preview window refreshes (WebBrowser control)
    Const DS As Integer = 21
    Const SP As Integer = &H2
    <DllImport("urlmon.dll")> _
    <PreserveSig> _
    Private Shared Function CoInternetSetFeatureEnabled(FeatureEntry As Integer, <MarshalAs(UnmanagedType.U4)> dSFlags As Integer, eEnable As Boolean) As <MarshalAs(UnmanagedType.[Error])> Integer
    End Function

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
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
        styleStr = fntSize.ToString & "pt   "
        styleStr &= If(myFont.Style And Drawing.FontStyle.Bold, " Bold", "")
        styleStr &= If(myFont.Style And Drawing.FontStyle.Italic, " Italic", "")
        styleStr &= If(myFont.Style And Drawing.FontStyle.Underline, " Underscore", "")
        styleStr &= If(((myFont.Style And Drawing.FontStyle.Bold) = False) And ((myFont.Style And Drawing.FontStyle.Italic) = False) And ((myFont.Style And Drawing.FontStyle.Underline) = False), " Normal", "")
        'styleStr &= If(My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Regular, " Normal", "")

        Select Case myCase
            Case "BookChapter"
                styleStr &= If(BookChapterSuperscriptBtn.Checked, " Superscript", "")
                styleStr &= If(BookChapterSubscriptBtn.Checked, " Subscript", "")
                BookChapterStyleLbl.Text = styleStr
            Case "VerseNumber"
                styleStr &= If(VerseNumberSuperscriptBtn.Checked, " Superscript", "")
                styleStr &= If(VerseNumberSubscriptBtn.Checked, " Subscript", "")
                VerseNumberStyleLbl.Text = styleStr
            Case "VerseText"
                styleStr &= If(VerseTextSuperscriptBtn.Checked, " Superscript", "")
                styleStr &= If(VerseTextSubscriptBtn.Checked, " Subscript", "")
                VerseTextStyleLbl.Text = styleStr
        End Select
    End Sub

    Private Sub setFontBtn(ByVal myCase As String)
        Select Case myCase
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
            Case "BookChapter"
                BookChapterBoldBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Bold)
                BookChapterItalicBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Italic)
                BookChapterUnderlineBtn.Checked = (My.Settings.BookChapterFont.Style And Drawing.FontStyle.Underline)
                BookChapterSuperscriptBtn.Checked = (My.Settings.BookChapterVAlign = "super")
                BookChapterSubscriptBtn.Checked = (My.Settings.BookChapterVAlign = "sub")
            Case "VerseNumber"
                VerseNumberBoldBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Bold)
                VerseNumberItalicBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Italic)
                VerseNumberUnderlineBtn.Checked = (My.Settings.VerseNumberFont.Style And Drawing.FontStyle.Underline)
                VerseNumberSuperscriptBtn.Checked = (My.Settings.VerseNumberVAlign = "super")
                VerseNumberSubscriptBtn.Checked = (My.Settings.VerseNumberVAlign = "sub")
            Case "VerseText"
                VerseTextBoldBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Bold)
                VerseTextItalicBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Italic)
                VerseTextUnderlineBtn.Checked = (My.Settings.VerseTextFont.Style And Drawing.FontStyle.Underline)
                VerseTextSuperscriptBtn.Checked = (My.Settings.VerseTextVAlign = "super")
                VerseTextSubscriptBtn.Checked = (My.Settings.VerseTextVAlign = "sub")
        End Select
    End Sub

    Private Sub checkBoxChanged(ByVal myCase As String, ByVal myBtn As String)
        Select Case myBtn
            Case "Bold"
                Select Case myCase
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

        Dim paragraphAlignment As String = My.Settings.ParagraphAlignment
        Dim paragraphLineSpacing As Decimal = My.Settings.Linespacing * 100
        Dim leftIndent As Short = My.Settings.Indent * 5

        Dim fontFamilyBookChapter As String = My.Settings.BookChapterFont.Name
        Dim fontSizeBookChapter As String = Math.Round(My.Settings.BookChapterFont.SizeInPoints).ToString
        Dim boldBookChapter As Boolean = BookChapterBoldBtn.Checked
        Dim italicBookChapter As Boolean = BookChapterItalicBtn.Checked
        Dim textColorBookChapter As String = If(Not My.Settings.BookChapterForeColor.IsEmpty, My.Settings.BookChapterForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorBookChapter As String = If(Not My.Settings.BookChapterBackColor.IsEmpty, My.Settings.BookChapterBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim vAlignBookChapter As String = My.Settings.BookChapterVAlign
        'System.Diagnostics.Debug.WriteLine(vAlignBookChapter)

        Dim fontFamilyVerseNumber As String = My.Settings.VerseNumberFont.Name
        Dim fontSizeVerseNumber As String = Math.Round(My.Settings.VerseNumberFont.SizeInPoints).ToString
        Dim boldVerseNumber As Boolean = VerseNumberBoldBtn.Checked
        Dim italicVerseNumber As Boolean = VerseNumberItalicBtn.Checked
        Dim textColorVerseNumber As String = If(Not My.Settings.VerseNumberForeColor.IsEmpty, My.Settings.VerseNumberForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorVerseNumber As String = If(Not My.Settings.VerseNumberBackColor.IsEmpty, My.Settings.VerseNumberBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        'System.Diagnostics.Debug.WriteLine(bgColorVerseNumber)
        Dim vAlignVerseNumber As String = My.Settings.VerseNumberVAlign
        'System.Diagnostics.Debug.WriteLine(vAlignVerseNumber)

        Dim fontFamilyVerseText As String = My.Settings.VerseTextFont.Name
        Dim fontSizeVerseText As String = Math.Round(My.Settings.VerseTextFont.SizeInPoints).ToString
        Dim boldVerseText As Boolean = VerseTextBoldBtn.Checked
        Dim italicVerseText As Boolean = VerseTextItalicBtn.Checked
        Dim textColorVerseText As String = If(Not My.Settings.VerseTextForeColor.IsEmpty, My.Settings.VerseTextForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorVerseText As String = If(Not My.Settings.VerseTextBackColor.IsEmpty, My.Settings.VerseTextBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim vAlignVerseText As String = My.Settings.VerseTextVAlign
        'System.Diagnostics.Debug.WriteLine(vAlignVerseText)

        previewDocument = "<!DOCTYPE html>"
        previewDocument &= "<head>"
        previewDocument &= "<meta charset=""UTF-8"">"
        previewDocument &= "<style type=""text/css"">"

        stylesheet = "body { padding: 6px; background-color: #FFFFFF; }"
        stylesheet &= "div.results { line-height: " & paragraphLineSpacing & "%; }"
        stylesheet &= "div.results { margin-left: " & leftIndent & "pt; }"
        stylesheet &= "div.results p.book { font-family: " & fontFamilyBookChapter & "; }"
        stylesheet &= "div.results p.book { font-size: " & fontSizeBookChapter & "pt; }"
        stylesheet &= "div.results p.book { font-weight: " & If(boldBookChapter, "bold", "normal") & "; }"
        stylesheet &= "div.results p.book { font-style: " & If(italicBookChapter, "italic", "normal") & "; }"
        stylesheet &= "div.results p.book { color: #" & textColorBookChapter & "; }"
        stylesheet &= "div.results p.book { background-color: #" & bgColorBookChapter & "; }"
        stylesheet &= "div.results p.book span { vertical-align: " & vAlignBookChapter & "; }"
        stylesheet &= "div.results p.verses { text-align: " & paragraphAlignment & "; }"
        stylesheet &= "div.results p.verses span.sup { font-family: " & fontFamilyVerseNumber & "; }"
        stylesheet &= "div.results p.verses span.sup { font-size:" & fontSizeVerseNumber & "pt; }"
        stylesheet &= "div.results p.verses span.sup { font-weight: " & If(boldVerseNumber, "bold", "normal") & "; }"
        stylesheet &= "div.results p.verses span.sup { font-style: " & If(italicVerseNumber, "italic", "normal") & "; }"
        stylesheet &= "div.results p.verses span.sup { color: #" & textColorVerseNumber & "; }"
        stylesheet &= "div.results p.verses span.sup { background-color: #" & bgColorVerseNumber & "; }"
        stylesheet &= "div.results p.verses span.sup { vertical-align: " & vAlignVerseNumber & "; }"
        stylesheet &= "div.results p.verses span.text { font-family: " & fontFamilyVerseText & "; }"
        stylesheet &= "div.results p.verses span.text { font-size:" & fontSizeVerseText & "pt; }"
        stylesheet &= "div.results p.verses span.text { font-weight: " & If(boldVerseText, "bold", "normal") & "; }"
        stylesheet &= "div.results p.verses span.text { font-style: " & If(italicVerseText, "italic", "normal") & "; }"
        stylesheet &= "div.results p.verses span.text { color: #" & textColorVerseText & "; }"
        stylesheet &= "div.results p.verses span.text { background-color: #" & bgColorVerseText & "; }"
        stylesheet &= "div.results p.verses span.text { vertical-align: " & vAlignVerseText & "; }"

        previewDocument &= stylesheet
        previewDocument &= "</style>"
        previewDocument &= "</head>"
        previewDocument &= "<body>"
        previewDocument &= "<div class=""results""><p class=""book""><span>"
        previewDocument &= __("Genesis") & "&nbsp;1"
        previewDocument &= "</span></p><p class=""verses"" style=""margin-top:0px;""><span class=""sup"">1</span><span class=""text"">"
        previewDocument &= __("In the beginning, when God created the heavens and the earth")
        previewDocument &= "</span><span class=""sup"">2</span><span class=""text"">"
        previewDocument &= __("and the earth was without form or shape, with darkness over the abyss and a mighty wind sweeping over the waters")
        previewDocument &= "</span><span class=""sup"">3</span><span class=""text"">"
        previewDocument &= __("Then God said: Let there be light, and there was light.")
        previewDocument &= "</span></p></div>"
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

        Me.Text = __("User Preferences")

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
            Case 1.0
                ComboBox1.SelectedIndex = 0
            Case 1.5
                ComboBox1.SelectedIndex = 1
            Case 2.0
                ComboBox1.SelectedIndex = 2
        End Select
        GroupBox5.Text = __("Paragraph")
        GroupBox6.Text = __("Alignment")
        GroupBox7.Text = __("Indent")
        GroupBox8.Text = __("Line-spacing")
        GroupBox9.Text = __("Override Bible Version Formatting")
        GroupBox1.Text = __("Book / Chapter")
        GroupBox2.Text = __("Verse Number")
        GroupBox3.Text = __("Verse Text")
        GroupBox4.Text = __("Preview")
        ToolTip1.SetToolTip(Label1, __("Some Bible versions have their own formatting. This is left by default to keep the text as close as possible to the original.<br> If however you need to have consistent formatting in your document, you may override the Bible version's own formatting."))

        Select Case My.Settings.ParagraphAlignment
            Case "left"
                RadioButton1.Checked = True
            Case "center"
                RadioButton2.Checked = True
            Case "right"
                RadioButton3.Checked = True
            Case "justify"
                RadioButton4.Checked = True
        End Select

        CheckBox1.Checked = My.Settings.NOVERSIONFORMATTING

        CoInternetSetFeatureEnabled(DS, SP, True)

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

    Private Sub BookChapterSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterSuperscriptBtn.CheckedChanged
        If BookChapterSuperscriptBtn.Checked Then
            BookChapterSubscriptBtn.Checked = False
            My.Settings.BookChapterVAlign = "super"
            My.Settings.Save()
            setStyleLable(My.Settings.BookChapterFont, "BookChapter")
            If Not initializing Then setPreviewDocument()
        ElseIf BookChapterSuperscriptBtn.Checked = False And BookChapterSubscriptBtn.Checked = False Then
            My.Settings.BookChapterVAlign = "baseline"
            My.Settings.Save()
            setStyleLable(My.Settings.BookChapterFont, "BookChapter")
            setPreviewDocument()
        End If
    End Sub

    Private Sub BookChapterSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterSubscriptBtn.CheckedChanged
        If BookChapterSubscriptBtn.Checked Then
            BookChapterSuperscriptBtn.Checked = False
            My.Settings.BookChapterVAlign = "sub"
            My.Settings.Save()
            setStyleLable(My.Settings.BookChapterFont, "BookChapter")
            If Not initializing Then setPreviewDocument()
        ElseIf BookChapterSubscriptBtn.Checked = False And BookChapterSuperscriptBtn.Checked = False Then
            My.Settings.BookChapterVAlign = "baseline"
            My.Settings.Save()
            setStyleLable(My.Settings.BookChapterFont, "BookChapter")
            setPreviewDocument()
        End If
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
            My.Settings.VerseNumberVAlign = "super"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSuperscriptBtn.Checked = False And VerseNumberSubscriptBtn.Checked = False Then
            My.Settings.VerseNumberVAlign = "baseline"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseNumberSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberSubscriptBtn.CheckedChanged
        If VerseNumberSubscriptBtn.Checked Then
            VerseNumberSuperscriptBtn.Checked = False
            My.Settings.VerseNumberVAlign = "sub"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSubscriptBtn.Checked = False And VerseNumberSuperscriptBtn.Checked = False Then
            My.Settings.VerseNumberVAlign = "baseline"
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

    Private Sub VerseTextSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextSuperscriptBtn.CheckedChanged
        If VerseTextSuperscriptBtn.Checked Then
            VerseTextSubscriptBtn.Checked = False
            My.Settings.VerseTextVAlign = "super"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseTextFont, "VerseText")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseTextSuperscriptBtn.Checked = False And VerseTextSubscriptBtn.Checked = False Then
            My.Settings.VerseTextVAlign = "baseline"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseTextFont, "VerseText")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseTextSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextSubscriptBtn.CheckedChanged
        If VerseTextSubscriptBtn.Checked Then
            VerseTextSuperscriptBtn.Checked = False
            My.Settings.VerseTextVAlign = "sub"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseTextFont, "VerseText")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseTextSubscriptBtn.Checked = False And VerseTextSuperscriptBtn.Checked = False Then
            My.Settings.VerseTextVAlign = "baseline"
            My.Settings.Save()
            setStyleLable(My.Settings.VerseTextFont, "VerseText")
            setPreviewDocument()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox1.Text = "ON"
            My.Settings.NOVERSIONFORMATTING = True
            CheckBox1.ForeColor = Drawing.Color.DarkGreen
        Else
            CheckBox1.Text = "OFF"
            My.Settings.NOVERSIONFORMATTING = False
            CheckBox1.ForeColor = Drawing.Color.DarkRed
        End If
        My.Settings.Save()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            My.Settings.ParagraphAlignment = "left"
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            My.Settings.ParagraphAlignment = "center"
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            My.Settings.ParagraphAlignment = "right"
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            My.Settings.ParagraphAlignment = "justify"
            My.Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim indent As Short = My.Settings.Indent
        indent += 1
        If indent > 20 Then indent = 20
        My.Settings.Indent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim indent As Short = My.Settings.Indent
        indent -= 1
        If indent < 0 Then indent = 0
        My.Settings.Indent = indent
        My.Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not initializing Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    My.Settings.Linespacing = 1.0
                Case 1
                    My.Settings.Linespacing = 1.5
                Case 2
                    My.Settings.Linespacing = 2.0
            End Select
            'Diagnostics.Debug.WriteLine("linespacing has been set to " + My.Settings.Linespacing.ToString)
            My.Settings.Save()
            setPreviewDocument()
        End If
    End Sub

End Class