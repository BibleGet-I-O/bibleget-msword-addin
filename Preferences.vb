'Imports mshtml
Imports System.Globalization
Imports System.Runtime.InteropServices

Public Class Preferences
    Private initializing As Boolean = True
    Private Settings As New MySettings

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
        'styleStr &= If(Settings.VerseNumberFont.Style And Drawing.FontStyle.Regular, " Normal", "")

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
                BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                BookChapterFontBtn.Text = Settings.BookChapterFont.Name
                BookChapterFontBtn.ForeColor = Settings.BookChapterForeColor
                BookChapterFontBtn.BackColor = Settings.BookChapterBackColor
            Case "VerseNumber"
                VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                VerseNumberFontBtn.Text = Settings.VerseNumberFont.Name
                VerseNumberFontBtn.ForeColor = Settings.VerseNumberForeColor
                VerseNumberFontBtn.BackColor = Settings.VerseNumberBackColor
            Case "VerseText"
                VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                VerseTextFontBtn.Text = Settings.VerseTextFont.Name
                VerseTextFontBtn.ForeColor = Settings.VerseTextForeColor
                VerseTextFontBtn.BackColor = Settings.VerseTextBackColor
        End Select

    End Sub

    Private Sub setCheckBtns(ByVal myCase As String)
        Select Case myCase
            Case "BookChapter"
                BookChapterBoldBtn.Checked = (Settings.BookChapterFont.Style And Drawing.FontStyle.Bold)
                BookChapterItalicBtn.Checked = (Settings.BookChapterFont.Style And Drawing.FontStyle.Italic)
                BookChapterUnderlineBtn.Checked = (Settings.BookChapterFont.Style And Drawing.FontStyle.Underline)
                BookChapterSuperscriptBtn.Checked = (Settings.BookChapterVAlign = "super")
                BookChapterSubscriptBtn.Checked = (Settings.BookChapterVAlign = "sub")
            Case "VerseNumber"
                VerseNumberBoldBtn.Checked = (Settings.VerseNumberFont.Style And Drawing.FontStyle.Bold)
                VerseNumberItalicBtn.Checked = (Settings.VerseNumberFont.Style And Drawing.FontStyle.Italic)
                VerseNumberUnderlineBtn.Checked = (Settings.VerseNumberFont.Style And Drawing.FontStyle.Underline)
                VerseNumberSuperscriptBtn.Checked = (Settings.VerseNumberVAlign = "super")
                VerseNumberSubscriptBtn.Checked = (Settings.VerseNumberVAlign = "sub")
            Case "VerseText"
                VerseTextBoldBtn.Checked = (Settings.VerseTextFont.Style And Drawing.FontStyle.Bold)
                VerseTextItalicBtn.Checked = (Settings.VerseTextFont.Style And Drawing.FontStyle.Italic)
                VerseTextUnderlineBtn.Checked = (Settings.VerseTextFont.Style And Drawing.FontStyle.Underline)
                VerseTextSuperscriptBtn.Checked = (Settings.VerseTextVAlign = "super")
                VerseTextSubscriptBtn.Checked = (Settings.VerseTextVAlign = "sub")
        End Select
    End Sub

    Private Sub checkBoxChanged(ByVal myCase As String, ByVal myBtn As String)
        Select Case myBtn
            Case "Bold"
                Select Case myCase
                    Case "BookChapter"
                        If BookChapterBoldBtn.Checked Then
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style Or Drawing.FontStyle.Bold)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style And Not Drawing.FontStyle.Bold)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberBoldBtn.Checked Then
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style Or Drawing.FontStyle.Bold)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Bold)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextBoldBtn.Checked Then
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style Or Drawing.FontStyle.Bold)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style And Not Drawing.FontStyle.Bold)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
            Case "Italic"
                Select Case myCase
                    Case "BookChapter"
                        If BookChapterItalicBtn.Checked Then
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style Or Drawing.FontStyle.Italic)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style And Not Drawing.FontStyle.Italic)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberItalicBtn.Checked Then
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style Or Drawing.FontStyle.Italic)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Italic)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextItalicBtn.Checked Then
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style Or Drawing.FontStyle.Italic)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style And Not Drawing.FontStyle.Italic)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
            Case "Underline"
                Select Case myCase
                    Case "BookChapter"
                        If BookChapterUnderlineBtn.Checked Then
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style Or Drawing.FontStyle.Underline)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.BookChapterFont = New Drawing.Font(Settings.BookChapterFont.Name, Settings.BookChapterFont.Size, Settings.BookChapterFont.Style And Not Drawing.FontStyle.Underline)
                            BookChapterFontBtn.Font = New Drawing.Font(Settings.BookChapterFont.Name, 12, Settings.BookChapterFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseNumber"
                        If VerseNumberUnderlineBtn.Checked Then
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style Or Drawing.FontStyle.Underline)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseNumberFont = New Drawing.Font(Settings.VerseNumberFont.Name, Settings.VerseNumberFont.Size, Settings.VerseNumberFont.Style And Not Drawing.FontStyle.Underline)
                            VerseNumberFontBtn.Font = New Drawing.Font(Settings.VerseNumberFont.Name, 12, Settings.VerseNumberFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                    Case "VerseText"
                        If VerseTextUnderlineBtn.Checked Then
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style Or Drawing.FontStyle.Underline)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        Else
                            Settings.VerseTextFont = New Drawing.Font(Settings.VerseTextFont.Name, Settings.VerseTextFont.Size, Settings.VerseTextFont.Style And Not Drawing.FontStyle.Underline)
                            VerseTextFontBtn.Font = New Drawing.Font(Settings.VerseTextFont.Name, 12, Settings.VerseTextFont.Style, Drawing.GraphicsUnit.Point)
                        End If
                End Select
        End Select
        Settings.Save()
    End Sub

    Private Sub setPreviewDocument()
        Dim previewDocument As String
        Dim stylesheet As String

        Dim paragraphAlignment As String = Settings.ParagraphAlignment
        Dim paragraphLineSpacing As Decimal = Settings.Linespacing * 100
        Dim leftIndent As Short = Settings.Indent * 5

        Dim fontFamilyBookChapter As String = Settings.BookChapterFont.Name
        Dim fontSizeBookChapter As String = Math.Round(Settings.BookChapterFont.SizeInPoints).ToString
        Dim boldBookChapter As Boolean = BookChapterBoldBtn.Checked
        Dim italicBookChapter As Boolean = BookChapterItalicBtn.Checked
        Dim textColorBookChapter As String = If(Not Settings.BookChapterForeColor.IsEmpty, Settings.BookChapterForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorBookChapter As String = If(Not Settings.BookChapterBackColor.IsEmpty, Settings.BookChapterBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim vAlignBookChapter As String = Settings.BookChapterVAlign
        'System.Diagnostics.Debug.WriteLine(vAlignBookChapter)

        Dim fontFamilyVerseNumber As String = Settings.VerseNumberFont.Name
        Dim fontSizeVerseNumber As String = Math.Round(Settings.VerseNumberFont.SizeInPoints).ToString
        Dim boldVerseNumber As Boolean = VerseNumberBoldBtn.Checked
        Dim italicVerseNumber As Boolean = VerseNumberItalicBtn.Checked
        Dim textColorVerseNumber As String = If(Not Settings.VerseNumberForeColor.IsEmpty, Settings.VerseNumberForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorVerseNumber As String = If(Not Settings.VerseNumberBackColor.IsEmpty, Settings.VerseNumberBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        'System.Diagnostics.Debug.WriteLine(bgColorVerseNumber)
        Dim vAlignVerseNumber As String = Settings.VerseNumberVAlign
        'System.Diagnostics.Debug.WriteLine(vAlignVerseNumber)

        Dim fontFamilyVerseText As String = Settings.VerseTextFont.Name
        Dim fontSizeVerseText As String = Math.Round(Settings.VerseTextFont.SizeInPoints).ToString
        Dim boldVerseText As Boolean = VerseTextBoldBtn.Checked
        Dim italicVerseText As Boolean = VerseTextItalicBtn.Checked
        Dim textColorVerseText As String = If(Not Settings.VerseTextForeColor.IsEmpty, Settings.VerseTextForeColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim bgColorVerseText As String = If(Not Settings.VerseTextBackColor.IsEmpty, Settings.VerseTextBackColor.ToArgb().ToString("X").Substring(2), "transparent")
        Dim vAlignVerseText As String = Settings.VerseTextVAlign
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
        setStyleLable(Settings.BookChapterFont, "BookChapter")
        setCheckBtns("BookChapter")

        setFontBtn("VerseNumber")
        setStyleLable(Settings.VerseNumberFont, "VerseNumber")
        setCheckBtns("VerseNumber")

        setFontBtn("VerseText")
        setStyleLable(Settings.VerseTextFont, "VerseText")
        setCheckBtns("VerseText")

        Select Case Settings.Linespacing
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

        Select Case Settings.ParagraphAlignment
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

    Private Sub BookChapterFontBtn_Click(sender As Object, e As EventArgs) Handles BookChapterFontBtn.Click
        FontDlg.Font = Settings.BookChapterFont
        FontDlg.ShowDialog()
        Settings.BookChapterFont = FontDlg.Font
        Settings.Save()

        setFontBtn("BookChapter")

        setStyleLable(Settings.BookChapterFont, "BookChapter")

        setCheckBtns("BookChapter")
        setPreviewDocument()
    End Sub

    Private Sub CheckBoxBold_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterBoldBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Bold")
        setStyleLable(Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub CheckBoxItalic_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterItalicBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Italic")
        setStyleLable(Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub CheckBoxUnderline_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterUnderlineBtn.CheckedChanged
        checkBoxChanged("BookChapter", "Underline")
        setStyleLable(Settings.BookChapterFont, "BookChapter")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub BookChapterColorBtn_Click(sender As Object, e As EventArgs) Handles BookChapterColorBtn.Click
        ColorDlg.Color = Settings.BookChapterForeColor
        ColorDlg.ShowDialog()
        Settings.BookChapterForeColor = ColorDlg.Color
        Settings.Save()
        BookChapterFontBtn.ForeColor = Settings.BookChapterForeColor
        setPreviewDocument()
    End Sub

    Private Sub BookChapterBGColorBtn_Click(sender As Object, e As EventArgs) Handles BookChapterBGColorBtn.Click
        ColorDlg.Color = Settings.BookChapterBackColor
        ColorDlg.ShowDialog()
        Settings.BookChapterBackColor = ColorDlg.Color
        Settings.Save()
        BookChapterFontBtn.BackColor = Settings.BookChapterBackColor
        setPreviewDocument()
    End Sub

    Private Sub BookChapterSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterSuperscriptBtn.CheckedChanged
        If BookChapterSuperscriptBtn.Checked Then
            BookChapterSubscriptBtn.Checked = False
            Settings.BookChapterVAlign = "super"
            Settings.Save()
            setStyleLable(Settings.BookChapterFont, "BookChapter")
            If Not initializing Then setPreviewDocument()
        ElseIf BookChapterSuperscriptBtn.Checked = False And BookChapterSubscriptBtn.Checked = False Then
            Settings.BookChapterVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.BookChapterFont, "BookChapter")
            setPreviewDocument()
        End If
    End Sub

    Private Sub BookChapterSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles BookChapterSubscriptBtn.CheckedChanged
        If BookChapterSubscriptBtn.Checked Then
            BookChapterSuperscriptBtn.Checked = False
            Settings.BookChapterVAlign = "sub"
            Settings.Save()
            setStyleLable(Settings.BookChapterFont, "BookChapter")
            If Not initializing Then setPreviewDocument()
        ElseIf BookChapterSubscriptBtn.Checked = False And BookChapterSuperscriptBtn.Checked = False Then
            Settings.BookChapterVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.BookChapterFont, "BookChapter")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseNumberFontBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberFontBtn.Click
        FontDlg.Font = Settings.VerseNumberFont
        FontDlg.ShowDialog()
        Settings.VerseNumberFont = FontDlg.Font
        Settings.Save()

        setFontBtn("VerseNumber")

        setStyleLable(Settings.VerseNumberFont, "VerseNumber")

        setCheckBtns("VerseNumber")

        setPreviewDocument()
    End Sub

    Private Sub VerseNumberColorBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberColorBtn.Click
        ColorDlg.Color = Settings.VerseNumberForeColor
        ColorDlg.ShowDialog()
        Settings.VerseNumberForeColor = ColorDlg.Color
        Settings.Save()
        VerseNumberFontBtn.ForeColor = Settings.VerseNumberForeColor
        setPreviewDocument()
    End Sub

    Private Sub VerseNumberBGColorBtn_Click(sender As Object, e As EventArgs) Handles VerseNumberBGColorBtn.Click
        ColorDlg.Color = Settings.VerseNumberBackColor
        ColorDlg.ShowDialog()
        Settings.VerseNumberBackColor = ColorDlg.Color
        Settings.Save()
        VerseNumberFontBtn.BackColor = Settings.VerseNumberBackColor
        setPreviewDocument()
    End Sub

    Private Sub VerseNumberBoldBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberBoldBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Bold")
        setStyleLable(Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberItalicBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberItalicBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Italic")
        setStyleLable(Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberUnderlineBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberUnderlineBtn.CheckedChanged
        checkBoxChanged("VerseNumber", "Underline")
        setStyleLable(Settings.VerseNumberFont, "VerseNumber")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseNumberSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberSuperscriptBtn.CheckedChanged
        If VerseNumberSuperscriptBtn.Checked Then
            VerseNumberSubscriptBtn.Checked = False
            Settings.VerseNumberVAlign = "super"
            Settings.Save()
            setStyleLable(Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSuperscriptBtn.Checked = False And VerseNumberSubscriptBtn.Checked = False Then
            Settings.VerseNumberVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.VerseNumberFont, "VerseNumber")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseNumberSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseNumberSubscriptBtn.CheckedChanged
        If VerseNumberSubscriptBtn.Checked Then
            VerseNumberSuperscriptBtn.Checked = False
            Settings.VerseNumberVAlign = "sub"
            Settings.Save()
            setStyleLable(Settings.VerseNumberFont, "VerseNumber")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseNumberSubscriptBtn.Checked = False And VerseNumberSuperscriptBtn.Checked = False Then
            Settings.VerseNumberVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.VerseNumberFont, "VerseNumber")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseTextFontBtn_Click(sender As Object, e As EventArgs) Handles VerseTextFontBtn.Click
        FontDlg.Font = Settings.VerseTextFont
        FontDlg.ShowDialog()
        Settings.VerseTextFont = FontDlg.Font
        Settings.Save()

        setFontBtn("VerseText")

        setStyleLable(Settings.VerseTextFont, "VerseText")

        setCheckBtns("VerseText")
        setPreviewDocument()
    End Sub

    Private Sub VerseTextColorBtn_Click(sender As Object, e As EventArgs) Handles VerseTextColorBtn.Click
        ColorDlg.Color = Settings.VerseTextForeColor
        ColorDlg.ShowDialog()
        Settings.VerseTextForeColor = ColorDlg.Color
        Settings.Save()
        VerseTextFontBtn.ForeColor = Settings.VerseTextForeColor
        setPreviewDocument()
    End Sub

    Private Sub VerseTextBGColorBtn_Click(sender As Object, e As EventArgs) Handles VerseTextBGColorBtn.Click
        ColorDlg.Color = Settings.VerseTextBackColor
        ColorDlg.ShowDialog()
        Settings.VerseTextBackColor = ColorDlg.Color
        Settings.Save()
        VerseTextFontBtn.BackColor = Settings.VerseTextBackColor
        setPreviewDocument()
    End Sub

    Private Sub VerseTextBoldBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextBoldBtn.CheckedChanged
        checkBoxChanged("VerseText", "Bold")
        setStyleLable(Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseTextItalicBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextItalicBtn.CheckedChanged
        checkBoxChanged("VerseText", "Italic")
        setStyleLable(Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseTextUnderlineBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextUnderlineBtn.CheckedChanged
        checkBoxChanged("VerseText", "Underline")
        setStyleLable(Settings.VerseTextFont, "VerseText")
        If Not initializing Then setPreviewDocument()
    End Sub

    Private Sub VerseTextSuperscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextSuperscriptBtn.CheckedChanged
        If VerseTextSuperscriptBtn.Checked Then
            VerseTextSubscriptBtn.Checked = False
            Settings.VerseTextVAlign = "super"
            Settings.Save()
            setStyleLable(Settings.VerseTextFont, "VerseText")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseTextSuperscriptBtn.Checked = False And VerseTextSubscriptBtn.Checked = False Then
            Settings.VerseTextVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.VerseTextFont, "VerseText")
            setPreviewDocument()
        End If
    End Sub

    Private Sub VerseTextSubscriptBtn_CheckedChanged(sender As Object, e As EventArgs) Handles VerseTextSubscriptBtn.CheckedChanged
        If VerseTextSubscriptBtn.Checked Then
            VerseTextSuperscriptBtn.Checked = False
            Settings.VerseTextVAlign = "sub"
            Settings.Save()
            setStyleLable(Settings.VerseTextFont, "VerseText")
            If Not initializing Then setPreviewDocument()
        ElseIf VerseTextSubscriptBtn.Checked = False And VerseTextSuperscriptBtn.Checked = False Then
            Settings.VerseTextVAlign = "baseline"
            Settings.Save()
            setStyleLable(Settings.VerseTextFont, "VerseText")
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
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Settings.ParagraphAlignment = "left"
            Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Settings.ParagraphAlignment = "center"
            Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            Settings.ParagraphAlignment = "right"
            Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            Settings.ParagraphAlignment = "justify"
            Settings.Save()
            If Not initializing Then setPreviewDocument()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim indent As Short = Settings.Indent
        indent += 1
        If indent > 20 Then indent = 20
        Settings.Indent = indent
        Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim indent As Short = Settings.Indent
        indent -= 1
        If indent < 0 Then indent = 0
        Settings.Indent = indent
        Settings.Save()
        setPreviewDocument()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not initializing Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    Settings.Linespacing = 1.0
                Case 1
                    Settings.Linespacing = 1.5
                Case 2
                    Settings.Linespacing = 2.0
            End Select
            Settings.Save()
            setPreviewDocument()
        End If
    End Sub
End Class