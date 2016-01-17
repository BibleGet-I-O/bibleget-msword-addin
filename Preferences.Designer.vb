<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Preferences
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Preferences))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BookChapterFontBtnn = New System.Windows.Forms.Button()
        Me.BookChapterSubscriptBtn = New System.Windows.Forms.CheckBox()
        Me.BookChapterSuperscriptBtn = New System.Windows.Forms.CheckBox()
        Me.BookChapterStyleLbl = New System.Windows.Forms.Label()
        Me.BookChapterUnderlineBtn = New System.Windows.Forms.CheckBox()
        Me.BookChapterItalicBtn = New System.Windows.Forms.CheckBox()
        Me.BookChapterBoldBtn = New System.Windows.Forms.CheckBox()
        Me.BookChapterBGColorBtn = New System.Windows.Forms.Button()
        Me.BookChapterColorBtn = New System.Windows.Forms.Button()
        Me.FontDlg = New System.Windows.Forms.FontDialog()
        Me.ColorDlg = New System.Windows.Forms.ColorDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.VerseNumberFontBtnn = New System.Windows.Forms.Button()
        Me.VerseNumberSubscriptBtn = New System.Windows.Forms.CheckBox()
        Me.VerseNumberSuperscriptBtn = New System.Windows.Forms.CheckBox()
        Me.VerseNumberStyleLbl = New System.Windows.Forms.Label()
        Me.VerseNumberUnderlineBtn = New System.Windows.Forms.CheckBox()
        Me.VerseNumberItalicBtn = New System.Windows.Forms.CheckBox()
        Me.VerseNumberBoldBtn = New System.Windows.Forms.CheckBox()
        Me.VerseNumberBGColorBtn = New System.Windows.Forms.Button()
        Me.VerseNumberColorBtn = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.VerseTextFontBtnn = New System.Windows.Forms.Button()
        Me.VerseTextSubscriptBtn = New System.Windows.Forms.CheckBox()
        Me.VerseTextSuperscriptBtn = New System.Windows.Forms.CheckBox()
        Me.VerseTextStyleLbl = New System.Windows.Forms.Label()
        Me.VerseTextUnderlineBtn = New System.Windows.Forms.CheckBox()
        Me.VerseTextItalicBtn = New System.Windows.Forms.CheckBox()
        Me.VerseTextBoldBtn = New System.Windows.Forms.CheckBox()
        Me.VerseTextBGColorBtn = New System.Windows.Forms.Button()
        Me.VerseTextColorBtn = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BookChapterFontBtnn)
        Me.GroupBox1.Controls.Add(Me.BookChapterSubscriptBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterSuperscriptBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterStyleLbl)
        Me.GroupBox1.Controls.Add(Me.BookChapterUnderlineBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterItalicBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterBoldBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterBGColorBtn)
        Me.GroupBox1.Controls.Add(Me.BookChapterColorBtn)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 126)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(633, 75)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Book / Chapter Formatting"
        '
        'BookChapterFontBtnn
        '
        Me.BookChapterFontBtnn.Location = New System.Drawing.Point(6, 16)
        Me.BookChapterFontBtnn.Name = "BookChapterFontBtnn"
        Me.BookChapterFontBtnn.Size = New System.Drawing.Size(240, 35)
        Me.BookChapterFontBtnn.TabIndex = 11
        Me.BookChapterFontBtnn.Text = "Button3"
        Me.BookChapterFontBtnn.UseVisualStyleBackColor = True
        '
        'BookChapterSubscriptBtn
        '
        Me.BookChapterSubscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.BookChapterSubscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.superscript
        Me.BookChapterSubscriptBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BookChapterSubscriptBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.BookChapterSubscriptBtn.Location = New System.Drawing.Point(554, 16)
        Me.BookChapterSubscriptBtn.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.BookChapterSubscriptBtn.Name = "BookChapterSubscriptBtn"
        Me.BookChapterSubscriptBtn.Padding = New System.Windows.Forms.Padding(0, 2, 0, 0)
        Me.BookChapterSubscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterSubscriptBtn.TabIndex = 10
        Me.BookChapterSubscriptBtn.UseVisualStyleBackColor = True
        '
        'BookChapterSuperscriptBtn
        '
        Me.BookChapterSuperscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.BookChapterSuperscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.subscript
        Me.BookChapterSuperscriptBtn.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BookChapterSuperscriptBtn.Location = New System.Drawing.Point(507, 16)
        Me.BookChapterSuperscriptBtn.Margin = New System.Windows.Forms.Padding(6, 3, 0, 3)
        Me.BookChapterSuperscriptBtn.Name = "BookChapterSuperscriptBtn"
        Me.BookChapterSuperscriptBtn.Padding = New System.Windows.Forms.Padding(0, 0, 0, 1)
        Me.BookChapterSuperscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterSuperscriptBtn.TabIndex = 9
        Me.BookChapterSuperscriptBtn.UseVisualStyleBackColor = True
        '
        'BookChapterStyleLbl
        '
        Me.BookChapterStyleLbl.AutoSize = True
        Me.BookChapterStyleLbl.BackColor = System.Drawing.SystemColors.Control
        Me.BookChapterStyleLbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.BookChapterStyleLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BookChapterStyleLbl.Location = New System.Drawing.Point(7, 54)
        Me.BookChapterStyleLbl.Name = "BookChapterStyleLbl"
        Me.BookChapterStyleLbl.Size = New System.Drawing.Size(60, 15)
        Me.BookChapterStyleLbl.TabIndex = 8
        Me.BookChapterStyleLbl.Text = "12pt   Bold"
        '
        'BookChapterUnderlineBtn
        '
        Me.BookChapterUnderlineBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.BookChapterUnderlineBtn.Image = Global.BibleGetIO.My.Resources.Resources.underline
        Me.BookChapterUnderlineBtn.Location = New System.Drawing.Point(450, 16)
        Me.BookChapterUnderlineBtn.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.BookChapterUnderlineBtn.Name = "BookChapterUnderlineBtn"
        Me.BookChapterUnderlineBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterUnderlineBtn.TabIndex = 7
        Me.BookChapterUnderlineBtn.UseVisualStyleBackColor = True
        '
        'BookChapterItalicBtn
        '
        Me.BookChapterItalicBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.BookChapterItalicBtn.Image = Global.BibleGetIO.My.Resources.Resources.italic
        Me.BookChapterItalicBtn.Location = New System.Drawing.Point(403, 16)
        Me.BookChapterItalicBtn.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.BookChapterItalicBtn.Name = "BookChapterItalicBtn"
        Me.BookChapterItalicBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterItalicBtn.TabIndex = 6
        Me.BookChapterItalicBtn.UseVisualStyleBackColor = True
        '
        'BookChapterBoldBtn
        '
        Me.BookChapterBoldBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.BookChapterBoldBtn.Checked = True
        Me.BookChapterBoldBtn.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BookChapterBoldBtn.Image = Global.BibleGetIO.My.Resources.Resources.bold
        Me.BookChapterBoldBtn.Location = New System.Drawing.Point(356, 16)
        Me.BookChapterBoldBtn.Margin = New System.Windows.Forms.Padding(6, 3, 0, 3)
        Me.BookChapterBoldBtn.Name = "BookChapterBoldBtn"
        Me.BookChapterBoldBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterBoldBtn.TabIndex = 5
        Me.BookChapterBoldBtn.UseVisualStyleBackColor = True
        '
        'BookChapterBGColorBtn
        '
        Me.BookChapterBGColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.background_color
        Me.BookChapterBGColorBtn.Location = New System.Drawing.Point(299, 16)
        Me.BookChapterBGColorBtn.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.BookChapterBGColorBtn.Name = "BookChapterBGColorBtn"
        Me.BookChapterBGColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterBGColorBtn.TabIndex = 3
        Me.BookChapterBGColorBtn.UseVisualStyleBackColor = True
        '
        'BookChapterColorBtn
        '
        Me.BookChapterColorBtn.AutoSize = True
        Me.BookChapterColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.text_color
        Me.BookChapterColorBtn.Location = New System.Drawing.Point(252, 16)
        Me.BookChapterColorBtn.Margin = New System.Windows.Forms.Padding(3, 3, 0, 3)
        Me.BookChapterColorBtn.Name = "BookChapterColorBtn"
        Me.BookChapterColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.BookChapterColorBtn.TabIndex = 0
        Me.BookChapterColorBtn.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.VerseNumberFontBtnn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberSubscriptBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberSuperscriptBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberStyleLbl)
        Me.GroupBox2.Controls.Add(Me.VerseNumberUnderlineBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberItalicBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberBoldBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberBGColorBtn)
        Me.GroupBox2.Controls.Add(Me.VerseNumberColorBtn)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 207)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(633, 75)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Verse Number Formatting"
        '
        'VerseNumberFontBtnn
        '
        Me.VerseNumberFontBtnn.Location = New System.Drawing.Point(7, 16)
        Me.VerseNumberFontBtnn.Name = "VerseNumberFontBtnn"
        Me.VerseNumberFontBtnn.Size = New System.Drawing.Size(239, 35)
        Me.VerseNumberFontBtnn.TabIndex = 11
        Me.VerseNumberFontBtnn.Text = "Button3"
        Me.VerseNumberFontBtnn.UseVisualStyleBackColor = True
        '
        'VerseNumberSubscriptBtn
        '
        Me.VerseNumberSubscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseNumberSubscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.superscript
        Me.VerseNumberSubscriptBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.VerseNumberSubscriptBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberSubscriptBtn.Location = New System.Drawing.Point(554, 16)
        Me.VerseNumberSubscriptBtn.Name = "VerseNumberSubscriptBtn"
        Me.VerseNumberSubscriptBtn.Padding = New System.Windows.Forms.Padding(0, 2, 0, 0)
        Me.VerseNumberSubscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberSubscriptBtn.TabIndex = 10
        Me.VerseNumberSubscriptBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberSuperscriptBtn
        '
        Me.VerseNumberSuperscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseNumberSuperscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.subscript
        Me.VerseNumberSuperscriptBtn.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.VerseNumberSuperscriptBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberSuperscriptBtn.Location = New System.Drawing.Point(507, 16)
        Me.VerseNumberSuperscriptBtn.Name = "VerseNumberSuperscriptBtn"
        Me.VerseNumberSuperscriptBtn.Padding = New System.Windows.Forms.Padding(0, 0, 0, 1)
        Me.VerseNumberSuperscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberSuperscriptBtn.TabIndex = 9
        Me.VerseNumberSuperscriptBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberStyleLbl
        '
        Me.VerseNumberStyleLbl.AutoSize = True
        Me.VerseNumberStyleLbl.BackColor = System.Drawing.SystemColors.Control
        Me.VerseNumberStyleLbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.VerseNumberStyleLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VerseNumberStyleLbl.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberStyleLbl.Location = New System.Drawing.Point(7, 54)
        Me.VerseNumberStyleLbl.Name = "VerseNumberStyleLbl"
        Me.VerseNumberStyleLbl.Size = New System.Drawing.Size(72, 15)
        Me.VerseNumberStyleLbl.TabIndex = 8
        Me.VerseNumberStyleLbl.Text = "12pt   Normal"
        '
        'VerseNumberUnderlineBtn
        '
        Me.VerseNumberUnderlineBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseNumberUnderlineBtn.Image = Global.BibleGetIO.My.Resources.Resources.underline
        Me.VerseNumberUnderlineBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberUnderlineBtn.Location = New System.Drawing.Point(450, 16)
        Me.VerseNumberUnderlineBtn.Name = "VerseNumberUnderlineBtn"
        Me.VerseNumberUnderlineBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberUnderlineBtn.TabIndex = 7
        Me.VerseNumberUnderlineBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberItalicBtn
        '
        Me.VerseNumberItalicBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseNumberItalicBtn.Image = Global.BibleGetIO.My.Resources.Resources.italic
        Me.VerseNumberItalicBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberItalicBtn.Location = New System.Drawing.Point(403, 16)
        Me.VerseNumberItalicBtn.Name = "VerseNumberItalicBtn"
        Me.VerseNumberItalicBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberItalicBtn.TabIndex = 6
        Me.VerseNumberItalicBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberBoldBtn
        '
        Me.VerseNumberBoldBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseNumberBoldBtn.Image = Global.BibleGetIO.My.Resources.Resources.bold
        Me.VerseNumberBoldBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberBoldBtn.Location = New System.Drawing.Point(356, 16)
        Me.VerseNumberBoldBtn.Margin = New System.Windows.Forms.Padding(6, 3, 0, 3)
        Me.VerseNumberBoldBtn.Name = "VerseNumberBoldBtn"
        Me.VerseNumberBoldBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberBoldBtn.TabIndex = 5
        Me.VerseNumberBoldBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberBGColorBtn
        '
        Me.VerseNumberBGColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.background_color
        Me.VerseNumberBGColorBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberBGColorBtn.Location = New System.Drawing.Point(299, 16)
        Me.VerseNumberBGColorBtn.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.VerseNumberBGColorBtn.Name = "VerseNumberBGColorBtn"
        Me.VerseNumberBGColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberBGColorBtn.TabIndex = 3
        Me.VerseNumberBGColorBtn.UseVisualStyleBackColor = True
        '
        'VerseNumberColorBtn
        '
        Me.VerseNumberColorBtn.AutoSize = True
        Me.VerseNumberColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.text_color
        Me.VerseNumberColorBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseNumberColorBtn.Location = New System.Drawing.Point(252, 16)
        Me.VerseNumberColorBtn.Margin = New System.Windows.Forms.Padding(3, 3, 0, 3)
        Me.VerseNumberColorBtn.Name = "VerseNumberColorBtn"
        Me.VerseNumberColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseNumberColorBtn.TabIndex = 0
        Me.VerseNumberColorBtn.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.VerseTextFontBtnn)
        Me.GroupBox3.Controls.Add(Me.VerseTextSubscriptBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextSuperscriptBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextStyleLbl)
        Me.GroupBox3.Controls.Add(Me.VerseTextUnderlineBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextItalicBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextBoldBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextBGColorBtn)
        Me.GroupBox3.Controls.Add(Me.VerseTextColorBtn)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 288)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(633, 75)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Verse Text Formatting"
        '
        'VerseTextFontBtnn
        '
        Me.VerseTextFontBtnn.Location = New System.Drawing.Point(7, 16)
        Me.VerseTextFontBtnn.Name = "VerseTextFontBtnn"
        Me.VerseTextFontBtnn.Size = New System.Drawing.Size(239, 35)
        Me.VerseTextFontBtnn.TabIndex = 11
        Me.VerseTextFontBtnn.Text = "Button3"
        Me.VerseTextFontBtnn.UseVisualStyleBackColor = True
        '
        'VerseTextSubscriptBtn
        '
        Me.VerseTextSubscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseTextSubscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.superscript
        Me.VerseTextSubscriptBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.VerseTextSubscriptBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextSubscriptBtn.Location = New System.Drawing.Point(554, 16)
        Me.VerseTextSubscriptBtn.Name = "VerseTextSubscriptBtn"
        Me.VerseTextSubscriptBtn.Padding = New System.Windows.Forms.Padding(0, 2, 0, 0)
        Me.VerseTextSubscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextSubscriptBtn.TabIndex = 10
        Me.VerseTextSubscriptBtn.UseVisualStyleBackColor = True
        '
        'VerseTextSuperscriptBtn
        '
        Me.VerseTextSuperscriptBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseTextSuperscriptBtn.Image = Global.BibleGetIO.My.Resources.Resources.subscript
        Me.VerseTextSuperscriptBtn.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.VerseTextSuperscriptBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextSuperscriptBtn.Location = New System.Drawing.Point(507, 16)
        Me.VerseTextSuperscriptBtn.Name = "VerseTextSuperscriptBtn"
        Me.VerseTextSuperscriptBtn.Padding = New System.Windows.Forms.Padding(0, 0, 0, 1)
        Me.VerseTextSuperscriptBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextSuperscriptBtn.TabIndex = 9
        Me.VerseTextSuperscriptBtn.UseVisualStyleBackColor = True
        '
        'VerseTextStyleLbl
        '
        Me.VerseTextStyleLbl.AutoSize = True
        Me.VerseTextStyleLbl.BackColor = System.Drawing.SystemColors.Control
        Me.VerseTextStyleLbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.VerseTextStyleLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VerseTextStyleLbl.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextStyleLbl.Location = New System.Drawing.Point(7, 54)
        Me.VerseTextStyleLbl.Name = "VerseTextStyleLbl"
        Me.VerseTextStyleLbl.Size = New System.Drawing.Size(72, 15)
        Me.VerseTextStyleLbl.TabIndex = 8
        Me.VerseTextStyleLbl.Text = "12pt   Normal"
        '
        'VerseTextUnderlineBtn
        '
        Me.VerseTextUnderlineBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseTextUnderlineBtn.Image = Global.BibleGetIO.My.Resources.Resources.underline
        Me.VerseTextUnderlineBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextUnderlineBtn.Location = New System.Drawing.Point(450, 16)
        Me.VerseTextUnderlineBtn.Name = "VerseTextUnderlineBtn"
        Me.VerseTextUnderlineBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextUnderlineBtn.TabIndex = 7
        Me.VerseTextUnderlineBtn.UseVisualStyleBackColor = True
        '
        'VerseTextItalicBtn
        '
        Me.VerseTextItalicBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseTextItalicBtn.Image = Global.BibleGetIO.My.Resources.Resources.italic
        Me.VerseTextItalicBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextItalicBtn.Location = New System.Drawing.Point(403, 16)
        Me.VerseTextItalicBtn.Name = "VerseTextItalicBtn"
        Me.VerseTextItalicBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextItalicBtn.TabIndex = 6
        Me.VerseTextItalicBtn.UseVisualStyleBackColor = True
        '
        'VerseTextBoldBtn
        '
        Me.VerseTextBoldBtn.Appearance = System.Windows.Forms.Appearance.Button
        Me.VerseTextBoldBtn.Image = Global.BibleGetIO.My.Resources.Resources.bold
        Me.VerseTextBoldBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextBoldBtn.Location = New System.Drawing.Point(356, 16)
        Me.VerseTextBoldBtn.Name = "VerseTextBoldBtn"
        Me.VerseTextBoldBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextBoldBtn.TabIndex = 5
        Me.VerseTextBoldBtn.UseVisualStyleBackColor = True
        '
        'VerseTextBGColorBtn
        '
        Me.VerseTextBGColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.background_color
        Me.VerseTextBGColorBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextBGColorBtn.Location = New System.Drawing.Point(299, 16)
        Me.VerseTextBGColorBtn.Name = "VerseTextBGColorBtn"
        Me.VerseTextBGColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextBGColorBtn.TabIndex = 3
        Me.VerseTextBGColorBtn.UseVisualStyleBackColor = True
        '
        'VerseTextColorBtn
        '
        Me.VerseTextColorBtn.AutoSize = True
        Me.VerseTextColorBtn.Image = Global.BibleGetIO.My.Resources.Resources.text_color
        Me.VerseTextColorBtn.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.VerseTextColorBtn.Location = New System.Drawing.Point(252, 16)
        Me.VerseTextColorBtn.Name = "VerseTextColorBtn"
        Me.VerseTextColorBtn.Size = New System.Drawing.Size(48, 48)
        Me.VerseTextColorBtn.TabIndex = 0
        Me.VerseTextColorBtn.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.WebBrowser1)
        Me.GroupBox4.Location = New System.Drawing.Point(13, 369)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(633, 187)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Preview"
        '
        'WebBrowser1
        '
        Me.WebBrowser1.AllowNavigation = False
        Me.WebBrowser1.AllowWebBrowserDrop = False
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.IsWebBrowserContextMenuEnabled = False
        Me.WebBrowser1.Location = New System.Drawing.Point(3, 16)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.ScrollBarsEnabled = False
        Me.WebBrowser1.Size = New System.Drawing.Size(627, 168)
        Me.WebBrowser1.TabIndex = 0
        Me.WebBrowser1.WebBrowserShortcutsEnabled = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.GroupBox9)
        Me.GroupBox5.Controls.Add(Me.GroupBox8)
        Me.GroupBox5.Controls.Add(Me.GroupBox7)
        Me.GroupBox5.Controls.Add(Me.GroupBox6)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(634, 108)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Paragraph Formatting"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.Label1)
        Me.GroupBox9.Controls.Add(Me.CheckBox1)
        Me.GroupBox9.Location = New System.Drawing.Point(438, 20)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(183, 82)
        Me.GroupBox9.TabIndex = 3
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Override Bible Version Formatting"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Help
        Me.Label1.Image = Global.BibleGetIO.My.Resources.Resources.help_large
        Me.Label1.Location = New System.Drawing.Point(134, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 28)
        Me.Label1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Label1, resources.GetString("Label1.ToolTip"))
        '
        'CheckBox1
        '
        Me.CheckBox1.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBox1.Font = New System.Drawing.Font("Gabriola", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.ForeColor = System.Drawing.Color.DarkRed
        Me.CheckBox1.Location = New System.Drawing.Point(6, 29)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(61, 44)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "OFF"
        Me.CheckBox1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.ComboBox1)
        Me.GroupBox8.Location = New System.Drawing.Point(328, 20)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(104, 82)
        Me.GroupBox8.TabIndex = 2
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Line-spacing"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"single", "1½", "double"})
        Me.ComboBox1.Location = New System.Drawing.Point(7, 20)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(91, 32)
        Me.ComboBox1.TabIndex = 0
        Me.ComboBox1.Text = "1½"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Button2)
        Me.GroupBox7.Controls.Add(Me.Button1)
        Me.GroupBox7.Location = New System.Drawing.Point(214, 20)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(107, 82)
        Me.GroupBox7.TabIndex = 1
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Indentation"
        '
        'Button2
        '
        Me.Button2.Image = Global.BibleGetIO.My.Resources.Resources.decrease_indent
        Me.Button2.Location = New System.Drawing.Point(53, 19)
        Me.Button2.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(48, 48)
        Me.Button2.TabIndex = 1
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.BibleGetIO.My.Resources.Resources.increase_indent
        Me.Button1.Location = New System.Drawing.Point(6, 19)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 3, 0, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(48, 48)
        Me.Button1.TabIndex = 0
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.RadioButton4)
        Me.GroupBox6.Controls.Add(Me.RadioButton3)
        Me.GroupBox6.Controls.Add(Me.RadioButton2)
        Me.GroupBox6.Controls.Add(Me.RadioButton1)
        Me.GroupBox6.Location = New System.Drawing.Point(7, 20)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(201, 82)
        Me.GroupBox6.TabIndex = 0
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Alignment"
        '
        'RadioButton4
        '
        Me.RadioButton4.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton4.Image = Global.BibleGetIO.My.Resources.Resources.align_justify
        Me.RadioButton4.Location = New System.Drawing.Point(148, 20)
        Me.RadioButton4.Margin = New System.Windows.Forms.Padding(0)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(48, 48)
        Me.RadioButton4.TabIndex = 3
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton3.Image = Global.BibleGetIO.My.Resources.Resources.align_right
        Me.RadioButton3.Location = New System.Drawing.Point(101, 20)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(0)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(48, 48)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton2.Image = Global.BibleGetIO.My.Resources.Resources.align_center
        Me.RadioButton2.Location = New System.Drawing.Point(54, 20)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(0)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(48, 48)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton1.Image = Global.BibleGetIO.My.Resources.Resources.align_left
        Me.RadioButton1.Location = New System.Drawing.Point(7, 20)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(0)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(48, 48)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 10000
        Me.ToolTip1.InitialDelay = 500
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 100
        Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.ToolTip1.ToolTipTitle = "More information"
        '
        'Preferences
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 568)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Preferences"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Preferences"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BookChapterColorBtn As System.Windows.Forms.Button
    Friend WithEvents FontDlg As System.Windows.Forms.FontDialog
    Friend WithEvents ColorDlg As System.Windows.Forms.ColorDialog
    Friend WithEvents BookChapterBGColorBtn As System.Windows.Forms.Button
    Friend WithEvents BookChapterUnderlineBtn As System.Windows.Forms.CheckBox
    Friend WithEvents BookChapterItalicBtn As System.Windows.Forms.CheckBox
    Friend WithEvents BookChapterBoldBtn As System.Windows.Forms.CheckBox
    Friend WithEvents BookChapterStyleLbl As System.Windows.Forms.Label
    Friend WithEvents BookChapterSubscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents BookChapterSuperscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents VerseNumberSubscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseNumberSuperscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseNumberStyleLbl As System.Windows.Forms.Label
    Friend WithEvents VerseNumberUnderlineBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseNumberItalicBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseNumberBoldBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseNumberBGColorBtn As System.Windows.Forms.Button
    Friend WithEvents VerseNumberColorBtn As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents VerseTextSubscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseTextSuperscriptBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseTextStyleLbl As System.Windows.Forms.Label
    Friend WithEvents VerseTextUnderlineBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseTextItalicBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseTextBoldBtn As System.Windows.Forms.CheckBox
    Friend WithEvents VerseTextBGColorBtn As System.Windows.Forms.Button
    Friend WithEvents VerseTextColorBtn As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BookChapterFontBtnn As System.Windows.Forms.Button
    Friend WithEvents VerseNumberFontBtnn As System.Windows.Forms.Button
    Friend WithEvents VerseTextFontBtnn As System.Windows.Forms.Button
End Class
