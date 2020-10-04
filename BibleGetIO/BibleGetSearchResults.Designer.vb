<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BibleGetSearchResults
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
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BibleGetSearchResults))
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.TermToSearch = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.BibleVersionForSearch = New System.Windows.Forms.ListView()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.FilterForTerm = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ExactMatchChkBox = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'TermToSearch
        '
        Me.TermToSearch.Location = New System.Drawing.Point(19, 297)
        Me.TermToSearch.Name = "TermToSearch"
        Me.TermToSearch.Size = New System.Drawing.Size(215, 22)
        Me.TermToSearch.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 277)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Term to search"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(184, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Bible version to search from"
        '
        'Label3
        '
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 625)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4)
        Me.Label3.Name = "Label3"
        Me.Label3.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Label3.Size = New System.Drawing.Size(279, 28)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "WAITING TO START REQUEST..."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 602)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "JOB STATUS:"
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(13, 661)
        Me.ProgressBar2.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(280, 21)
        Me.ProgressBar2.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar2.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Label5.Location = New System.Drawing.Point(537, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(118, 20)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Search results"
        '
        'BibleVersionForSearch
        '
        Me.BibleVersionForSearch.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.BibleVersionForSearch.HideSelection = False
        Me.BibleVersionForSearch.Location = New System.Drawing.Point(19, 64)
        Me.BibleVersionForSearch.MultiSelect = False
        Me.BibleVersionForSearch.Name = "BibleVersionForSearch"
        Me.BibleVersionForSearch.Size = New System.Drawing.Size(516, 191)
        Me.BibleVersionForSearch.TabIndex = 12
        Me.BibleVersionForSearch.UseCompatibleStateImageBehavior = False
        '
        'WebBrowser1
        '
        Me.WebBrowser1.AllowNavigation = False
        Me.WebBrowser1.AllowWebBrowserDrop = False
        Me.WebBrowser1.IsWebBrowserContextMenuEnabled = False
        Me.WebBrowser1.Location = New System.Drawing.Point(541, 64)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(1136, 803)
        Me.WebBrowser1.TabIndex = 13
        Me.WebBrowser1.WebBrowserShortcutsEnabled = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.InfoText
        Me.TextBox1.ForeColor = System.Drawing.Color.LimeGreen
        Me.TextBox1.Location = New System.Drawing.Point(12, 689)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(515, 178)
        Me.TextBox1.TabIndex = 14
        Me.TextBox1.Text = "Ready to send request for search term to the BibleGet service endpoint."
        '
        'FilterForTerm
        '
        Me.FilterForTerm.Location = New System.Drawing.Point(20, 417)
        Me.FilterForTerm.Name = "FilterForTerm"
        Me.FilterForTerm.Size = New System.Drawing.Size(214, 22)
        Me.FilterForTerm.TabIndex = 15
        Me.FilterForTerm.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 397)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(198, 17)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Filter results with another term"
        Me.Label6.Visible = False
        '
        'Button3
        '
        Me.Button3.Image = Global.BibleGetIO.My.Resources.Resources.Sort_16
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.Location = New System.Drawing.Point(352, 386)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(110, 62)
        Me.Button3.TabIndex = 18
        Me.Button3.Text = "Order by Reference"
        Me.Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'Button2
        '
        Me.Button2.Image = Global.BibleGetIO.My.Resources.Resources.filter
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button2.Location = New System.Drawing.Point(240, 386)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(106, 62)
        Me.Button2.TabIndex = 17
        Me.Button2.Text = "Apply filter"
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Image = Global.BibleGetIO.My.Resources.Resources.search_small
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(429, 257)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(106, 62)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Search"
        Me.Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ExactMatchChkBox
        '
        Me.ExactMatchChkBox.AutoSize = True
        Me.ExactMatchChkBox.Location = New System.Drawing.Point(240, 297)
        Me.ExactMatchChkBox.Name = "ExactMatchChkBox"
        Me.ExactMatchChkBox.Size = New System.Drawing.Size(153, 21)
        Me.ExactMatchChkBox.TabIndex = 19
        Me.ExactMatchChkBox.Text = "Only exact matches"
        Me.ToolTip1.SetToolTip(Me.ExactMatchChkBox, resources.GetString("ExactMatchChkBox.ToolTip"))
        Me.ExactMatchChkBox.UseVisualStyleBackColor = True
        '
        'ToolTip1
        '
        Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.ToolTip1.ToolTipTitle = "More info"
        '
        'BibleGetSearchResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1689, 879)
        Me.Controls.Add(Me.ExactMatchChkBox)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.FilterForTerm)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.BibleVersionForSearch)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ProgressBar2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TermToSearch)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "BibleGetSearchResults"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Search for Bible Verses"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
    Friend WithEvents TermToSearch As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents ProgressBar2 As Windows.Forms.ProgressBar
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents BibleVersionForSearch As Windows.Forms.ListView
    Friend WithEvents WebBrowser1 As Windows.Forms.WebBrowser
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents FilterForTerm As Windows.Forms.TextBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents ExactMatchChkBox As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
End Class
