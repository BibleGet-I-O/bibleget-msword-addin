<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutBibleGet
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

    Friend WithEvents TableLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents LabelProductName As System.Windows.Forms.Label
    Friend WithEvents LabelVersion As System.Windows.Forms.Label
    Friend WithEvents LabelCopyright As System.Windows.Forms.Label

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AboutBibleGet))
        Me.TableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.LabelProductName = New System.Windows.Forms.Label()
        Me.LabelVersion = New System.Windows.Forms.Label()
        Me.LabelCopyright = New System.Windows.Forms.Label()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.CurrentInfo = New System.Windows.Forms.Label()
        Me.ServerData = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ServerDataLangs = New System.Windows.Forms.TextBox()
        Me.ServerDataLangsCount = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel.SuspendLayout()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel
        '
        Me.TableLayoutPanel.ColumnCount = 2
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75.0!))
        Me.TableLayoutPanel.Controls.Add(Me.LogoPictureBox, 0, 0)
        Me.TableLayoutPanel.Controls.Add(Me.LabelProductName, 1, 0)
        Me.TableLayoutPanel.Controls.Add(Me.LabelVersion, 1, 1)
        Me.TableLayoutPanel.Controls.Add(Me.LabelCopyright, 1, 2)
        Me.TableLayoutPanel.Controls.Add(Me.WebBrowser1, 1, 4)
        Me.TableLayoutPanel.Controls.Add(Me.Panel1, 0, 5)
        Me.TableLayoutPanel.Controls.Add(Me.Panel2, 0, 6)
        Me.TableLayoutPanel.Controls.Add(Me.Panel3, 0, 7)
        Me.TableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel.Location = New System.Drawing.Point(12, 11)
        Me.TableLayoutPanel.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel.Name = "TableLayoutPanel"
        Me.TableLayoutPanel.RowCount = 8
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 44.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 92.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.0!))
        Me.TableLayoutPanel.Size = New System.Drawing.Size(819, 726)
        Me.TableLayoutPanel.TabIndex = 0
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LogoPictureBox.Image = Global.BibleGetIO.My.Resources.Resources.holy_bible_x128_B
        Me.LogoPictureBox.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.LogoPictureBox.Location = New System.Drawing.Point(4, 4)
        Me.LogoPictureBox.Margin = New System.Windows.Forms.Padding(4)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.TableLayoutPanel.SetRowSpan(Me.LogoPictureBox, 5)
        Me.LogoPictureBox.Size = New System.Drawing.Size(196, 313)
        Me.LogoPictureBox.TabIndex = 0
        Me.LogoPictureBox.TabStop = False
        '
        'LabelProductName
        '
        Me.LabelProductName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelProductName.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.LabelProductName.Location = New System.Drawing.Point(212, 0)
        Me.LabelProductName.Margin = New System.Windows.Forms.Padding(8, 0, 4, 0)
        Me.LabelProductName.MaximumSize = New System.Drawing.Size(0, 21)
        Me.LabelProductName.Name = "LabelProductName"
        Me.LabelProductName.Size = New System.Drawing.Size(603, 21)
        Me.LabelProductName.TabIndex = 0
        Me.LabelProductName.Text = "BibleGet I/O Plugin for MSWord 2007+"
        Me.LabelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelVersion
        '
        Me.LabelVersion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelVersion.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.LabelVersion.Location = New System.Drawing.Point(212, 25)
        Me.LabelVersion.Margin = New System.Windows.Forms.Padding(8, 0, 4, 0)
        Me.LabelVersion.MaximumSize = New System.Drawing.Size(0, 21)
        Me.LabelVersion.Name = "LabelVersion"
        Me.LabelVersion.Size = New System.Drawing.Size(603, 21)
        Me.LabelVersion.TabIndex = 0
        Me.LabelVersion.Text = "current version"
        Me.LabelVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelCopyright
        '
        Me.LabelCopyright.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelCopyright.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.LabelCopyright.Location = New System.Drawing.Point(212, 50)
        Me.LabelCopyright.Margin = New System.Windows.Forms.Padding(8, 0, 4, 0)
        Me.LabelCopyright.MaximumSize = New System.Drawing.Size(0, 21)
        Me.LabelCopyright.Name = "LabelCopyright"
        Me.LabelCopyright.Size = New System.Drawing.Size(603, 21)
        Me.LabelCopyright.TabIndex = 0
        Me.LabelCopyright.Text = "© 2015 John Romano D'Orazio"
        Me.LabelCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'WebBrowser1
        '
        Me.WebBrowser1.AllowNavigation = False
        Me.WebBrowser1.AllowWebBrowserDrop = False
        Me.WebBrowser1.IsWebBrowserContextMenuEnabled = False
        Me.WebBrowser1.Location = New System.Drawing.Point(208, 104)
        Me.WebBrowser1.Margin = New System.Windows.Forms.Padding(4)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(27, 25)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(607, 213)
        Me.WebBrowser1.TabIndex = 1
        Me.WebBrowser1.WebBrowserShortcutsEnabled = False
        '
        'Panel1
        '
        Me.TableLayoutPanel.SetColumnSpan(Me.Panel1, 2)
        Me.Panel1.Controls.Add(Me.ListView1)
        Me.Panel1.Controls.Add(Me.CurrentInfo)
        Me.Panel1.Controls.Add(Me.ServerData)
        Me.Panel1.Location = New System.Drawing.Point(4, 325)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(811, 270)
        Me.Panel1.TabIndex = 3
        '
        'ListView1
        '
        Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.Location = New System.Drawing.Point(0, 49)
        Me.ListView1.Margin = New System.Windows.Forms.Padding(4)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(810, 221)
        Me.ListView1.TabIndex = 4
        Me.ListView1.TabStop = False
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'CurrentInfo
        '
        Me.CurrentInfo.AutoSize = True
        Me.CurrentInfo.Location = New System.Drawing.Point(5, 30)
        Me.CurrentInfo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.CurrentInfo.Name = "CurrentInfo"
        Me.CurrentInfo.Size = New System.Drawing.Size(71, 17)
        Me.CurrentInfo.TabIndex = 3
        Me.CurrentInfo.Text = "UserCode"
        '
        'ServerData
        '
        Me.ServerData.AutoSize = True
        Me.ServerData.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ServerData.Location = New System.Drawing.Point(4, 0)
        Me.ServerData.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.ServerData.Name = "ServerData"
        Me.ServerData.Size = New System.Drawing.Size(109, 25)
        Me.ServerData.TabIndex = 2
        Me.ServerData.Text = "UserCode"
        '
        'Panel2
        '
        Me.TableLayoutPanel.SetColumnSpan(Me.Panel2, 2)
        Me.Panel2.Controls.Add(Me.ServerDataLangs)
        Me.Panel2.Controls.Add(Me.ServerDataLangsCount)
        Me.Panel2.Location = New System.Drawing.Point(4, 603)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(811, 84)
        Me.Panel2.TabIndex = 5
        '
        'ServerDataLangs
        '
        Me.ServerDataLangs.Location = New System.Drawing.Point(4, 20)
        Me.ServerDataLangs.Margin = New System.Windows.Forms.Padding(4)
        Me.ServerDataLangs.Multiline = True
        Me.ServerDataLangs.Name = "ServerDataLangs"
        Me.ServerDataLangs.ReadOnly = True
        Me.ServerDataLangs.Size = New System.Drawing.Size(801, 61)
        Me.ServerDataLangs.TabIndex = 5
        Me.ServerDataLangs.TabStop = False
        '
        'ServerDataLangsCount
        '
        Me.ServerDataLangsCount.AutoSize = True
        Me.ServerDataLangsCount.Location = New System.Drawing.Point(5, 0)
        Me.ServerDataLangsCount.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.ServerDataLangsCount.Name = "ServerDataLangsCount"
        Me.ServerDataLangsCount.Size = New System.Drawing.Size(156, 17)
        Me.ServerDataLangsCount.TabIndex = 4
        Me.ServerDataLangsCount.Text = "ServerDataLangsCount"
        '
        'Panel3
        '
        Me.TableLayoutPanel.SetColumnSpan(Me.Panel3, 2)
        Me.Panel3.Controls.Add(Me.ProgressBar1)
        Me.Panel3.Controls.Add(Me.OKButton)
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Location = New System.Drawing.Point(4, 695)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(811, 27)
        Me.Panel3.TabIndex = 7
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(391, 0)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(312, 28)
        Me.ProgressBar1.TabIndex = 8
        Me.ProgressBar1.Visible = False
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(711, 0)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 7
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(383, 28)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'AboutBibleGet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(843, 748)
        Me.Controls.Add(Me.TableLayoutPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutBibleGet"
        Me.Padding = New System.Windows.Forms.Padding(12, 11, 12, 11)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About BibleGet I/O"
        Me.TableLayoutPanel.ResumeLayout(False)
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents ServerData As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CurrentInfo As System.Windows.Forms.Label
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ServerDataLangsCount As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ServerDataLangs As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar

End Class
