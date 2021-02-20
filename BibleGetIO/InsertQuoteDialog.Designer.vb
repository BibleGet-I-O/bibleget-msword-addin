<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InsertQuoteDialog
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InsertQuoteDialog))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PreferHebrewOriginLbl = New System.Windows.Forms.Label()
        Me.PreferGreekOriginLbl = New System.Windows.Forms.Label()
        Me.PreferOriginToggle = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(16, 357)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(968, 111)
        Me.TextBox1.TabIndex = 10
        Me.TextBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(705, 329)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(280, 21)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(701, 270)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "SERVER STATUS:"
        '
        'Label2
        '
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(705, 293)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Label2.Size = New System.Drawing.Size(279, 28)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "WAITING FOR REQUEST..."
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(152, 41)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 17)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "UserCode"
        '
        'TextBox2
        '
        Me.TextBox2.ForeColor = System.Drawing.Color.Gray
        Me.TextBox2.Location = New System.Drawing.Point(156, 62)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(516, 22)
        Me.TextBox2.TabIndex = 0
        Me.TextBox2.Text = "e.g. John 3:16;1 John 4,7-8"
        '
        'ListView1
        '
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(156, 129)
        Me.ListView1.Margin = New System.Windows.Forms.Padding(4)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(516, 191)
        Me.ListView1.TabIndex = 1
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(156, 106)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(143, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Choose Bible Version"
        '
        'PreferHebrewOriginLbl
        '
        Me.PreferHebrewOriginLbl.AutoSize = True
        Me.PreferHebrewOriginLbl.Location = New System.Drawing.Point(730, 183)
        Me.PreferHebrewOriginLbl.MaximumSize = New System.Drawing.Size(80, 0)
        Me.PreferHebrewOriginLbl.Name = "PreferHebrewOriginLbl"
        Me.PreferHebrewOriginLbl.Size = New System.Drawing.Size(60, 51)
        Me.PreferHebrewOriginLbl.TabIndex = 12
        Me.PreferHebrewOriginLbl.Text = "Prefer Hebrew origin"
        Me.PreferHebrewOriginLbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PreferGreekOriginLbl
        '
        Me.PreferGreekOriginLbl.AutoSize = True
        Me.PreferGreekOriginLbl.Location = New System.Drawing.Point(888, 183)
        Me.PreferGreekOriginLbl.MaximumSize = New System.Drawing.Size(80, 0)
        Me.PreferGreekOriginLbl.MinimumSize = New System.Drawing.Size(60, 0)
        Me.PreferGreekOriginLbl.Name = "PreferGreekOriginLbl"
        Me.PreferGreekOriginLbl.Size = New System.Drawing.Size(60, 51)
        Me.PreferGreekOriginLbl.TabIndex = 13
        Me.PreferGreekOriginLbl.Text = "Prefer Greek origin"
        Me.PreferGreekOriginLbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PreferOriginToggle
        '
        Me.PreferOriginToggle.Appearance = System.Windows.Forms.Appearance.Button
        Me.PreferOriginToggle.AutoSize = True
        Me.PreferOriginToggle.BackColor = System.Drawing.Color.Transparent
        Me.PreferOriginToggle.FlatAppearance.BorderSize = 0
        Me.PreferOriginToggle.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent
        Me.PreferOriginToggle.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.PreferOriginToggle.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.PreferOriginToggle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.PreferOriginToggle.Image = Global.BibleGetIO.My.Resources.Resources.toggle_button_state_left
        Me.PreferOriginToggle.Location = New System.Drawing.Point(796, 181)
        Me.PreferOriginToggle.Name = "PreferOriginToggle"
        Me.PreferOriginToggle.Size = New System.Drawing.Size(86, 54)
        Me.PreferOriginToggle.TabIndex = 11
        Me.PreferOriginToggle.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Image = Global.BibleGetIO.My.Resources.Resources.arrow_down
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Location = New System.Drawing.Point(715, 62)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Padding = New System.Windows.Forms.Padding(0, 0, 9, 0)
        Me.Button1.Size = New System.Drawing.Size(251, 73)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "GET BIBLE QUOTE"
        Me.Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.Button1.UseVisualStyleBackColor = True
        '
        'InsertQuoteDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1001, 484)
        Me.Controls.Add(Me.PreferGreekOriginLbl)
        Me.Controls.Add(Me.PreferHebrewOriginLbl)
        Me.Controls.Add(Me.PreferOriginToggle)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "InsertQuoteDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "InsertQuoteDialog"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents PreferOriginToggle As Windows.Forms.CheckBox
    Friend WithEvents PreferHebrewOriginLbl As Windows.Forms.Label
    Friend WithEvents PreferGreekOriginLbl As Windows.Forms.Label
End Class
