<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectSingleBibleVersion
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectSingleBibleVersion))
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SetBibleVersionBtn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(12, 42)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(516, 191)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(227, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select Bible version to search from"
        '
        'SetBibleVersionBtn
        '
        Me.SetBibleVersionBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.SetBibleVersionBtn.Location = New System.Drawing.Point(430, 259)
        Me.SetBibleVersionBtn.Name = "SetBibleVersionBtn"
        Me.SetBibleVersionBtn.Size = New System.Drawing.Size(75, 23)
        Me.SetBibleVersionBtn.TabIndex = 2
        Me.SetBibleVersionBtn.Text = "SET"
        Me.SetBibleVersionBtn.UseVisualStyleBackColor = True
        '
        'SelectSingleBibleVersion
        '
        Me.AcceptButton = Me.SetBibleVersionBtn
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 309)
        Me.Controls.Add(Me.SetBibleVersionBtn)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SelectSingleBibleVersion"
        Me.Text = "Select Bible version"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListView1 As Windows.Forms.ListView
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents SetBibleVersionBtn As Windows.Forms.Button
End Class
