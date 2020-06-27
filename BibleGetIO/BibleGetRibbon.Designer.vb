Partial Class BibleGetRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BibleGetTab = Me.Factory.CreateRibbonTab
        Me.BibleGetTabGroup1 = Me.Factory.CreateRibbonGroup
        Me.InsertBibleQuoteFromDialogBtn = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.InsertBibleQuoteFromTextSelectionBtn = Me.Factory.CreateRibbonButton
        Me.BibleGetTabGroup2 = Me.Factory.CreateRibbonGroup
        Me.SearchBtn = Me.Factory.CreateRibbonButton
        Me.BibleGetTabGroup3 = Me.Factory.CreateRibbonGroup
        Me.PreferencesBtn = Me.Factory.CreateRibbonButton
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.HelpBtn = Me.Factory.CreateRibbonButton
        Me.BibleGetTabGroup4 = Me.Factory.CreateRibbonGroup
        Me.SendFeedbackBtn = Me.Factory.CreateRibbonButton
        Me.MakeContributionBtn = Me.Factory.CreateRibbonButton
        Me.AboutBtn = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.StatusBtn = Me.Factory.CreateRibbonButton
        Me.BibleGetTab.SuspendLayout()
        Me.BibleGetTabGroup1.SuspendLayout()
        Me.BibleGetTabGroup2.SuspendLayout()
        Me.BibleGetTabGroup3.SuspendLayout()
        Me.BibleGetTabGroup4.SuspendLayout()
        Me.SuspendLayout()
        '
        'BibleGetTab
        '
        Me.BibleGetTab.Groups.Add(Me.BibleGetTabGroup1)
        Me.BibleGetTab.Groups.Add(Me.BibleGetTabGroup2)
        Me.BibleGetTab.Groups.Add(Me.BibleGetTabGroup3)
        Me.BibleGetTab.Groups.Add(Me.BibleGetTabGroup4)
        Me.BibleGetTab.KeyTip = "Q"
        Me.BibleGetTab.Label = "BIBLEGET I/O"
        Me.BibleGetTab.Name = "BibleGetTab"
        Me.BibleGetTab.Position = Me.Factory.RibbonPosition.AfterOfficeId("TabAddIn")
        '
        'BibleGetTabGroup1
        '
        Me.BibleGetTabGroup1.Items.Add(Me.InsertBibleQuoteFromDialogBtn)
        Me.BibleGetTabGroup1.Items.Add(Me.Separator2)
        Me.BibleGetTabGroup1.Items.Add(Me.InsertBibleQuoteFromTextSelectionBtn)
        Me.BibleGetTabGroup1.Label = "Insert Bible Quote"
        Me.BibleGetTabGroup1.Name = "BibleGetTabGroup1"
        '
        'InsertBibleQuoteFromDialogBtn
        '
        Me.InsertBibleQuoteFromDialogBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.InsertBibleQuoteFromDialogBtn.Image = Global.BibleGetIO.My.Resources.Resources.quotefrominput_small
        Me.InsertBibleQuoteFromDialogBtn.KeyTip = "B"
        Me.InsertBibleQuoteFromDialogBtn.Label = "Insert Bible Quote from Dialog"
        Me.InsertBibleQuoteFromDialogBtn.Name = "InsertBibleQuoteFromDialogBtn"
        Me.InsertBibleQuoteFromDialogBtn.ScreenTip = "Get your desired Bible quote using the dialog and choosing preferred Bible Versio" &
    "ns"
        Me.InsertBibleQuoteFromDialogBtn.ShowImage = True
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'InsertBibleQuoteFromTextSelectionBtn
        '
        Me.InsertBibleQuoteFromTextSelectionBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.InsertBibleQuoteFromTextSelectionBtn.Image = Global.BibleGetIO.My.Resources.Resources.quotefromselection_large
        Me.InsertBibleQuoteFromTextSelectionBtn.KeyTip = "T"
        Me.InsertBibleQuoteFromTextSelectionBtn.Label = "Insert bible quote from Text selection"
        Me.InsertBibleQuoteFromTextSelectionBtn.Name = "InsertBibleQuoteFromTextSelectionBtn"
        Me.InsertBibleQuoteFromTextSelectionBtn.ScreenTip = "Select your desired Bible quote directly in your document"
        Me.InsertBibleQuoteFromTextSelectionBtn.ShowImage = True
        '
        'BibleGetTabGroup2
        '
        Me.BibleGetTabGroup2.Items.Add(Me.SearchBtn)
        Me.BibleGetTabGroup2.Label = "Search"
        Me.BibleGetTabGroup2.Name = "BibleGetTabGroup2"
        '
        'SearchBtn
        '
        Me.SearchBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SearchBtn.Image = Global.BibleGetIO.My.Resources.Resources.search_32x32
        Me.SearchBtn.KeyTip = "S"
        Me.SearchBtn.Label = "Search for verses by keyword"
        Me.SearchBtn.Name = "SearchBtn"
        Me.SearchBtn.ShowImage = True
        '
        'BibleGetTabGroup3
        '
        Me.BibleGetTabGroup3.Items.Add(Me.PreferencesBtn)
        Me.BibleGetTabGroup3.Items.Add(Me.Separator3)
        Me.BibleGetTabGroup3.Items.Add(Me.HelpBtn)
        Me.BibleGetTabGroup3.Label = "Settings"
        Me.BibleGetTabGroup3.Name = "BibleGetTabGroup3"
        '
        'PreferencesBtn
        '
        Me.PreferencesBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PreferencesBtn.Image = Global.BibleGetIO.My.Resources.Resources.preferences_large
        Me.PreferencesBtn.KeyTip = "P"
        Me.PreferencesBtn.Label = "Preferences"
        Me.PreferencesBtn.Name = "PreferencesBtn"
        Me.PreferencesBtn.ScreenTip = "Choose your preferred text and paragraph formatting"
        Me.PreferencesBtn.ShowImage = True
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'HelpBtn
        '
        Me.HelpBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.HelpBtn.Image = Global.BibleGetIO.My.Resources.Resources.help_large
        Me.HelpBtn.KeyTip = "H"
        Me.HelpBtn.Label = "Help"
        Me.HelpBtn.Name = "HelpBtn"
        Me.HelpBtn.ScreenTip = "Instructions for the usage of this AddIn"
        Me.HelpBtn.ShowImage = True
        '
        'BibleGetTabGroup4
        '
        Me.BibleGetTabGroup4.Items.Add(Me.SendFeedbackBtn)
        Me.BibleGetTabGroup4.Items.Add(Me.MakeContributionBtn)
        Me.BibleGetTabGroup4.Items.Add(Me.AboutBtn)
        Me.BibleGetTabGroup4.Items.Add(Me.Separator1)
        Me.BibleGetTabGroup4.Items.Add(Me.StatusBtn)
        Me.BibleGetTabGroup4.Label = "About"
        Me.BibleGetTabGroup4.Name = "BibleGetTabGroup4"
        '
        'SendFeedbackBtn
        '
        Me.SendFeedbackBtn.Image = Global.BibleGetIO.My.Resources.Resources.email_smallB
        Me.SendFeedbackBtn.KeyTip = "F"
        Me.SendFeedbackBtn.Label = "Send feedback"
        Me.SendFeedbackBtn.Name = "SendFeedbackBtn"
        Me.SendFeedbackBtn.ScreenTip = "Any bug reports or suggestions for enhancement can be submitted here"
        Me.SendFeedbackBtn.ShowImage = True
        '
        'MakeContributionBtn
        '
        Me.MakeContributionBtn.Image = Global.BibleGetIO.My.Resources.Resources.paypal_small
        Me.MakeContributionBtn.KeyTip = "G"
        Me.MakeContributionBtn.Label = "Make a contribution"
        Me.MakeContributionBtn.Name = "MakeContributionBtn"
        Me.MakeContributionBtn.ScreenTip = "Help support the BibleGet I/O Project"
        Me.MakeContributionBtn.ShowImage = True
        '
        'AboutBtn
        '
        Me.AboutBtn.Image = Global.BibleGetIO.My.Resources.Resources.info_small
        Me.AboutBtn.KeyTip = "V"
        Me.AboutBtn.Label = "About"
        Me.AboutBtn.Name = "AboutBtn"
        Me.AboutBtn.ScreenTip = "Information about this plugin."
        Me.AboutBtn.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'StatusBtn
        '
        Me.StatusBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.StatusBtn.Image = Global.BibleGetIO.My.Resources.Resources.red_x_wrong_mark
        Me.StatusBtn.Label = "STATUS: NOT READY"
        Me.StatusBtn.Name = "StatusBtn"
        Me.StatusBtn.ShowImage = True
        '
        'BibleGetRibbon
        '
        Me.Name = "BibleGetRibbon"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.BibleGetTab)
        Me.BibleGetTab.ResumeLayout(False)
        Me.BibleGetTab.PerformLayout()
        Me.BibleGetTabGroup1.ResumeLayout(False)
        Me.BibleGetTabGroup1.PerformLayout()
        Me.BibleGetTabGroup2.ResumeLayout(False)
        Me.BibleGetTabGroup2.PerformLayout()
        Me.BibleGetTabGroup3.ResumeLayout(False)
        Me.BibleGetTabGroup3.PerformLayout()
        Me.BibleGetTabGroup4.ResumeLayout(False)
        Me.BibleGetTabGroup4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BibleGetTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents BibleGetTabGroup1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InsertBibleQuoteFromDialogBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents InsertBibleQuoteFromTextSelectionBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents PreferencesBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents HelpBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents SendFeedbackBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MakeContributionBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AboutBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents StatusBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BibleGetTabGroup3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BibleGetTabGroup4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BibleGetTabGroup2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SearchBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property BibleGetRibbon() As BibleGetRibbon
        Get
            Return Me.GetRibbon(Of BibleGetRibbon)()
        End Get
    End Property
End Class
