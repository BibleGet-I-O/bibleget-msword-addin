Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class BibleGetHelp

    Private HtmlStr0 As String

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    'turn off the annoying clicking sound when the preview window refreshes (WebBrowser control)
    Const DS As Integer = 21
    Const SP As Integer = &H2
    <DllImport("urlmon.dll")> _
    <PreserveSig> _
    Private Shared Function CoInternetSetFeatureEnabled(FeatureEntry As Integer, <MarshalAs(UnmanagedType.U4)> dSFlags As Integer, eEnable As Boolean) As <MarshalAs(UnmanagedType.[Error])> Integer
    End Function

    'TODO: complete BibleGetHelp Class based on NetBeans project BibleGetHelp class

    Private Sub BibleGetHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CoInternetSetFeatureEnabled(DS, SP, True)

        'Diagnostics.Debug.WriteLine(Now.ToShortTimeString + ": BibleGetHelp load event being issued")
        Me.Text = __("Instructions")
        TreeView1.Nodes.Clear()

        Dim rootNode As TreeNode = New TreeNode(__("Help"))
        rootNode.NodeFont = New System.Drawing.Font("Garamond", 14, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
        Dim usageNode As TreeNode = New TreeNode(__("Usage of the Plugin"))
        usageNode.NodeFont = New System.Drawing.Font("Garamond", 12, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
        Dim formulationNode As TreeNode = New TreeNode(__("Formulation of the Queries"))
        formulationNode.NodeFont = New System.Drawing.Font("Garamond", 12, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
        Dim booksNode As TreeNode = New TreeNode(__("Biblical Books and Abbreviations"))
        booksNode.NodeFont = New System.Drawing.Font("Garamond", 12, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)

        TreeView1.Nodes.Add(rootNode)
        rootNode.Nodes.Add(usageNode)
        rootNode.Nodes.Add(formulationNode)
        rootNode.Nodes.Add(booksNode)

        rootNode.Checked = True
        rootNode.Expand()

        'TODO: Populate children of booksNode with language variants
        Me.HtmlStr0 = "<html><head><meta charset=""utf-8""></head><body>"
        Me.HtmlStr0 &= "<h1>" + __("Help for BibleGet (Open Office Writer)") + "</h1>"
        Me.HtmlStr0 &= "<p>" + __("This Help dialog window introduces the user to the usage of the BibleGet I/O plugin for Open Office Writer.") + "</p>"
        Me.HtmlStr0 &= "<p>" + __("The Help is divided into three sections:") + "</p>"
        Me.HtmlStr0 &= "<ul>"
        Me.HtmlStr0 &= "<li>" + __("Usage of the Plugin") + "</li>"
        Me.HtmlStr0 &= "<li>" + __("Formulation of the Queries") + "</li>"
        Me.HtmlStr0 &= "<li>" + __("Biblical Books and Abbreviations") + "</li>"
        Me.HtmlStr0 &= "</ul>"
        Me.HtmlStr0 &= "<p><b>" + __("AUTHOR") + ":</b> " + __("John R. D'Orazio (chaplain at Roma Tre University)") + "</p>"
        Me.HtmlStr0 &= "<p><b>" + __("COLLABORATORS") + ":</b> " + __("Giovanni Gregori (computing) and Simone Urbinati (MUG Roma Tre)") + "</p>"
        Me.HtmlStr0 &= "<p><b>" + __("Version").ToUpper + ":</b> " & My.Application.Info.Version.ToString + "</p>"
        Me.HtmlStr0 &= "<p>© <b>Copyright 2016 BibleGet I/O by John R. D'Orazio</b> <a href=""mailto:john.dorazio@cappellaniauniroma3.org"">john.dorazio@cappellaniauniroma3.org</a></p>"
        Me.HtmlStr0 &= "<p><b>" + __("PROJECT WEBSITE") + ": </b><a href=""http://www.bibleget.io"">http://www.bibleget.io</a> | <b>" + __("EMAIL ADDRESS FOR INFORMATION OR FEEDBACK ON THE PROJECT") + ":</b> <a href=""mailto:bibleget.io@gmail.com"">bibleget.io@gmail.com</a></p>"
        Me.HtmlStr0 &= "<p>Cappellania Università degli Studi Roma Tre - Piazzale San Paolo 1/E - 00120 Città del Vaticano - +39 06.69.88.08.09 - <a href=""mailto:cappellania.uniroma3@gmail.com"">cappellania.uniroma3@gmail.com</a></p></body></html>"

        SetPreviewDocument(__("Help"))

    End Sub



    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        SetPreviewDocument(e.Node.Text)
    End Sub

    Private Sub SetPreviewDocument(ByVal node As String)
        Dim previewDocument As String = String.Empty
        Select Case node
            Case __("Help")
                previewDocument = Me.HtmlStr0
            Case __("Usage of the Plugin")
            Case __("Formulation of the Queries")
            Case __("Biblical Books and Abbreviations")
                'WebBrowser1.
        End Select

        If WebBrowser1.Document Is Nothing Then
            WebBrowser1.DocumentText = previewDocument
        Else
            WebBrowser1.Document.Write(String.Empty)
            WebBrowser1.Document.Write(previewDocument)
            WebBrowser1.Refresh()
        End If

    End Sub
End Class