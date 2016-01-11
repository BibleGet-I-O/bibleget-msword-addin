Imports System.Windows.Forms

Public Class BibleGetHelp
    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    'TODO: complete BibleGetHelp Class based on NetBeans project BibleGetHelp class

    Private Sub BibleGetHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load        
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


    End Sub
End Class