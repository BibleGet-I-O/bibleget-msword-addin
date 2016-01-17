Imports Microsoft.Office.Tools.Ribbon
Imports System.Globalization
Imports System.Diagnostics

Public Class Ribbon1

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        InsertBibleQuoteFromDialogBtn.Label = __("Insert quote from input window")
        InsertBibleQuoteFromTextSelectionBtn.Label = __("Insert quote from text selection")
        PreferencesBtn.Label = __("User Preferences")
        HelpBtn.Label = __("Help")
        SendFeedbackBtn.Label = __("Send feedback")
        MakeContributionBtn.Label = __("Contribute")
        AboutBtn.Label = __("About this plugin")
        Dim bibleGetDB As New BibleGetDatabase
        If bibleGetDB.INITIALIZED Then
            StatusBtn.Image = My.Resources.green_checkmark
            StatusBtn.Label = "STATUS: READY"
        End If

    End Sub

    Private Sub PreferencesBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles PreferencesBtn.Click
        Dim oForm As Preferences = New Preferences
        oForm.Show()
    End Sub

    Private Sub InsertBibleQuoteFromDialogBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertBibleQuoteFromDialogBtn.Click
        Dim oForm As InsertQuoteDialog = New InsertQuoteDialog
        oForm.Show()
    End Sub

    Private Sub AboutBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AboutBtn.Click
        Dim oForm As AboutBibleGet = New AboutBibleGet
        oForm.Show()
    End Sub

    Private Sub MakeContributionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MakeContributionBtn.Click
        Dim webAddress As String = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=HDS7XQKGFHJ58"
        Process.Start(webAddress)
    End Sub

    Private Sub StatusBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StatusBtn.Click
        Dim curTime As Date = TimeOfDay
        MsgBox("It is " & curTime.ToShortTimeString & " and all is well! And yes, the Database is correctly initialized and you may proceed to utilize this plugin.")
    End Sub

    Private Sub HelpBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles HelpBtn.Click
        Dim oForm As BibleGetHelp = New BibleGetHelp
        oForm.Show()
        'oForm.ShowDialog()
    End Sub


    Private Sub InsertBibleQuoteFromTextSelectionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertBibleQuoteFromTextSelectionBtn.Click
        Dim progressBar As New ProgressBar
        progressBar.Show()
    End Sub

End Class
