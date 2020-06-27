Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Net.Mail
Imports System.ComponentModel
Imports System.IO

Public Class Feedback

    Private mailSent As Boolean = False
    Private invalid As Boolean
    Private client As SmtpClient
    Private message As MailMessage


    Public Function IsValidEmail(ByVal strIn As String) As Boolean
        invalid = False
        If String.IsNullOrEmpty(strIn) Then
            Return False
        End If

        ' Use IdnMapping class to convert Unicode domain names.
        Try
            strIn = Regex.Replace(strIn, "(@)(.+)$", AddressOf DomainMapper, RegexOptions.None)
        Catch e As System.TimeoutException
            Return False
        End Try

        If invalid Then
            Return False
        End If

        ' Return true if strIn is in valid e-mail format.
        Try
            Dim result As Boolean = Regex.IsMatch(strIn,
                   "^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                   "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                   RegexOptions.IgnoreCase)
            If result = False Then
            End If
            Return result
        Catch e As System.TimeoutException
            Return False
        End Try
    End Function

    Private Function DomainMapper(match As Match) As String
        ' IdnMapping class with default property values.
        Dim idn As New IdnMapping()

        Dim domainName As String = match.Groups(2).Value
        Try
            domainName = idn.GetAscii(domainName)
        Catch e As ArgumentException
            invalid = True
        End Try
        Return match.Groups(1).Value + domainName
    End Function

    Private Sub SendCompletedCallback(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
        If InvokeRequired Then
            BeginInvoke(New Action(Of AsyncCompletedEventArgs)(AddressOf SendCompleted), e)
        Else
            SendCompleted(e)
        End If
    End Sub

    Private Sub SendCompleted(ByVal e As AsyncCompletedEventArgs)
        ' Get the unique identifier for this asynchronous operation.
        client.Dispose()
        message.Dispose()
        Cursor = Windows.Forms.Cursors.Default
        Dim token As String = CStr(e.UserState)
        Dim str1 As String = String.Empty
        If e.Cancelled Then
            str1 = String.Format("[{0}] Send canceled.", token)
            MsgBox(str1)
        Else
            If e.Error IsNot Nothing Then
                str1 = String.Format("[{0}] {1}", token, e.Error.ToString())
                MsgBox(str1)
            Else
                str1 = String.Format("Thank you for your feedback.")
                mailSent = True
                MsgBox(str1)
                Dispose()
            End If
        End If
    End Sub

    Private Function EnableMessageSend() As Boolean
        Dim ready As Boolean = True
        If Not IsValidEmail(TextBox1.Text) Or String.IsNullOrEmpty(ComboBox1.Text) Or String.IsNullOrEmpty(TextBox2.Text) Then
            ready = False
        End If
        Return ready
    End Function

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        ErrorProvider1.SetError(TextBox1, String.Empty)
    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        Button1.Enabled = EnableMessageSend()
        If Not IsValidEmail(TextBox1.Text) Then
            e.Cancel = True
            ErrorProvider1.SetError(TextBox1, "Invalid email address.")
            TextBox1.Select(0, TextBox1.Text.Length)
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Button1.Enabled = EnableMessageSend()
        If Not String.IsNullOrEmpty(TextBox2.Text) Then
            ErrorProvider1.SetError(TextBox2, String.Empty)
        End If
    End Sub

    Private Sub TextBox2_Validated(sender As Object, e As EventArgs) Handles TextBox2.Validated
        ErrorProvider1.SetError(TextBox2, String.Empty)
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        Button1.Enabled = EnableMessageSend()
        If String.IsNullOrEmpty(TextBox2.Text) Then
            e.Cancel = True
            ErrorProvider1.SetError(TextBox2, "Body cannot be empty.")
        End If
    End Sub

    Private Sub ComboBox1_Validated(sender As Object, e As EventArgs) Handles ComboBox1.Validated
        ErrorProvider1.SetError(ComboBox1, String.Empty)
    End Sub


    Private Sub ComboBox1_Validating(sender As Object, e As CancelEventArgs) Handles ComboBox1.Validating
        Button1.Enabled = EnableMessageSend()
        If String.IsNullOrEmpty(ComboBox1.Text) Then
            e.Cancel = True
            ErrorProvider1.SetError(ComboBox1, "Subject cannot be empty.")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        client = New SmtpClient("mail.bibleget.io")
        If Button1.Text = "SEND FEEDBACK" Then
            Button1.Text = "CANCEL"
            Cursor = Windows.Forms.Cursors.WaitCursor
            message = New MailMessage(New MailAddress(TextBox1.Text), New MailAddress("admin@bibleget.io"))
            message.Subject = "Word Plugin Feedback [" + ComboBox1.Text + "]"
            message.SubjectEncoding = System.Text.Encoding.UTF8
            message.IsBodyHtml = False
            message.Body = TextBox2.Text
            message.BodyEncoding = System.Text.Encoding.UTF8
            If File.Exists(BibleGetAddIn.logFile) Then message.Attachments.Add(New Attachment(BibleGetAddIn.logFile))

            AddHandler client.SendCompleted, AddressOf SendCompletedCallback
            Dim userState As String = "feedback message"
            client.SendAsync(message, userState)
        Else
            If mailSent = False Then
                client.SendAsyncCancel()
            End If
            Button1.Text = "SEND FEEDBACK"
        End If
    End Sub

    Private Sub Feedback_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If client IsNot Nothing Then
            client.Dispose()
        End If
        If message IsNot Nothing Then
            message.Dispose()
        End If
    End Sub

    Private Sub Feedback_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = False
    End Sub
End Class

