Imports System.Timers
Imports System.Net
Imports System.IO
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class ProgressBar

    Private Application As Word.Application = Globals.ThisAddIn.Application
    Private PreferredVersions As List(Of String) = My.Settings.PreferredVersions.Split(",").ToList
    Private currentSelection As Word.Selection = Application.Selection

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If BackgroundWorker1.IsBusy Then
            If BackgroundWorker1.WorkerSupportsCancellation Then BackgroundWorker1.CancelAsync()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim result As BibleGetWorker = Nothing

        Dim x As BibleGetWorker = e.Argument
        Dim y As Integer = 0
        If x.Command = "SENDQUERY" Then
            y = 10
            worker.ReportProgress(y)
            Dim queryString As String = x.QueryString
            Dim request As WebRequest = WebRequest.Create(queryString)
            Try
                Dim response As WebResponse = request.GetResponse()
                y += 5
                worker.ReportProgress(y)
                result = New BibleGetWorker("WEBREQUESTCOMPLETE", response)
            Catch ex As Exception
                result = New BibleGetWorker("WEBREQUESTFAILED", ex.Message)
            End Try
        ElseIf x.Command = "ELABORATEWEBRESPONSE" Then
            worker.ReportProgress(20)
            Dim responseFromServer As String = x.QueryString
            Dim honeyBee As BibleGetDocInject = New BibleGetDocInject(worker, e)
            currentSelection.Text = String.Empty
            Dim finalString As String = honeyBee.InsertTextAtCurrentSelection(responseFromServer)
            worker.ReportProgress(100)
            result = New BibleGetWorker("WEBRESPONSEELABORATED", finalString)
        End If

        e.Result = result

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If InvokeRequired Then
            BeginInvoke(New Action(Of ProgressChangedEventArgs)(AddressOf UpdateProgressBar), e)
        Else
            UpdateProgressBar(e)
        End If
        'UpdateProgressBar(e)
    End Sub

    Private Sub UpdateProgressBar(ByVal e As ProgressChangedEventArgs)
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If InvokeRequired Then
            BeginInvoke(New Action(Of RunWorkerCompletedEventArgs)(AddressOf DoWorkCompleted), e)
        Else
            DoWorkCompleted(e)
        End If
        'DoWorkCompleted(e)
    End Sub

    Private Sub DoWorkCompleted(ByVal e As RunWorkerCompletedEventArgs)
        If e.Cancelled = True Then
            STATUS.Text = "REQUEST CANCELED"
            ''Button1.Text = __("Send query")
            ProgressBar1.Value = 0
            Dim tmr As New System.Timers.Timer()
            tmr.Interval = 1000
            tmr.Enabled = True
            tmr.Start()
            AddHandler tmr.Elapsed, AddressOf OnTimedEvent
        ElseIf e.Error IsNot Nothing Then
            STATUS.Text = "ERROR: " & e.Error.Message
        Else
            Dim x As BibleGetWorker = e.Result
            Dim command As String = x.Command
            If command = "WEBREQUESTCOMPLETE" Then

                Dim response As HttpWebResponse = x.WebResponse
                'Status of Response
                'CType(response, HttpWebResponse).StatusDescription
                STATUS.Text = "HTTP " & response.StatusDescription

                If response.StatusCode = HttpStatusCode.OK Then
                    Dim dataStream As Stream = response.GetResponseStream()
                    Dim reader As New StreamReader(dataStream)
                    Dim responseFromServer As String = reader.ReadToEnd()
                    reader.Close()
                    response.Close()

                    Dim y As BibleGetWorker = New BibleGetWorker("ELABORATEWEBRESPONSE", responseFromServer)
                    BackgroundWorker1.RunWorkerAsync(y)
                Else
                    Label2.Text = __("There was a problem communicating with the BibleGet server. Please try again.")
                    'Button1.Text = __("Send query")
                End If

            ElseIf command = "WEBRESPONSEELABORATED" Then
                STATUS.Text = "REQUEST COMPLETE"
                'Button1.Text = __("Send query")
                'Label2.Text = x.QueryString
                Label2.Text = String.Empty
                Dim tmr As New System.Timers.Timer()
                tmr.Interval = 2000
                tmr.Enabled = True
                tmr.Start()
                AddHandler tmr.Elapsed, AddressOf OnTimedEvent
            ElseIf command = "WEBREQUESTFAILED" Then
                STATUS.Text = "INTERNET ERROR"
                Label2.Text = x.QueryString
                'Button1.Text = __("Send query")
                ProgressBar1.Value = 0
            End If

        End If
    End Sub

    Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        CloseForm()
    End Sub

    Private Sub CloseForm()
        If InvokeRequired Then
            BeginInvoke(New System.Action(AddressOf CloseForm))
        Else
            Me.Close()
        End If
    End Sub

    Private Sub ProgressBar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Text = __("Cancel")
        STATUS.Text = "ELABORATING REQUEST..."
        ProgressBar1.Value = 20
        Dim queryString As String = currentSelection.Text
        'Diagnostics.Debug.WriteLine("queryString = <" + queryString + ">")
        Dim helperFunctions As BibleGetHelper = New BibleGetHelper
        If queryString Is String.Empty Or queryString = vbCrLf Or queryString = vbLf Or queryString = vbCr Then
            Label2.BackColor = Drawing.Color.LightPink
            Label2.ForeColor = Drawing.Color.DarkRed
            Label2.Text = __("You cannot send an empty query.")
            STATUS.Text = "REQUEST FAILED"
            ProgressBar1.Value = 100
            Me.Cursor = Cursors.Default
            Button1.Cursor = Cursors.Default
        Else
            queryString = New String(queryString.Where(Function(x) (Char.IsWhiteSpace(x) Or Char.IsLetterOrDigit(x) Or x = "," Or x = "." Or x = ":" Or x = "-" Or x = ";")).ToArray())
            Dim integrityResult As Boolean = helperFunctions.integrityCheck(queryString, PreferredVersions.ToArray)
            If integrityResult Then
                queryString = System.Uri.EscapeDataString(queryString)
                Dim queryVersions As String = System.Uri.EscapeDataString(String.Join(",", PreferredVersions))
                Dim serverRequestString As String = "http://query.bibleget.io/index2.php?query=" & queryString & "&version=" & queryVersions & "&return=json&appid=msword&pluginversion=" & My.Application.Info.Version.ToString

                Dim x As BibleGetWorker = New BibleGetWorker("SENDQUERY", serverRequestString)
                BackgroundWorker1.RunWorkerAsync(x)
            Else
                Label2.BackColor = Drawing.Color.Pink
                Label2.ForeColor = Drawing.Color.DarkRed
                Label2.Text = ""
                Dim counter As Integer = 0
                For Each errMessage As String In helperFunctions.errorMessages
                    Label2.Text = Label2.Text & (counter & ") ERROR" & ": " & errMessage & Environment.NewLine)
                    counter += 1
                Next
                'Button1.Text = __("Send query")
                STATUS.Text = "REQUEST ABORTED"
                ProgressBar1.Value = 100
                Cursor = Cursors.Default
                Button1.Cursor = Cursors.Default
            End If
        End If

    End Sub
End Class