Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SQLite
Imports System.Data
Imports Newtonsoft.Json.Linq
Imports System.Collections
Imports System.Globalization
Imports System.Timers
Imports System.Text.RegularExpressions


Public Class InsertQuoteDialog

    Private Application As Word.Application = Globals.ThisAddIn.Application
    Private PreferredVersions As List(Of String) = My.Settings.PreferredVersions.Split(",").ToList
    Private listItems As New Dictionary(Of Integer, String)
    Private helperFunctions As BibleGetHelper = New BibleGetHelper
    Private INITIALIZING As Boolean = True

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub InsertQuoteDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = __("Insert quote from input window")
        Label3.Text = __("Type the desired Bible Quote using standard notation:")
        TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)")
        Label4.Text = __("Choose version (or versions)")
        Button1.Text = __("Send query")
        ToolTip1.SetToolTip(Button1, __("Sends the request to the server and returns the results to the document."))
        LoadBibleVersions()
        TextBox2.Focus()
        INITIALIZING = False
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not BackgroundWorker1.IsBusy And Button1.Text() = __("Send query") Then
            Button1.Text = __("Cancel")
            Label2.Text = "ELABORATING REQUEST..."
            Dim queryString As String = TextBox2.Text
            'Diagnostics.Debug.WriteLine("dirty queryString = " + queryString)
            queryString = New String(queryString.Where(Function(x) (Char.IsWhiteSpace(x) Or Char.IsLetterOrDigit(x) Or x = "," Or x = "." Or x = ":" Or x = "-" Or x = ";")).ToArray())
            'Diagnostics.Debug.WriteLine("clean queryString = " + queryString)
            'First we perform some verifications to make sure we are dealing with a good query
            Dim integrityResult As Boolean = helperFunctions.integrityCheck(queryString, PreferredVersions.ToArray)

            If integrityResult Then
                TextBox1.Clear()
                TextBox1.BackColor = Drawing.Color.White
                TextBox1.ForeColor = Drawing.Color.Black
                'When we are sure that this is a good query, we can finally prepare it for the web request
                queryString = System.Uri.EscapeDataString(queryString)
                Dim queryVersions As String = System.Uri.EscapeDataString(String.Join(",", PreferredVersions))
                Dim serverRequestString As String = "http://query.bibleget.io/index2.php?query=" & queryString & "&version=" & queryVersions & "&return=json&appid=msword&pluginversion=" & My.Application.Info.Version.ToString

                Dim x As BibleGetWorker = New BibleGetWorker("SENDQUERY", serverRequestString)
                BackgroundWorker1.RunWorkerAsync(x)
            Else
                TextBox1.BackColor = Drawing.Color.Pink
                TextBox1.ForeColor = Drawing.Color.DarkRed

                Dim counter As Integer = 0
                TextBox1.Clear()
                For Each errMessage As String In helperFunctions.errorMessages
                    TextBox1.AppendText(counter & ") ERROR" & ": " & errMessage & Environment.NewLine)
                    counter += 1
                Next
                Button1.Text = __("Send query")
                Label2.Text = "REQUEST ABORTED"
            End If
        ElseIf BackgroundWorker1.IsBusy Or Button1.Text() = __("Cancel") Then
            Button1.Text = __("Send query")
            If BackgroundWorker1.WorkerSupportsCancellation Then
                BackgroundWorker1.CancelAsync()
            End If
        End If
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ' Get the BackgroundWorker object that raised this event.
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
    End Sub

    Private Sub DoWorkCompleted(ByVal e As RunWorkerCompletedEventArgs)
        If e.Cancelled = True Then
            Label2.Text = "REQUEST CANCELED"
            Button1.Text = __("Send query")
            ProgressBar1.Value = 0
        ElseIf e.Error IsNot Nothing Then
            Label2.Text = "ERROR: " & e.Error.Message
        Else
            Dim x As BibleGetWorker = e.Result
            Dim command As String = x.Command
            If command = "WEBREQUESTCOMPLETE" Then

                Dim response As HttpWebResponse = x.WebResponse
                'Status of Response
                'CType(response, HttpWebResponse).StatusDescription
                Label2.Text = "HTTP " & response.StatusDescription

                If response.StatusCode = HttpStatusCode.OK Then
                    Dim dataStream As Stream = response.GetResponseStream()
                    Dim reader As New StreamReader(dataStream)
                    Dim responseFromServer As String = reader.ReadToEnd()
                    reader.Close()
                    response.Close()

                    Dim y As BibleGetWorker = New BibleGetWorker("ELABORATEWEBRESPONSE", responseFromServer)
                    BackgroundWorker1.RunWorkerAsync(y)
                Else
                    TextBox1.Text = __("There was a problem communicating with the BibleGet server. Please try again.")
                    Button1.Text = __("Send query")
                End If

            ElseIf command = "WEBRESPONSEELABORATED" Then
                Label2.Text = "REQUEST COMPLETE"
                Button1.Text = __("Send query")
                'TextBox1.Text = x.QueryString
                TextBox1.Text = String.Empty
                Dim tmr As New System.Timers.Timer()
                tmr.Interval = 1000
                tmr.Enabled = True
                tmr.Start()
                AddHandler tmr.Elapsed, AddressOf OnTimedEvent
            ElseIf command = "WEBREQUESTFAILED" Then
                Label2.Text = "INTERNET ERROR"
                TextBox1.Text = x.QueryString
                Button1.Text = __("Send query")
                ProgressBar1.Value = 0
            End If

        End If
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        If TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)") Then
            TextBox2.ForeColor = Drawing.Color.Black
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text = "" Then
            TextBox2.ForeColor = Drawing.Color.Gray
            TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)")
        End If
    End Sub


    Private Sub LoadBibleVersions()
        Dim versionCount As Integer
        Dim versionLangs As Integer
        Dim bibleGetDB As New BibleGetDatabase
        Dim conn As SQLiteConnection
        If bibleGetDB.INITIALIZED Then
            conn = bibleGetDB.connect()
            If conn IsNot Nothing And conn.State = ConnectionState.Open Then
                Using conn
                    Using sqlQuery As New SQLiteCommand(conn)
                        Dim queryString As String = "SELECT VERSIONS FROM METADATA WHERE ID=0"
                        sqlQuery.CommandText = queryString
                        Dim versionsString As String = sqlQuery.ExecuteScalar()
                        'Diagnostics.Debug.WriteLine("versionsString = " + versionsString)
                        Dim versionsObj As JObject = JObject.Parse(versionsString)
                        Dim keys() As String = versionsObj.Properties().Select(Function(p) p.Name).ToArray()
                        versionCount = keys.Length
                        Dim BibleVersions As New ArrayList()

                        Dim lvGroups As New Dictionary(Of String, ListViewGroup)

                        For Each s As String In keys
                            Dim versionStr As String = versionsObj.SelectToken(s).ToString
                            Dim strArray() As String = versionStr.Split("|")
                            Dim myCulture As CultureInfo = New CultureInfo(strArray(2), False)
                            Dim fullLanguageName As String = myCulture.DisplayName
                            'Diagnostics.Debug.WriteLine(fullLanguageName)
                            Dim languageName As String = fullLanguageName.ToUpper
                            BibleVersions.Add(New BibleVersion(s, strArray(0), strArray(1), languageName))
                        Next

                        BibleVersions.Sort(New VersionCompareByLang())

                        For Each el As BibleVersion In BibleVersions
                            If Not lvGroups.ContainsKey(el.Lang) Then
                                Dim lvGroup As New ListViewGroup(el.Lang)
                                lvGroups.Add(el.Lang, lvGroup)
                                ListView1.Groups.Add(lvGroup)
                                versionLangs += 1
                            End If
                            Dim lvItem As ListViewItem = New ListViewItem()
                            lvItem.Group = lvGroups.Item(el.Lang)
                            lvItem.Text = el.Abbrev & " - " & el.Fullname & " (" & el.Year & ")"
                            ListView1.Items.Add(lvItem)
                            listItems.Add(lvItem.Index, el.Abbrev)
                        Next
                        ListView1.View = View.Details
                        Dim colHeader As ColumnHeader = New ColumnHeader()
                        colHeader.Text = "Available Bible Versions"
                        colHeader.Width = -2
                        colHeader.TextAlign = HorizontalAlignment.Left
                        ListView1.Columns.Add(colHeader)
                        ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
                        ListView1.Columns(0).Width = ListView1.Width - 4 - SystemInformation.VerticalScrollBarWidth
                        For Each item As ListViewItem In ListView1.Items
                            'Diagnostics.Debug.WriteLine("item " + item.Index.ToString + ": " + item.Text + ": " + listItems(item.Index))
                            If Array.IndexOf(PreferredVersions.ToArray, listItems(item.Index)) <> -1 Then
                                'Diagnostics.Debug.WriteLine("item " + item.Index.ToString + " is in the PreferredVersions Array!")
                                item.Selected = True
                            End If
                        Next
                    End Using
                End Using
            Else
                'Diagnostics.Debug.WriteLine("we seem to have a null connection... arghhh!")
            End If
        End If


    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        If Not INITIALIZING Then
            'Diagnostics.Debug.WriteLine("I just noticed a change in the ListView selected indices!")
            Dim selectedItems As ListView.SelectedListViewItemCollection = ListView1.SelectedItems
            Dim item As ListViewItem
            Dim versionsList(selectedItems.Count - 1) As String
            Dim counter As Integer = 0
            For Each item In selectedItems
                versionsList(counter) = listItems(item.Index)
                counter += 1
            Next
            PreferredVersions = versionsList.ToList
            My.Settings.PreferredVersions = String.Join(",", PreferredVersions)
        Else
            'Diagnostics.Debug.WriteLine("I am ignoring a change in the ListView selected indices, lalalala I cannot hear a thing")
        End If
        If TextBox2.Text = String.Empty Or TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)") Or ListView1.SelectedItems.Count < 1 Then
            Button1.Enabled = False
        ElseIf TextBox2.Text <> String.Empty And ListView1.SelectedItems.Count > 0 Then
            Button1.Enabled = True
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = String.Empty Or TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)") Or ListView1.SelectedItems.Count < 1 Then
            Button1.Enabled = False
        ElseIf TextBox2.Text <> String.Empty And ListView1.SelectedItems.Count > 0 Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub CloseForm()
        If InvokeRequired Then
            BeginInvoke(New System.Action(AddressOf CloseForm))
        Else
            Me.Close()
        End If
    End Sub


    Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        CloseForm()
    End Sub


End Class





