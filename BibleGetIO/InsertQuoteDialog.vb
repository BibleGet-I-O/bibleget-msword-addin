Imports System.Net
Imports System.IO
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SQLite
Imports Newtonsoft.Json.Linq
Imports System.Collections
Imports System.Globalization
Imports System.Timers


Public Class InsertQuoteDialog

    'Private Application As Word.Application = Globals.ThisAddIn.Application
    Private PreferredVersions As List(Of String) = My.Settings.PreferredVersions.Split(",").ToList
    Private listItems As New Dictionary(Of Integer, String)
    Private colHeader As ColumnHeader
    Private helperFunctions As BibleGetHelper = New BibleGetHelper
    Private INITIALIZING As Boolean = True
    Private Timer1 As Timers.Timer
    Private DEBUG_MODE As Boolean

    Private Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = BibleGetAddIn.RM.GetString(myStr, BibleGetAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub InsertQuoteDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DEBUG_MODE = My.Settings.DEBUG_MODE
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Loading InsertQuoteDialog")
        Text = __("Insert quote from input window")
        Label3.Text = __("Type the desired Bible Quote using standard notation:")
        TextBox2.Text = __("(e.g. Mt 1,1-10.12-15;5,3-4;Jn 3,16)")
        Label4.Text = __("Choose version (or versions)")
        Button1.Text = __("Send query")
        ToolTip1.SetToolTip(Button1, __("Sends the request to the server and returns the results to the document."))
        INITIALIZING = False
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "about to call LoadBibleVersions")
        LoadBibleVersions(ListView1)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "LoadBibleVersions called")
        TextBox2.Focus()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not BackgroundWorker1.IsBusy And Button1.Text() = __("Send query") Then
            Button1.Text = __("Cancel")
            Label2.Text = "ELABORATING REQUEST..."
            Dim queryString As String = TextBox2.Text
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "dirty queryString = " + queryString)
            queryString = New String(queryString.Where(Function(x) (Char.IsWhiteSpace(x) Or Char.IsLetterOrDigit(x) Or x = "," Or x = "." Or x = ":" Or x = "-" Or x = ";")).ToArray())
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "clean queryString = " + queryString)
            'First we perform some verifications to make sure we are dealing with a good query
            Dim integrityResult As Boolean = helperFunctions.IntegrityCheck(queryString, PreferredVersions.ToArray)

            If integrityResult Then
                TextBox1.Clear()
                TextBox1.BackColor = Drawing.Color.White
                TextBox1.ForeColor = Drawing.Color.Black
                'When we are sure that this is a good query, we can finally prepare it for the web request
                queryString = Uri.EscapeDataString(queryString)
                Dim queryVersions As String = Uri.EscapeDataString(String.Join(",", PreferredVersions))
                Dim serverRequestString As String = BibleGetAddIn.BGET_ENDPOINT & "?query=" & queryString & "&version=" & queryVersions & "&return=json&appid=msword&pluginversion=" & My.Application.Info.Version.ToString

                Dim x As BibleGetWorker = New BibleGetWorker("SENDQUERY", serverRequestString)
                BackgroundWorker1.RunWorkerAsync(x)
            Else
                TextBox1.BackColor = Drawing.Color.Pink
                TextBox1.ForeColor = Drawing.Color.DarkRed

                Dim counter As Integer = 0
                TextBox1.Clear()
                For Each errMessage As String In helperFunctions.ErrorMessages
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
            Catch ex As WebException
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
                Timer1 = New Timers.Timer()
                Timer1.Interval = 1000
                Timer1.Enabled = True
                Timer1.Start()
                AddHandler Timer1.Elapsed, AddressOf OnTimedEvent
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


    Private Sub LoadBibleVersions(myListView As ListView)
        'Dim versionCount As Integer
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Entering LoadBibleVersions Sub...")
        Dim versionLangs As Integer
        Dim bibleGetDB As New BibleGetDatabase
        If bibleGetDB.IsInitialized Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "bibleGetDB.IsInitialized")
            Using conn As New SQLiteConnection(bibleGetDB.connectionStr)
                If conn IsNot Nothing Then
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "SQLiteConnection OK")
                    conn.Open()
                    Using sqlQuery As New SQLiteCommand(conn)
                        Dim queryString As String = "SELECT VERSIONS FROM METADATA WHERE ID=0"
                        sqlQuery.CommandText = queryString
                        Dim versionsString As String = sqlQuery.ExecuteScalar()
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "versionsString = " + versionsString)
                        Dim versionsObj As JObject = JObject.Parse(versionsString)
                        Dim keys() As String = versionsObj.Properties().Select(Function(p) p.Name).ToArray()
                        'versionCount = keys.Length
                        Dim BibleVersions As New ArrayList()

                        For Each s As String In keys
                            Dim versionStr As String = versionsObj.SelectToken(s).Value(Of String)
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "versionStr = " + versionStr)
                            Dim strArray() As String = versionStr.Split("|")
                            Dim fullLanguageName As String = ""
                            Dim languageName As String
                            Try
                                Dim myCulture As CultureInfo = New CultureInfo(strArray(2), False)
                                fullLanguageName = myCulture.DisplayName
                                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & fullLanguageName)
                                languageName = fullLanguageName.ToUpper(CultureInfo.CurrentUICulture)
                            Catch e As CultureNotFoundException
                                If strArray(2) = "la" Then
                                    Select Case CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
                                        Case "en"
                                            fullLanguageName = "Latin"
                                        Case "es"
                                            fullLanguageName = "Latín"
                                        Case "fr"
                                            fullLanguageName = "Latin"
                                        Case "it"
                                            fullLanguageName = "Latino"
                                        Case "de"
                                            fullLanguageName = "Lateinische"
                                        Case "ar"
                                            fullLanguageName = "لاتينية"
                                        Case "pt"
                                            fullLanguageName = "Latim"
                                        Case "sr"
                                            fullLanguageName = "Латински"
                                        Case Else
                                            fullLanguageName = "Latin"
                                    End Select
                                End If
                            Catch e As Exception
                                MsgBox("There was an error: " & e.Message & ". Please send feedback about this error to the add-in author using the Send Feedback menu item.", MsgBoxStyle.Critical, "ERROR!")
                            End Try
                            languageName = fullLanguageName.ToUpper(CultureInfo.CurrentUICulture)
                            BibleVersions.Add(New BibleVersion(s, strArray(0), strArray(1), languageName))

                        Next
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "BibleVersions ArrayList should now be built")

                        BibleVersions.Sort(New VersionCompareByLang())
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "BibleVersions ArrayList should now be sorted using VersionCompareByLang")

                        Dim lvGroups As New Dictionary(Of String, ListViewGroup)
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "About to start building listview based on Bible versions sorted by lang")
                        For Each el As BibleVersion In BibleVersions
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "BibleVersions ArrayList should now be sorted using VersionCompareByLang")
                            If Not lvGroups.ContainsKey(el.Lang) Then
                                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & el.Lang & " cannot be found in the list view groups dictionary, now creating new listviewgroup and adding to lvGroups dictionary...")
                                Dim lvGroup As New ListViewGroup(el.Lang)
                                lvGroups.Add(el.Lang, lvGroup)
                                myListView.Groups.Add(lvGroup)
                                versionLangs += 1
                            End If
                            Dim lvItem As ListViewItem = New ListViewItem()
                            lvItem.Group = lvGroups.Item(el.Lang)
                            lvItem.Text = el.Abbrev & " - " & el.Fullname & " (" & el.Year & ")"
                            myListView.Items.Add(lvItem)
                            listItems.Add(lvItem.Index, el.Abbrev)
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "created new listview item with index " & lvItem.Index & " and abbreviation value " & el.Abbrev)
                        Next
                        myListView.View = View.Details
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set myListView.View to View.Details")
                        colHeader = New ColumnHeader()
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "created new ColumnHeader")
                        colHeader.Text = String.Empty
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set colHeader Text to empty string")
                        colHeader.Width = -2
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set colHeader width to -2")
                        colHeader.TextAlign = HorizontalAlignment.Left
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set colHeader.TextAlign to HorizontalAlignment.Left")
                        myListView.Columns.Add(colHeader)
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "added colHeader to myListView.Columns")
                        myListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set myListView.HeaderStyle to ColumnHeaderStyle.None")
                        myListView.Columns(0).Width = myListView.Width - 4 - SystemInformation.VerticalScrollBarWidth
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "set myListView.Columns(0).Width to myListView.Width-4-SystemInformation.VerticalScrollBarWidth")

                        For Each item As ListViewItem In myListView.Items
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "item " + item.Index.ToString + ": " + item.Text + ": " + listItems(item.Index))
                            If Array.IndexOf(PreferredVersions.ToArray, listItems(item.Index)) <> -1 Then
                                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "item " + item.Index.ToString + " is in the PreferredVersions Array!")
                                item.Selected = True
                            End If
                        Next
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "ListView should now be populated")
                    End Using

                Else
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "we seem to have a null connection... arghhh!")
                End If
            End Using
        End If


    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        If Not INITIALIZING Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "I just noticed a change in the ListView selected indices!")
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
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "I am ignoring a change in the ListView selected indices, lalalala I cannot hear a thing")
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
            If colHeader IsNot Nothing Then colHeader.Dispose()
            Close()
        End If
    End Sub


    Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Timer1.Dispose()
        CloseForm()
    End Sub


End Class





