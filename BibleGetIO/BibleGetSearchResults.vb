Imports System.Collections
Imports System.Data.SQLite
Imports System.Globalization
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports System.ComponentModel
Imports System.Net
Imports System.IO
Imports System.Data
Imports Newtonsoft.Json
Imports System.Diagnostics
Imports System.Text.RegularExpressions

Public Class BibleGetSearchResults

    Private DEBUG_MODE As Boolean
    Private PlaceholderText As String
    Private listItems As New Dictionary(Of Integer, String)
    Private colHeader As ColumnHeader
    Private Timer1 As Timers.Timer
    'Private WithEvents _document As HtmlDocument 'do we even use this?
    Private localizedBookNames As LocalizedBibleBooks
    Private searchResultsDT As New DataTable
    Private previewDocumentHead As String
    Private previewDocumentBodyOpen As String
    Private previewDocumentBodyClose As String


    Private Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = BibleGetAddIn.RM.GetString(myStr, BibleGetAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub BibleGetSearchResults_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox("Term to search = " & BibleGetRibbon.TermToSearch & " & version for search = " & BibleGetRibbon.BibleVersionForSearch)
        Text = __("Search for Bible Verses")
        Label2.Text = __("Bible version to search from")
        DEBUG_MODE = My.Settings.DEBUG_MODE
        LoadBibleVersions(BibleVersionForSearch)
        PlaceholderText = __("e.g. creation")
        Label1.Text = __("Term to search")
        Label6.Text = __("Filter results with another term")
        Button1.Text = __("Search")
        Button2.Text = __("Apply filter")
        Button3.Text = __("Order by Reference")
        TermToSearch.Text = PlaceholderText
        Label5.Text = __("Search results")
        localizedBookNames = New LocalizedBibleBooks()
        searchResultsDT.CaseSensitive = False
        searchResultsDT.Columns.Add("IDX", Type.GetType("System.Int32"))
        searchResultsDT.Columns.Add("BOOK", Type.GetType("System.Int32"))
        searchResultsDT.Columns.Add("CHAPTER", Type.GetType("System.Int32"))
        searchResultsDT.Columns.Add("VERSE", Type.GetType("System.String"))
        searchResultsDT.Columns.Add("VERSETEXT", Type.GetType("System.String"))
        searchResultsDT.Columns.Add("SEARCHTERM", Type.GetType("System.String"))
        searchResultsDT.Columns.Add("JSONSTR", Type.GetType("System.String"))

        previewDocumentHead = "<!DOCTYPE html>"
        previewDocumentHead &= "<head>"
        previewDocumentHead &= "<meta charset=""UTF-8"">"
        previewDocumentHead &= "<style type=""text/css"">"
        previewDocumentHead &= "html,body { margin: 0; padding: 0; }
body { border: 1px solid Black; }
#bibleGetSearchResultsTableContainer {
	border: 1px solid #963;
	overflow-y: auto;
    overflow-x: hidden;
    max-height: 100vh;
    width: 100vh;
}

#bibleGetSearchResultsTableContainer table {
  width: 100%;
}

#bibleGetSearchResultsTableContainer th, td { padding: 8px 16px; }

#bibleGetSearchResultsTableContainer thead th {
	position: fixed;
    top: 0;
	background: #C96;
	border-left: 1px solid #EB8;
	border-right: 1px solid #B74;
	border-top: 1px solid #EB8;
	font-weight: normal;
	text-align: center;
    color: White;
    font-weight: bold;
}

#bibleGetSearchResultsTableContainer tbody td {
  border-bottom: 1px groove White;
  background-color: #EFEFEF;
}

#bibleGetSearchResultsTableContainer mark {
  font-weight: bold;
}

a.mark { background-color: yellow; font-weight: bold; padding: 2px 4px; }
a.submark { background-color: lightyellow; padding: 2px 0px; }
a.bmark { background-color: pink; font-weight: bold; padding: 2px 4px; }
a.bsubmark { background-color: #ffe1e6; padding: 2px 0px; }
a.button { padding: 6px; color: DarkBlue; font-weight: bold; background-color: LightBlue; border: 2px outset Blue; border-radius: 3px; display: inline-block; box-shadow: 2px 2px 4px 4px DarkBlue; cursor: pointer; text-decoration: none; }
a.button:hover { background-color: #EEF; }
"
        previewDocumentHead &= "</style>"
        previewDocumentHead &= "</head>"

        previewDocumentBodyOpen = "<body><div id=""bibleGetSearchResultsTableContainer"">
								<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""scrollTable"" id=""SearchResultsTable"">
									<thead class=""fixedHeader"">
										<tr class=""alternateRow""><th>" & __("Action") & "</th><th>" & __("Verse Reference") & "</th><th>" & __("Verse Text") & "</th></tr>
									</thead>
									<tbody class=""scrollContent"">"

        previewDocumentBodyClose = "</tbody></table></div></body>"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TermToSearch.Text = "" Or TermToSearch.Text = PlaceholderText Then
            MsgBox(__("You must type a term to search for"), MsgBoxStyle.Exclamation, "Invalid input")
            Exit Sub
        End If
        If BibleVersionForSearch.SelectedItems.Count < 1 Then
            MsgBox(__("You must choose a Bible version from which to search"), MsgBoxStyle.Exclamation, "Invalid input")
            Exit Sub
        End If
        'perform search
        'MsgBox("Will now perform search for verses that contain """ & TermToSearch.Text & """ in version """ & BibleVersionForSearch.SelectedItems.Item(0).Tag & """...", MsgBoxStyle.Information, "This is a test")

        If Not BackgroundWorker1.IsBusy And Button1.Text = __("Search") Then
            searchResultsDT.Rows.Clear()
            Button3.Text = __("Order by Reference")
            Button2.Text = __("Apply filter")
            Button2.Image = My.Resources.filter
            FilterForTerm.Text = String.Empty
            Button3.Visible = False
            Button2.Visible = False
            FilterForTerm.Visible = False
            Label6.Visible = False

            Button1.Text = __("Cancel")
            Label3.Text = "ELABORATING REQUEST..."
            Dim queryString As String = TermToSearch.Text

            'TextBox1.Clear()
            'TextBox1.BackColor = Drawing.Color.White
            'TextBox1.ForeColor = Drawing.Color.Black
            queryString = queryString.TrimStart
            queryString = queryString.TrimEnd
            'only allow search for one term
            If queryString.Contains(" ") Then
                queryString = queryString.Split(" ").First
            End If
            Dim serverRequestString As String = BibleGetAddIn.BGET_SEARCH_ENDPOINT & "?query=keywordsearch&keyword=" & queryString & "&version=" & BibleVersionForSearch.SelectedItems.Item(0).Tag & "&return=json&appid=msword&pluginversion=" & My.Application.Info.Version.ToString
            If Me.ExactMatchChkBox.Checked Then
                serverRequestString &= "&exactmatch=true"
            End If

            Dim x As BibleGetWorker = New BibleGetWorker("SENDQUERY", serverRequestString)
            BackgroundWorker1.RunWorkerAsync(x)
        ElseIf BackgroundWorker1.IsBusy Or Button1.Text() = __("Cancel") Then
            Button1.Text = __("Send query")
            If BackgroundWorker1.WorkerSupportsCancellation Then
                BackgroundWorker1.CancelAsync()
            End If
        End If

    End Sub

    Private Sub LoadBibleVersions(myListView As ListView)
        'Dim versionCount As Integer
        Dim PreferredVersion As String = My.Settings.PreferredVersions.Split(",").First
        Dim versionLangs As Integer
        Dim bibleGetDB As New BibleGetDatabase
        If bibleGetDB.IsInitialized Then
            Using conn As New SQLiteConnection(bibleGetDB.connectionStr)
                If conn IsNot Nothing Then
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

                        Dim lvGroups As New Dictionary(Of String, ListViewGroup)

                        For Each s As String In keys
                            Dim versionStr As String = versionsObj.SelectToken(s).ToString
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

                        BibleVersions.Sort(New VersionCompareByLang())

                        For Each el As BibleVersion In BibleVersions
                            If Not lvGroups.ContainsKey(el.Lang) Then
                                Dim lvGroup As New ListViewGroup(el.Lang)
                                lvGroups.Add(el.Lang, lvGroup)
                                myListView.Groups.Add(lvGroup)
                                versionLangs += 1
                            End If
                            Dim lvItem As ListViewItem = New ListViewItem()
                            lvItem.Group = lvGroups.Item(el.Lang)
                            lvItem.Text = el.Abbrev & " - " & el.Fullname & " (" & el.Year & ")"
                            lvItem.Tag = el.Abbrev
                            myListView.Items.Add(lvItem)
                            listItems.Add(lvItem.Index, el.Abbrev)
                        Next
                        myListView.View = View.Details
                        colHeader = New ColumnHeader()
                        colHeader.Text = String.Empty
                        colHeader.Width = -2
                        colHeader.TextAlign = HorizontalAlignment.Left
                        myListView.Columns.Add(colHeader)
                        myListView.HeaderStyle = ColumnHeaderStyle.None
                        myListView.Columns(0).Width = myListView.Width - 4 - SystemInformation.VerticalScrollBarWidth

                        For Each item As ListViewItem In myListView.Items
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "item " + item.Index.ToString + ": " + item.Text + ": " + listItems(item.Index))
                            If listItems(item.Index) = PreferredVersion Then
                                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "item " + item.Index.ToString + " is in the PreferredVersions Array!")
                                item.Selected = True
                            End If
                        Next
                    End Using

                Else
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "we seem to have a null connection... arghhh!")
                End If
            End Using
        End If


    End Sub

    Private Sub TermToSearch_GotFocus(sender As Object, e As EventArgs) Handles TermToSearch.GotFocus
        If TermToSearch.Text = PlaceholderText Then
            TermToSearch.Text = ""
        End If
    End Sub

    Private Sub TermToSearch_LostFocus(sender As Object, e As EventArgs) Handles TermToSearch.LostFocus
        If TermToSearch.Text = "" Then
            TermToSearch.Text = PlaceholderText
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
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12
            Dim queryString As String = x.QueryString
            Dim request As HttpWebRequest = WebRequest.Create(queryString)
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
            'the following instruction performs a non safe cross thread operation, look into how to fix this
            'Dim finalString As String = PopulateTableWithSearchResults(responseFromServer, worker)
            result = New BibleGetWorker("WEBRESPONSEELABORATED", responseFromServer)
        ElseIf x.Command = "DOCINJECT" Then
            Dim resultToInject As String = x.QueryString
            Dim honeyBee As BibleGetDocInject = New BibleGetDocInject(worker, e)
            Dim finalString As String = honeyBee.InsertTextAtCurrentSelection(resultToInject)
            worker.ReportProgress(100)
            result = New BibleGetWorker("INJECTIONCOMPLETED", finalString)
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
        ProgressBar2.Value = e.ProgressPercentage
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
            Label3.Text = "REQUEST CANCELED"
            Button1.Text = __("Search")
            ProgressBar2.Value = 0
        ElseIf e.Error IsNot Nothing Then
            Label3.Text = "ERROR: " & e.Error.Message
        Else
            Dim x As BibleGetWorker = e.Result
            Dim command As String = x.Command
            If command = "WEBREQUESTCOMPLETE" Then

                Dim response As HttpWebResponse = x.WebResponse
                'Status of Response
                'CType(response, HttpWebResponse).StatusDescription
                Label3.Text = "HTTP " & response.StatusDescription

                If response.StatusCode = HttpStatusCode.OK Then
                    Dim dataStream As Stream = response.GetResponseStream()
                    Dim reader As New StreamReader(dataStream)
                    Dim responseFromServer As String = reader.ReadToEnd()
                    reader.Close()
                    response.Close()

                    Dim y As BibleGetWorker = New BibleGetWorker("ELABORATEWEBRESPONSE", responseFromServer)
                    BackgroundWorker1.RunWorkerAsync(y)
                Else
                    MsgBox(__("There was a problem communicating with the BibleGet server. Please try again."), MsgBoxStyle.Information, "Status")
                    Button1.Text = __("Search")
                End If

            ElseIf command = "WEBRESPONSEELABORATED" Then
                Label3.Text = "REQUEST COMPLETE"
                Button1.Text = __("Search")
                Dim responseFromServer As String = x.QueryString
                'TextBox1.AppendText(x.QueryString)
                'TextBox1.AppendText(Environment.NewLine)
                Dim result As String = PopulateTableWithSearchResults(responseFromServer)
                TextBox1.AppendText(result)
                TextBox1.AppendText(Environment.NewLine)
                ProgressBar2.Value = 100
                'TextBox1.Text = String.Empty
                'Timer1 = New Timers.Timer()
                'Timer1.Interval = 1000
                'Timer1.Enabled = True
                'Timer1.Start()
                'AddHandler Timer1.Elapsed, AddressOf OnTimedEvent
            ElseIf command = "WEBREQUESTFAILED" Then
                Label3.Text = "INTERNET ERROR"
                TextBox1.AppendText(x.QueryString)
                TextBox1.AppendText(Environment.NewLine)
                'TextBox1.Text = x.QueryString
                Button1.Text = __("Search")
                ProgressBar2.Value = 0
            ElseIf command = "INJECTIONCOMPLETED" Then
                TextBox1.AppendText("Bible quote injection into document completed")
                TextBox1.AppendText(Environment.NewLine)
            End If

        End If
    End Sub

    'Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
    '    Timer1.Dispose()
    '    CloseForm()
    'End Sub

    'Private Sub CloseForm()
    '    If InvokeRequired Then
    '        BeginInvoke(New System.Action(AddressOf CloseForm))
    '    Else
    '        If colHeader IsNot Nothing Then colHeader.Dispose()
    '        Close()
    '    End If
    'End Sub

    Private Function PopulateTableWithSearchResults(SearchResults As String)
        Dim jsObj As JToken = JObject.Parse(SearchResults)
        Dim jRRArray As JArray = jsObj.SelectToken("results")
        Dim infoObj As JObject = jsObj.SelectToken("info")
        Dim searchTerm As String = infoObj.SelectToken("keyword").Value(Of String)()
        Dim versionSearched As String = infoObj.SelectToken("version").Value(Of String)()
        Dim previewDocument As String
        Dim rowsSearchResultsTable As String = ""

        Dim book As Integer
        Dim chapter As Integer
        Dim versenumber As String 'we use string and not integer because some verses contain a letter! otherwise conversion exceptions will be generated
        Dim versetext As String
        Dim resultJsonStr As String

        'Dim curLangIsoCode As String = BibleGetAddIn.locale.TwoLetterISOLanguageName
        'Dim curLangDisplayName As String = New CultureInfo(curLangIsoCode).DisplayName

        Dim numResults As Integer = jRRArray.Count
        Label5.Text = __("Search results") & ": " & numResults & " verses found containing the term """ & searchTerm & """ in version """ & versionSearched & """ "

        ProgressBar2.Value = 25
        If numResults > 0 Then

            Dim workerProgressChunk = Math.Floor(75 / numResults)
            Dim resultCounter As Integer = 0
            For Each result As JToken In jRRArray
                book = result.SelectToken("univbooknum").Value(Of Integer)()
                Dim localizedBook As LocalizedBibleBook = localizedBookNames.GetBookByIndex(book - 1)
                chapter = result.SelectToken("chapter").Value(Of Integer)()
                versenumber = result.SelectToken("verse").Value(Of String)() 'we use string and not integer because some verses contain a letter! otherwise conversion exceptions will be generated
                versetext = result.SelectToken("text").Value(Of String)()
                Dim matchpattern As String = "<(?:[^>=]|='[^']*'|=""[^""]*""|=[^'""][^\s>]*)*>"
                versetext = Regex.Replace(versetext, matchpattern, "")
                versetext = AddMark(versetext, {searchTerm, stripDiacritics(searchTerm)})
                resultJsonStr = JsonConvert.SerializeObject(result, Formatting.None)
                searchResultsDT.Rows.Add(New Object() {resultCounter, book, chapter, versenumber, versetext, searchTerm, resultJsonStr})
                'Debug.Print(resultJsonStr)
                'TextBox1.AppendText(versetext)
                'TextBox1.AppendText(Environment.NewLine)
                rowsSearchResultsTable &= "<tr><td><a href=""#"" class=""button"" id=""row" & resultCounter & """>" & __("Select") & "</a></td><td>" & localizedBook.Fullname & " " & chapter & ":" & versenumber & "</td><td>" & versetext & "</td></tr>"
                ProgressBar2.Value = (ProgressBar2.Value + workerProgressChunk)
                resultCounter += 1
            Next
            Button3.Visible = True
            Button2.Visible = True
            FilterForTerm.Visible = True
            Label6.Visible = True
        Else
            rowsSearchResultsTable &= "<tr><td></td><td></td><td></td></tr>"
        End If
        previewDocument = previewDocumentHead & previewDocumentBodyOpen & rowsSearchResultsTable & previewDocumentBodyClose
        If WebBrowser1.Document Is Nothing Then
            WebBrowser1.DocumentText = previewDocument
        Else
            WebBrowser1.Document.Write(String.Empty)
            WebBrowser1.Document.Write(previewDocument)
            WebBrowser1.Refresh()
            Dim curState As WebBrowserReadyState = WebBrowserReadyState.Uninitialized
            While WebBrowser1.ReadyState < WebBrowserReadyState.Complete
                If WebBrowser1.ReadyState <> curState Then
                    curState = WebBrowser1.ReadyState
                End If
                Application.DoEvents()
            End While
            Dim oLink As HtmlElement
            Dim oLinks As HtmlElementCollection = WebBrowser1.Document.Links
            For Each oLink In oLinks
                oLink.AttachEventHandler("onclick", AddressOf LinkClicked)
            Next
        End If

        '_document = WebBrowser1.Document
        Return "Ok! All done."
    End Function

    Private Function AddMark(verseText As String, searchTerm() As String)
        Dim pattern As String = "\b(\w*?)(" & Join(searchTerm, "|") & ")(\w*?)\b"
        Dim replacement As String = "<a class=""submark"">$1</a><a class=""mark"">$2</a><a class=""submark"">$3</a>"
        Return AddBMark(Regex.Replace(verseText, pattern, replacement, RegexOptions.IgnoreCase), searchTerm)
        'Return Replace(verseText, searchTerm, "<a class=""mark"">" & searchTerm & "</a>", 1, -1, CompareMethod.Text)
    End Function

    Private Function AddBMark(verseText As String, searchTerm() As String)
        Dim upgradedTerm() As String = searchTerm.Select(Function(x) addDiacritics(x)).ToArray()
        'searchTerm = Array.ConvertAll(searchTerm, Function(x) addDiacritics(x))
        Dim pattern As String = "\b(\w*?)(?:(?!" & Join(searchTerm, "|") & "))(" & Join(upgradedTerm, "|") & ")(\w*?)\b"
        Dim replacement As String = "<a class=""bsubmark"">$1</a><a class=""bmark"">$2</a><a class=""bsubmark"">$3</a>"
        Return Regex.Replace(verseText, pattern, replacement, RegexOptions.IgnoreCase)
    End Function

    Private Function stripDiacritics(inText As String)
        Dim normalizedString = inText.Normalize(NormalizationForm.FormD)
        Dim StringBuilder = New StringBuilder()
        Dim c
        For Each c In normalizedString
            Dim UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c)
            If (UnicodeCategory <> UnicodeCategory.NonSpacingMark) Then
                StringBuilder.Append(c)
            End If
        Next

        Return StringBuilder.ToString().Normalize(NormalizationForm.FormC)
    End Function

    Private Function UpgradeDiacritics(ByVal c As Match) As String
        Select Case c.ToString
            Case "a", "A"
                Return "[aA\xC0-\xC5\xE0-\xE5\u0100-\u0105\u01CD\u01CE\u01DE-\u01E1\u01FA\u01FB\u0200-\u0203\u0226\u0227\u023A\u0250-\u0252]"
            Case "e", "E"
                Return "[eE\xC8-\xCB\xE8-\xEB\x12-\x1B\u0204-\u0207\u0228\u0229\u0400\u0401]"
            Case "i", "I"
                Return "[iI\xcc-\xCF\xEC-\xEF\u0128-\u0131\u0196\u0197\u0208-\u020B\u0406\u0407]"
            Case "o", "O"
                Return "[oO\xD2-\xD6\xD8\xF0\xF2-\xF6\xF8\u014C-\u0151\u01A0\u01A1\u01D1\u01D2\u01EA-\u01ED\u01FE\u01FF\u01EA-\u01ED\u01FE\u01FF\u020C-\u020F\u022A-\u0231]"
            Case "u", "U"
                Return "[uU\xD9-\xDC\xF9-\xFC\u0168-\u0173\u01AF-\u01B0\u01D3-\u01DC\u0214-\u0217]"
            Case "y", "Y"
                Return "[yY\xDD\xFD\xFF\u0176-\u0178\u01B3\u01B4\u0232\u0233]"
            Case "c", "C"
                Return "[cC\xC7\xE7\u0106-\u010D\u0187\u0188\u023B\u023C]"
            Case "n", "N"
                Return "[nN\xD1\xF1\u0143-\u014B\u019D\u019E\u01F8\u01F9\u0235]"
            Case "d", "D"
                Return "[dD\xD0\u010E-\u0111\u0189\u0190\u0221]"
            Case "g", "G"
                Return "[gG\u011C-\u0123\u0193-\u0194\u01E4-\u01E7\u01F4\u01F5]"
            Case "h", "H"
                Return "[hH\u0124-\u0127\u021E\u021F]"
            Case "j", "J"
                Return "[jJ\u0134\u0135]"
            Case "k", "K"
                Return "[kk\u0136-\u0138\u0198\u0199\u01E8\u01E9]"
            Case "l", "L"
                Return "[lL\u0139-\u0142\u019A\u019B\u0234\u023D]"
            Case "r", "R"
                Return "[rR\u0154-\u0159\u0210-\u0213]"
            Case "s", "S"
                Return "[sS\u015A-\u0161\u017F\u0218\u0219\u023F]"
            Case "t", "T"
                Return "[tT\u0162-\u0167\u01AB-\u01AE\u021A\u021B\u0236\u023E]"
            Case "w", "W"
                Return "[wW\u0174-\u0175]"
            Case "z", "Z"
                Return "[zZ\u0179-\u017E\u01B5\u01B6\u0224\u0225]"
            Case "b", "B"
                Return "[bB\u0180-\u0183]"
            Case "f", "F"
                Return "[fF\u0191-\u0192]"
            Case "m", "M"
                Return "[mM\u019C]"
            Case "p", "P"
                Return "[\u01A4-\u01A5]"
            Case Else
                Return c.ToString
        End Select
    End Function

    Private Function addDiacritics(term As String) As String
        Dim pattern As String = "."
        Dim r As Regex = New Regex(pattern)
        Dim replacement As MatchEvaluator = New MatchEvaluator(AddressOf UpgradeDiacritics)
        Return r.Replace(term, replacement)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = __("Order by Reference") Then
            searchResultsDT.DefaultView.Sort = "BOOK ASC,CHAPTER ASC,VERSE ASC"
            Button3.Text = __("Order by Importance")
        Else
            searchResultsDT.DefaultView.Sort = "IDX ASC"
            Button3.Text = __("Order by Reference")
        End If
        RefreshSearchResults()
    End Sub

    Private Sub RefreshSearchResults()
        Dim previewDocument As String
        Dim rowsSearchResultsTable As String = ""
        Dim filterTerm As String = String.Empty
        If FilterForTerm.Text IsNot String.Empty Then
            filterTerm = FilterForTerm.Text.TrimStart
            filterTerm = filterTerm.TrimEnd
            If filterTerm.Contains(" ") Then
                filterTerm = filterTerm.Split(" ").First
            End If
        End If
        For Each rowView As DataRowView In searchResultsDT.DefaultView
            Dim row As DataRow = rowView.Row
            Dim book As Integer = row("BOOK")
            Dim bookIdx As Integer = book
            Dim localizedBook As LocalizedBibleBook = localizedBookNames.GetBookByIndex(bookIdx - 1)
            Dim chapter As Integer = row("CHAPTER")
            Dim versenumber As Integer = row("VERSE")
            Dim versetext As String = row("VERSETEXT")
            Dim searchTerm As String = row("SEARCHTERM")
            Dim rowIdx As Integer = row("IDX")
            Dim resultJsonStr As String = row("JSONSTR")
            Dim searchArray() As String = {searchTerm, stripDiacritics(searchTerm)}
            If filterTerm IsNot String.Empty Then
                'versetext = AddMark(versetext, {filterTerm})
                Array.Resize(searchArray, searchArray.Length + 1)
                searchArray(searchArray.Length - 1) = filterTerm
            End If
            versetext = AddMark(versetext, searchArray)
            rowsSearchResultsTable &= "<tr><td><a href=""#"" class=""button"" id=""row" & rowIdx & """>" & __("Select") & "</a></td><td>" & localizedBook.Fullname & " " & chapter & ":" & versenumber & "</td><td>" & versetext & "</td></tr>"
        Next
        previewDocument = previewDocumentHead & previewDocumentBodyOpen & rowsSearchResultsTable & previewDocumentBodyClose
        WebBrowser1.Document.DetachEventHandler("onclick", AddressOf LinkClicked)
        If WebBrowser1.Document Is Nothing Then
            WebBrowser1.DocumentText = previewDocument
        Else
            WebBrowser1.Document.Write(String.Empty)
            WebBrowser1.Document.Write(previewDocument)
            WebBrowser1.Refresh()
        End If
        '_document = WebBrowser1.Document
        Dim curState As WebBrowserReadyState = WebBrowserReadyState.Uninitialized
        While WebBrowser1.ReadyState < WebBrowserReadyState.Complete
            If WebBrowser1.ReadyState <> curState Then
                curState = WebBrowser1.ReadyState
                'Debug.Print("Web Browser state =")
                'Select Case curState
                '    Case WebBrowserReadyState.Uninitialized
                '        Debug.Print("UNINITIALIZED" & Environment.NewLine)
                '    Case WebBrowserReadyState.Loading
                '        Debug.Print("LOADING" & Environment.NewLine)
                '    Case WebBrowserReadyState.Loaded
                '        Debug.Print("LOADED" & Environment.NewLine)
                '    Case WebBrowserReadyState.Interactive
                '        Debug.Print("INTERACTIVE" & Environment.NewLine)
                '    Case WebBrowserReadyState.Complete
                '        Debug.Print("COMPLETE" & Environment.NewLine)
                'End Select
            End If
            Application.DoEvents()
        End While
        Dim oLink As HtmlElement
        Dim oLinks As HtmlElementCollection = WebBrowser1.Document.Links
        For Each oLink In oLinks
            oLink.AttachEventHandler("onclick", AddressOf LinkClicked)
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = __("Apply filter") Then
            If FilterForTerm.Text = "" Then
                MsgBox("Filter term cannot be empty!", MsgBoxStyle.Exclamation, "Error")
            Else
                Button2.Text = __("Remove filter")
                Button2.Image = My.Resources.remove_filter
                Dim filterTerm As String = FilterForTerm.Text.TrimStart
                filterTerm = filterTerm.TrimEnd
                If filterTerm.Contains(" ") Then
                    filterTerm = filterTerm.Split(" ").First
                End If
                'searchResultsDT.CaseSensitive = False
                searchResultsDT.DefaultView.RowFilter = "VERSETEXT LIKE '%" & filterTerm & "%'"
            End If
        Else
            Button2.Text = __("Apply filter")
            Button2.Image = My.Resources.filter
            searchResultsDT.DefaultView.RowFilter = ""
            FilterForTerm.Text = String.Empty
        End If
        RefreshSearchResults()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Dim oLink As HtmlElement
        Dim oLinks As HtmlElementCollection = WebBrowser1.Document.Links
        For Each oLink In oLinks
            oLink.AttachEventHandler("onclick", AddressOf LinkClicked)
        Next

    End Sub

    Private Sub LinkClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim link As HtmlElement = WebBrowser1.Document.ActiveElement
        'MsgBox("a link was clicked", MsgBoxStyle.Information, "Html Document button click event")
        Dim resultIdx As Integer = Integer.Parse(Replace(link.GetAttribute("id"), "row", ""))
        Dim data As String = "{""results"": [" & searchResultsDT.Rows.Item(resultIdx)("JSONSTR") & "]}"
        link.InnerText = __("Inserted")
        link.Style = "color:Purple;background-color:Gray;border: 2px inset Blue;cursor:default;"
        link.DetachEventHandler("onclick", AddressOf LinkClicked)
        Dim x As BibleGetWorker = New BibleGetWorker("DOCINJECT", data)
        BackgroundWorker1.RunWorkerAsync(x)
        'Debug.Print("data(" & resultIdx & ") = " & data & Environment.NewLine)
        'MsgBox("button for row " & resultIdx & " was clicked")
    End Sub

End Class