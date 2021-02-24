Imports System.Data.SQLite
Imports Newtonsoft.Json.Linq
Imports System.Collections
Imports System.Globalization
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Net
Imports System.Diagnostics
Imports System.IO


Public NotInheritable Class AboutBibleGet

    Private langcodes As New Dictionary(Of String, String)
    Private bibleGetDB As BibleGetDatabase
    Private versionCount As Integer
    Private versionLangs As Integer
    Private booksLangs As Integer
    Private langsLocalized As List(Of String) = New List(Of String)
    Private colHeader As ColumnHeader
    Private localFile As String
    'Private eventHandled As Boolean
    Private WithEvents updateProcess As New Process
    Private elapsedTime As TimeSpan
    Private DEBUG_MODE As Boolean = My.Settings.DEBUG_MODE
    Private WordApplication As Word.Application = Globals.BibleGetAddIn.Application

    Private Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = BibleGetAddIn.RM.GetString(myStr, BibleGetAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub AboutBibleGet_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If colHeader IsNot Nothing Then colHeader.Dispose()
    End Sub

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Imposta il titolo del form.
        'Dim ApplicationTitle As String
        'If My.Application.Info.Title <> "" Then
        '    ApplicationTitle = My.Application.Info.Title
        'Else
        '    ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        'End If

        'Me.Text = String.Format("Informazioni su {0}", ApplicationTitle)
        Text = __("About this plugin")

        ' Inizializza tutto il testo visualizzato nella finestra di dialogo Informazioni su.
        LabelProductName.Text = __(My.Application.Info.ProductName)
        If My.Settings.NewVersionExists Then
            LabelVersion.Text = __("Version") & " " & My.Application.Info.Version.ToString & " (there is a newer version " & My.Settings.NewVersion & ", click here to update)"
            LabelVersion.ForeColor = Drawing.Color.DarkBlue
            LabelVersion.BackColor = Drawing.Color.LightYellow
            LabelVersion.Cursor = Cursors.Hand
        Else
            LabelVersion.Text = __("Version") & " " & My.Application.Info.Version.ToString
            LabelVersion.Cursor = Cursors.Default
        End If
        LabelCopyright.Text = My.Application.Info.Copyright

        'Me.TextBoxDescription.Text = __(My.Application.Info.Description)
        Dim descr As String
        descr = __("This plugin was developed by <b>John R. D'Orazio</b>, a priest in the diocese of Rome.") _
                + " " _
                + String.Format(__("It is a part of the <b>BibleGet Project</b> at {0}."), "<span style='color:Blue;'>https://www.bibleget.io</span>") _
                + " " _
                + __("The author would like to thank <b>Giovanni Gregori</b> and <b>Simone Urbinati</b> for their code contributions.") _
                + " " _
                + __("The <b>BibleGet Project</b> is an independent project born from the personal initiative of John R. D'Orazio, and is not funded by any kind of corporation.") _
                + " " _
                + __("All of the expenses of the project server and domain, which amount to €200 a year, are accounted for personally by the author. All code contributions and development are entirely volunteered.") _
                + " " _
                + __("If you like the plugin and find it useful, please consider contributing even a small amount to help keep this project running. Even just €1 can make a difference. You can contribute using the appropriate menu item in this plugin's menu.")
        WebBrowser1.DocumentText = "<!DOCTYPE html><head></head><body style=""background-color:transparent;margin:0;padding:0 1em;line-height:120%;""><div style=""font-size:10pt;text-align:justify;"">" & descr & "</div></body>"
        'Me.TextBoxDescription.Text =

        ServerData.Text = __("Current information from the BibleGet Server:")
        Button1.Text = __("RENEW SERVER DATA")

        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "about to call BuildLangCodes Sub...")
        BuildLangCodes()
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "BuildLangCodes Sub completed")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "about to call prepareDynamicInformation Sub...")
        prepareDynamicInformation()
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "prepareDynamicInformation Sub completed")

        CurrentInfo.Text = String.Format(__("The BibleGet database currently supports {0} versions of the Bible in {1} different languages:"), versionCount, versionLangs)
        ServerDataLangsCount.Text = String.Format(__("The BibleGet engine currently understands the names of the books of the Bible in {0} different languages:"), booksLangs)
        ServerDataLangs.Text = String.Join(", ", langsLocalized)

        UpdateCheckBtn.Text = "Check for Updates (last check was " & My.Settings.UpdateCheck.ToLongDateString & " at " & My.Settings.UpdateCheck.ToLongTimeString & ")"
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
    End Sub


    Private Function localizeLanguage(ByVal language As String) As String
        language = language.ToUpper
        Dim langCode As String = String.Empty
        If langcodes.TryGetValue(language, langCode) Then
            Dim myCulture As CultureInfo = New CultureInfo(langCode, False)
            Return myCulture.DisplayName.ToUpper
        Else
            Return language
        End If
        Return Nothing
    End Function

    Private Sub BuildLangCodes()
        'ISO language codes supported by Microsoft, taken from https://msdn.microsoft.com/it-it/goglobal/bb896001.aspx
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Entering BuildLangCodes Sub...")
        langcodes.Add("AFRIKAANS", "af")
        langcodes.Add("ALBANIAN", "sq")
        langcodes.Add("AMHARIC", "am")
        langcodes.Add("ARABIC", "ar")
        langcodes.Add("ARMENIAN", "hy")
        langcodes.Add("ASSAMESE", "as")
        langcodes.Add("AZERBAIJANI", "az")
        langcodes.Add("AZERI", "az")
        langcodes.Add("BASHKIR", "ba")
        langcodes.Add("BASQUE", "eu")
        langcodes.Add("BELARUSIAN", "be")
        langcodes.Add("BENGALI", "bn")
        langcodes.Add("BOSNIAN", "bs")
        langcodes.Add("BRETON", "br")
        langcodes.Add("BULGARIAN", "bg")
        langcodes.Add("CAMBODIAN", "km")
        langcodes.Add("CATALAN", "ca")
        langcodes.Add("CHINESE", "zh")
        langcodes.Add("CORSICAN", "co")
        langcodes.Add("CROATIAN", "hr")
        langcodes.Add("CZECH", "cs")
        langcodes.Add("DANISH", "da")
        langcodes.Add("DARI", "prs")
        langcodes.Add("DIVEHI", "div")
        langcodes.Add("DUTCH", "nl")
        langcodes.Add("ENGLISH", "en")
        langcodes.Add("ESTONIAN", "et")
        langcodes.Add("FAROESE", "fo")
        langcodes.Add("FILIPINO", "fil")
        langcodes.Add("FINNISH", "fi")
        langcodes.Add("FRENCH", "fr")
        langcodes.Add("FRISIAN", "fy")
        langcodes.Add("GALICIAN", "gl")
        langcodes.Add("GEORGIAN", "ka")
        langcodes.Add("GERMAN", "de")
        langcodes.Add("GREEK", "el")
        langcodes.Add("GREENLANDIC", "kl")
        langcodes.Add("GUJARATI", "gu")
        langcodes.Add("HAUSA", "ha")
        langcodes.Add("HEBREW", "he")
        langcodes.Add("HINDI", "hi")
        langcodes.Add("HUNGARIAN", "hu")
        langcodes.Add("ICELANDIC", "is")
        langcodes.Add("IGBO", "ig")
        langcodes.Add("INDONESIAN", "id")
        langcodes.Add("INUKTITUT", "iu")
        langcodes.Add("IRISH", "ga")
        langcodes.Add("ISIXHOSA", "xh")
        langcodes.Add("ISIZULU", "zu")
        langcodes.Add("ITALIAN", "it")
        langcodes.Add("JAPANESE", "ja")
        langcodes.Add("KANNADA", "kn")
        langcodes.Add("KAZAKH", "kk")
        langcodes.Add("KHMER", "km")
        langcodes.Add("K'ICHE", "qut")
        langcodes.Add("KINYARWANDA", "rw")
        langcodes.Add("KISWAHILI", "sw")
        langcodes.Add("KONKANI", "kok")
        langcodes.Add("KOREAN", "ko")
        langcodes.Add("KYRGYZ", "ky")
        langcodes.Add("LATIN", "la")
        langcodes.Add("LAO", "lo")
        langcodes.Add("LAOTHIAN", "lo")
        langcodes.Add("LATVIAN", "lv")
        langcodes.Add("LITHUANIAN", "lt")
        langcodes.Add("LOWER_SORBIAN", "wee")
        langcodes.Add("LUXEMBOURGHISH", "lb")
        langcodes.Add("MACEDONIAN", "mk")
        langcodes.Add("MALAY", "ms")
        langcodes.Add("MALAYALAM", "ml")
        langcodes.Add("MALTESE", "mt")
        langcodes.Add("MAORI", "mi")
        langcodes.Add("MAPUDUNGUN", "arn")
        langcodes.Add("MARATHI", "mr")
        langcodes.Add("MOHAWK", "moh")
        langcodes.Add("MONGOLIAN", "mn")
        langcodes.Add("NEPALI", "ne")
        langcodes.Add("NORWEGIAN", "no")
        langcodes.Add("OCCITAN", "oc")
        langcodes.Add("ORIYA", "or")
        langcodes.Add("PASHTO", "ps")
        langcodes.Add("PERSIAN", "fa")
        langcodes.Add("POLISH", "pl")
        langcodes.Add("PORTUGUESE", "pt")
        langcodes.Add("PUNJABI", "pa")
        langcodes.Add("QUECHUA", "quz")
        langcodes.Add("ROMANIAN", "ro")
        langcodes.Add("ROMANSH", "rm")
        langcodes.Add("RUSSIAN", "ru")
        langcodes.Add("SAMI_INARI", "smn")
        langcodes.Add("SAMI_LULE", "smj")
        langcodes.Add("SAMI_NORTHERN", "se")
        langcodes.Add("SAMI_SKOLT", "sms")
        langcodes.Add("SAMI_SOUTHERN", "sma")
        langcodes.Add("SANSKRIT", "sa")
        langcodes.Add("GAELIC", "ga")
        langcodes.Add("SERBIAN", "sr")
        langcodes.Add("SESOTHO", "nso")
        langcodes.Add("SETSWANA", "tn")
        langcodes.Add("SINHALA", "si")
        langcodes.Add("SINHALESE", "si")
        langcodes.Add("SLOVAK", "sk")
        langcodes.Add("SLOVENIAN", "sl")
        langcodes.Add("SPANISH", "es")
        langcodes.Add("SWAHILI", "sw")
        langcodes.Add("SWEDISH", "sv")
        langcodes.Add("SYRIAC", "syr")
        langcodes.Add("TAJIK", "tg")
        langcodes.Add("TAMAZIGHT", "tzm")
        langcodes.Add("TAMIL", "ta")
        langcodes.Add("TATAR", "tt")
        langcodes.Add("TELUGU", "te")
        langcodes.Add("THAI", "th")
        langcodes.Add("TIBETAN", "bo")
        langcodes.Add("TURKISH", "tr")
        langcodes.Add("TURKMEN", "tk")
        langcodes.Add("UIGHUR", "ug")
        langcodes.Add("UKRAINIAN", "uk")
        langcodes.Add("UPPER_SORBIAN", "wen")
        langcodes.Add("URDU", "ur")
        langcodes.Add("UZBEK", "uz")
        langcodes.Add("VIETNAMESE", "vi")
        langcodes.Add("WELSH", "cy")
        langcodes.Add("WOLOF", "wo")
        langcodes.Add("XHOSA", "xh")
        langcodes.Add("YAKUT", "sah")
        langcodes.Add("YI", "ii")
        langcodes.Add("YORUBA", "yo")
        langcodes.Add("ZULU", "zu")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "finishing BuildLangCodes Sub...")
    End Sub

    Private Sub prepareDynamicInformation()
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Entering prepareDynamicInformation Sub...")
        langsLocalized.Clear()
        ListView1.Clear()
        If bibleGetDB Is Nothing Then bibleGetDB = New BibleGetDatabase
        If bibleGetDB.IsInitialized Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "BibleGetDatabase is correctly initialized...")
            Using conn As New SQLiteConnection(bibleGetDB.connectionStr)
                If conn IsNot Nothing Then
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Connection to BibleGet sqlite database succeeded")
                    conn.Open()
                    Using sqlQuery As New SQLiteCommand(conn)
                        Dim queryString As String = "SELECT VERSIONS FROM METADATA WHERE ID=0"
                        Dim queryString2 As String = "SELECT LANGUAGES FROM METADATA WHERE ID=0"
                        sqlQuery.CommandText = queryString
                        Dim versionsString As String = sqlQuery.ExecuteScalar()
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "versionsString = " & versionsString)
                        sqlQuery.CommandText = queryString2
                        Dim langsSupported As String = sqlQuery.ExecuteScalar
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "langsSupported = " & langsSupported)

                        Dim versionsObj As JObject = JObject.Parse(versionsString)
                        Dim keys() As String = versionsObj.Properties().Select(Function(p) p.Name).ToArray()
                        versionCount = keys.Length
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "versionCount = " & versionCount)
                        Dim BibleVersions As New ArrayList()

                        Dim lvGroups As New Dictionary(Of String, ListViewGroup)

                        For Each s As String In keys
                            Dim versionStr As String = versionsObj.SelectToken(s).Value(Of String)
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
                            End Try
                            languageName = fullLanguageName.ToUpper(CultureInfo.CurrentUICulture)
                            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "lang code <" & strArray(2) & "> expanded to localized language name as: " & languageName)
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
                        Next
                        ListView1.View = View.Details
                        colHeader = New ColumnHeader()
                        colHeader.Text = String.Empty
                        colHeader.Width = -2
                        colHeader.TextAlign = HorizontalAlignment.Left
                        ListView1.Columns.Add(colHeader)
                        ListView1.HeaderStyle = ColumnHeaderStyle.None
                        ListView1.Columns(0).Width = ListView1.Width - 4 - SystemInformation.VerticalScrollBarWidth

                        Dim langsObj As JArray = JArray.Parse(langsSupported)
                        booksLangs = langsObj.Count
                        For Each jsonValue As JValue In langsObj
                            langsLocalized.Add(localizeLanguage(jsonValue.ToString))
                        Next
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & String.Join(",", langsLocalized))
                    End Using
                Else
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "we seem to have a null connection... arghhh!")
                End If
            End Using
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor
        If bibleGetDB.UpdateMetaDataTable() Then
            prepareDynamicInformation()
            BibleGetAddIn.checkForUpdate()
            CurrentInfo.Text = String.Format(__("The BibleGet database currently supports {0} versions of the Bible in {1} different languages:"), versionCount, versionLangs)
            ServerDataLangsCount.Text = String.Format(__("The BibleGet engine currently understands the names of the books of the Bible in {0} different languages:"), booksLangs)
            ServerDataLangs.Text = String.Join(", ", langsLocalized)
            If My.Settings.NewVersionExists Then
                LabelVersion.Text = __("Version") & " " & My.Application.Info.Version.ToString & " (there is a newer version " & My.Settings.NewVersion & ", click here to update)"
                LabelVersion.ForeColor = Drawing.Color.DarkBlue
                LabelVersion.BackColor = Drawing.Color.Yellow
                LabelVersion.Cursor = Cursors.Hand
            Else
                LabelVersion.Text = __("Version") & " " & My.Application.Info.Version.ToString
                LabelVersion.ForeColor = Drawing.Color.Black
                LabelVersion.BackColor = Drawing.Color.Transparent
                LabelVersion.Cursor = Cursors.Default
            End If
            MsgBox("Data was correctly updated from the BibleGet server.", MsgBoxStyle.Information)
        Else
            MsgBox("Error renewing data from server. Please try again later.", MsgBoxStyle.Critical)
        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub OKButton_Click_1(sender As Object, e As EventArgs) Handles OKButton.Click
        Close()
    End Sub

    Private Sub LabelVersion_Click(sender As Object, e As EventArgs) Handles LabelVersion.Click
        If My.Settings.NewVersionExists Then
            DoVersionUpdate()
        End If
    End Sub

    Private Sub LabelVersion_MouseEnter(sender As Object, e As EventArgs) Handles LabelVersion.MouseEnter
        If My.Settings.NewVersionExists Then
            LabelVersion.ForeColor = Drawing.Color.Blue
            LabelVersion.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    Private Sub LabelVersion_MouseLeave(sender As Object, e As EventArgs) Handles LabelVersion.MouseLeave
        If My.Settings.NewVersionExists Then
            LabelVersion.ForeColor = Drawing.Color.DarkBlue
            LabelVersion.BackColor = Drawing.Color.LightYellow
        End If
    End Sub

    Private Sub OnDownloadComplete(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
        If InvokeRequired Then
            BeginInvoke(New Action(Of AsyncCompletedEventArgs)(AddressOf DoDownloadCompleted), e)
        Else
            DoDownloadCompleted(e)
        End If
    End Sub

    Private Sub UpdateDownloadProgress(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
        If InvokeRequired Then
            BeginInvoke(New Action(Of DownloadProgressChangedEventArgs)(AddressOf UpdateProgressBar), e)
        Else
            UpdateProgressBar(e)
        End If
    End Sub

    Private Sub UpdateProgressBar(ByVal e As DownloadProgressChangedEventArgs)
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub UpdateProgressBar(ByVal e As ProgressChangedEventArgs)
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub DoDownloadCompleted(ByVal e As AsyncCompletedEventArgs)
        If Not e.Cancelled AndAlso e.Error Is Nothing Then
            'MessageBox.Show("Download success")
            LabelVersion.Cursor = Cursors.Default
            Try
                updateProcess.StartInfo.FileName = localFile
                updateProcess.EnableRaisingEvents = True
                updateProcess.Start()
            Catch ex As Exception
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
            End Try
        Else
            MessageBox.Show("Download failed")
        End If
    End Sub

    ' Handle Exited event and display process information.
    Private Sub updateProcess_Exited(ByVal sender As Object,
            ByVal e As System.EventArgs) Handles updateProcess.Exited
        elapsedTime = updateProcess.ExitTime - updateProcess.StartTime
        'eventHandled = True
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Start time:    " & updateProcess.StartTime)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Exit time:    " & updateProcess.ExitTime)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Exit code:    " & updateProcess.ExitCode)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Elapsed time: " & elapsedTime.TotalSeconds)

        If updateProcess.ExitCode = 0 Then
            My.Settings.NewVersionExists = False
            Dim restart As Windows.Forms.DialogResult
            restart = MessageBox.Show("BibleGet Plugin for MSWord 2007+ updated successfully. Restart now? (Current document will be saved.)", "Update Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
            If restart = Windows.Forms.DialogResult.OK Then
                Dim oWord As Word.Application
                oWord = CreateObject("Word.Application")
                oWord.Visible = True
                oWord.Documents.Add()
                WordApplication.Quit(Word.WdSaveOptions.wdPromptToSaveChanges, Word.WdOriginalFormat.wdPromptUser)
            Else
                MessageBox.Show("Word needs to be restarted in order to load the updated plugin. Please restart.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            MessageBox.Show("Plugin installer did not complete. Plugin was not updated.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub ListView1_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        e.Item.Selected = False
    End Sub

    Private Sub UpdateCheckBtn_Click(sender As Object, e As EventArgs) Handles UpdateCheckBtn.Click
        Cursor = Cursors.WaitCursor
        If BackgroundWorker1.IsBusy <> True Then
            ' Start the asynchronous operation.
            ProgressBar1.Visible = True
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Debug.WriteLine("Starting background work...")
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        worker.ReportProgress(10)
        Debug.WriteLine("now doing online check...")
        Dim onlineVersion As Version = HTTPCaller.GetCurrentVersion
        worker.ReportProgress(50)
        My.Settings.NewVersion = onlineVersion.ToString
        Debug.WriteLine("result of online check: online version = " & onlineVersion.ToString)
        If Version.op_GreaterThan(onlineVersion, My.Application.Info.Version) Then
            'Console.WriteLine("Detected online version is greater than current version")
            Debug.WriteLine("Detected online version is greater than current version")
            My.Settings.NewVersionExists = True
        ElseIf Version.op_LessThan(onlineVersion, My.Application.Info.Version) Then
            Debug.WriteLine("Detected online version is less than current version, it seems you are using an unreleased beta version?")
            My.Settings.NewVersionExists = False
        Else
            Debug.WriteLine("Detected online version is the same as the current version")
            My.Settings.NewVersionExists = False
        End If
        My.Settings.UpdateCheck = DateTime.Now
        My.Settings.Save()
        worker.ReportProgress(100)
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If InvokeRequired Then
            BeginInvoke(New Action(Of ProgressChangedEventArgs)(AddressOf UpdateProgressBar), e)
        Else
            UpdateProgressBar(e)
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If InvokeRequired Then
            BeginInvoke(New Action(Of AsyncCompletedEventArgs)(AddressOf DoUpdateCheckCompleted), e)
        Else
            DoUpdateCheckCompleted(e)
        End If
    End Sub

    Private Sub DoVersionUpdate()
        Dim remoteUri As New Uri("https://sourceforge.net/projects/bibleget/files/latest/download")
        localFile = Path.GetTempPath & "BibleGetIOMSWordAddInSetup_" & My.Settings.NewVersion.Replace(".", "") & ".exe"
        If File.Exists(localFile) Then
            Try
                updateProcess.StartInfo.FileName = localFile
                updateProcess.EnableRaisingEvents = True
                updateProcess.Start()
            Catch ex As Exception
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
                MessageBox.Show("Plugin installer was interrupted. Plugin was not updated.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            LabelVersion.Cursor = Cursors.WaitCursor
            ProgressBar1.Visible = True
            Dim webClient As New WebClient
            AddHandler webClient.DownloadProgressChanged, AddressOf UpdateDownloadProgress
            AddHandler webClient.DownloadFileCompleted, AddressOf OnDownloadComplete
            webClient.DownloadFileAsync(remoteUri, localFile)
        End If
    End Sub

    Private Sub DoUpdateCheckCompleted(ByVal e As AsyncCompletedEventArgs)
        If Not e.Cancelled AndAlso e.Error Is Nothing Then
            If My.Settings.NewVersionExists Then
                Dim updateChoice As DialogResult = MessageBox.Show("Version " & My.Settings.NewVersion & " is available online. Would you like to update to the new version now?", "Update available", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Select Case updateChoice
                    Case DialogResult.Yes

                    Case DialogResult.No
                        'nothing to do here
                End Select
            Else
                MessageBox.Show("You have the latest version of the BibleGet add-on for Microsoft Word.")
            End If
            ProgressBar1.Visible = False
            Cursor = Cursors.Default
        End If
    End Sub
End Class
