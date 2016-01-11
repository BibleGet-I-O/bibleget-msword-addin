Imports System.Data.SQLite
Imports System.Data
Imports Newtonsoft.Json.Linq
Imports System.Collections
Imports System.Globalization
Imports System.ComponentModel
Imports System.Windows.Forms


Public NotInheritable Class AboutBibleGet

    Private Function __(ByVal myStr As String) As String
        Dim myTranslation As String = ThisAddIn.RM.GetString(myStr, ThisAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Imposta il titolo del form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        'Me.Text = String.Format("Informazioni su {0}", ApplicationTitle)
        Me.Text = __("About this plugin")
        ' Inizializza tutto il testo visualizzato nella finestra di dialogo Informazioni su.
        Me.LabelProductName.Text = __(My.Application.Info.ProductName)
        Me.LabelVersion.Text = __("Version") & " " & My.Application.Info.Version.ToString
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        'Me.TextBoxDescription.Text = __(My.Application.Info.Description)
        Dim descr As String
        descr = __("This plugin was developed by <b>John R. D'Orazio</b>, a priest in the diocese of Rome, chaplain at Roma Tre University.") _
                + " " _
                + String.Format(__("It is a part of the <b>BibleGet Project</b> at {0}."), "<span style='color:Blue;'>http://www.bibleget.io</span>") _
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


        Dim versionCount As Integer
        Dim versionLangs As Integer
        Dim bibleGetDB As New BibleGetDatabase
        Dim conn As SQLiteConnection
        If bibleGetDB.INITIALIZED Then
            conn = bibleGetDB.connect()
            If conn IsNot Nothing Then
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
                        Next
                        ListView1.View = View.Details
                        Dim colHeader As ColumnHeader = New ColumnHeader()
                        colHeader.Text = "Available Bible Versions"
                        colHeader.Width = -2
                        colHeader.TextAlign = HorizontalAlignment.Left
                        ListView1.Columns.Add(colHeader)
                        ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
                        ListView1.Columns(0).Width = ListView1.Width - 4 - SystemInformation.VerticalScrollBarWidth
                        ListView1.Enabled = False
                    End Using
                End Using
            Else
                Diagnostics.Debug.WriteLine("we seem to have a null connection... arghhh!")
            End If
        End If

        CurrentInfo.Text = String.Format(__("The BibleGet database currently supports {0} versions of the Bible in {1} different languages:"), versionCount, versionLangs)

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

End Class
