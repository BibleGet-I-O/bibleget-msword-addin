Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Data.SQLite
Imports Newtonsoft.Json.Linq
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class BibleGetHelp

    Private HtmlStr0 As String
    Private HtmlStr1 As String
    Private HtmlStr2 As String
    Private HtmlStr3 As String
    Private HtmlStr3Table As String
    Private HtmlStr3Closing As String
    Private stylesheet As String
    Private packagepath As String
    Private lastNode As Windows.Forms.TreeNode
    Private booksLangs As Integer
    Private booksStr As String
    Private langsObj As JArray
    Private langcodes As New Dictionary(Of String, String)
    Private langsLocalized As List(Of String) = New List(Of String)
    Private curLang As String
    Private booksAndAbbreviations As New Dictionary(Of String, String)
    Private nodeFont As New System.Drawing.Font("Garamond", 12, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
    Private DEBUG_MODE As Boolean

    Private Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = BibleGetAddIn.RM.GetString(myStr, BibleGetAddIn.locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Dim rgx As New Regex("''")
            myTranslation = rgx.Replace(myTranslation, "'")
            Return myTranslation
        Else
            Return myStr
        End If
    End Function


    'turn off the annoying clicking sound when the preview window refreshes (WebBrowser control)
    Const DS As Integer = 21
    Const SP As Integer = &H2

    Private Sub BibleGetHelp_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        nodeFont.Dispose()
    End Sub


    Private Sub BibleGetHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DEBUG_MODE = My.Settings.DEBUG_MODE
        NativeMethods.CoInternetSetFeatureEnabled(DS, SP, True) 'Dim clickOff As Boolean = 
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Entering BibleGetHelp load event")
        Text = __("Instructions")
        TreeView1.Nodes.Clear()

        packagepath = "data:image/png;base64,"

        Dim rootNode As TreeNode = New TreeNode(__("Help"))
        rootNode.NodeFont = New Font("Garamond", 14, FontStyle.Regular, GraphicsUnit.Point)

        Dim usageNode As TreeNode = New TreeNode(__("Usage of the Plugin"))
        usageNode.NodeFont = nodeFont
        Dim formulationNode As TreeNode = New TreeNode(__("Formulation of the Queries"))
        formulationNode.NodeFont = nodeFont
        Dim booksNode As TreeNode = New TreeNode(__("Biblical Books and Abbreviations"))
        booksNode.NodeFont = nodeFont

        TreeView1.Nodes.Add(rootNode)
        rootNode.Nodes.Add(usageNode)
        rootNode.Nodes.Add(formulationNode)
        rootNode.Nodes.Add(booksNode)

        rootNode.Checked = True
        rootNode.Expand()

        BuildLangCodes()
        RetrieveInfoFromDB(booksNode)

        WebBrowser1.ObjectForScripting = True


        stylesheet = "body { background-color: #FFFFDD; border: 2px inset #CC9900; font-size: 10pt; }" & Environment.NewLine
        stylesheet &= "h1 { color: #0000AA; }" & Environment.NewLine
        stylesheet &= "h2 { color: #0000AA; }" & Environment.NewLine
        stylesheet &= "h3 { color: #0000AA; }" & Environment.NewLine
        stylesheet &= "p { text-align: justify; }" & Environment.NewLine
        stylesheet &= "div#tablecontainer { text-align: center; }" & Environment.NewLine
        stylesheet &= "table { border-collapse: collapse; width: 400px; margin: 10px auto; }" & Environment.NewLine
        stylesheet &= "th { text-align: center; border: 4px ridge #DEB887; background-color: #F5F5DC; padding: 3px; }" & Environment.NewLine
        stylesheet &= "td { text-align: justify; border: 3px ridge #DEB887; background-color: #F5F5DC; padding: 3px; }" & Environment.NewLine


        'TODO: Populate children of booksNode with language variants
        HtmlStr0 = "<html><head><meta charset=""utf-8""><style type=""text/css"">"
        HtmlStr0 &= stylesheet
        HtmlStr0 &= "</style></head><body>"
        HtmlStr0 &= "<h2>" + __("Help for BibleGet (Open Office Writer)") + "</h2>"
        HtmlStr0 &= "<p>" + __("This Help dialog window introduces the user to the usage of the BibleGet I/O plugin for Open Office Writer.") + "</p>"
        HtmlStr0 &= "<p>" + __("The Help is divided into three sections:") + "</p>"
        HtmlStr0 &= "<ul>"
        HtmlStr0 &= "<li>" + __("Usage of the Plugin") + "</li>"
        HtmlStr0 &= "<li>" + __("Formulation of the Queries") + "</li>"
        HtmlStr0 &= "<li>" + __("Biblical Books and Abbreviations") + "</li>"
        HtmlStr0 &= "</ul>"
        HtmlStr0 &= "<p><b>" + __("AUTHOR") + ":</b> " + __("John R. D'Orazio (priest in the Diocese of Rome)") + "</p>"
        HtmlStr0 &= "<p><b>" + __("COLLABORATORS") + ":</b> " + __("Giovanni Gregori (computing) and Simone Urbinati (MUG Roma Tre)") + "</p>"
        HtmlStr0 &= "<p><b>" + __("Version").ToUpper(CultureInfo.CurrentCulture) + ":</b> " & My.Application.Info.Version.ToString + "</p>"
        HtmlStr0 &= "<p>© <b>Copyright 2016 BibleGet I/O by John R. D'Orazio</b> <a href=""mailto:john.dorazio@cappellaniauniroma3.org"">john.dorazio@cappellaniauniroma3.org</a></p>"
        HtmlStr0 &= "<p><b>" + __("PROJECT WEBSITE") + ": </b><a href=""http://www.bibleget.io"">http://www.bibleget.io</a><br>"
        HtmlStr0 &= "<b>" + __("EMAIL ADDRESS FOR INFORMATION OR FEEDBACK ON THE PROJECT") + ":</b> <a href=""mailto:bibleget.io@gmail.com"">bibleget.io@gmail.com</a></p>"
        HtmlStr0 &= "<p>Cappellania Università degli Studi Roma Tre - Piazzale San Paolo 1/E - 00120 Città del Vaticano - +39 06.69.88.08.09 - <a href=""mailto:cappellania.uniroma3@gmail.com"">cappellania.uniroma3@gmail.com</a></p></body></html>"

        Dim strfmt1 As String = __("Insert quote from input window")
        Dim strfmt2 As String = __("About this plugin")
        Dim strfmt3 As String = __("RENEW SERVER DATA")
        Dim strfmt4 As String = strfmt1
        Dim strfmt5 As String = __("Insert quote from text selection")
        Dim strfmt6 As String = strfmt1

        Dim emailBase64 As String = ImageToBase64(My.Resources.email)
        Dim paypalBase64 As String = ImageToBase64(My.Resources.paypal)
        Dim infoBase64 As String = ImageToBase64(My.Resources.info)
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "x_wrong_mark format is: " + GetMimeType(My.Resources.red_x_wrong_mark))

        Dim screenshotPath As String = IO.Path.Combine(IO.Path.GetTempPath, "screenshot.png")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "screenshotPath = " + screenshotPath)
        If Not File.Exists(screenshotPath) Then
            My.Resources.screenshot_ribbon.Save(screenshotPath)
        End If

        Dim screenshotPath1 As String = IO.Path.Combine(IO.Path.GetTempPath, "screenshot1.png")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "screenshotPath1 = " + screenshotPath1)
        If Not File.Exists(screenshotPath1) Then
            My.Resources.screenshot_input_window.Save(screenshotPath1)
        End If

        Dim screenshotPath2 As String = IO.Path.Combine(IO.Path.GetTempPath, "screenshot2.png")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "screenshotPath2 = " + screenshotPath2)
        If Not File.Exists(screenshotPath2) Then
            My.Resources.screenshot_text_selection.Save(screenshotPath2)
        End If

        Dim screenshotPath3 As String = IO.Path.Combine(IO.Path.GetTempPath, "screenshot3.png")
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "screenshotPath3 = " + screenshotPath3)
        If Not File.Exists(screenshotPath3) Then
            My.Resources.screenshot_user_preferences.Save(screenshotPath3)
        End If

        'Dim ms As MemoryStream = New MemoryStream()
        'My.Resources.red_x_wrong_mark.Save(ms, My.Resources.red_x_wrong_mark.RawFormat)
        'Dim imageBytes() As Byte = ms.ToArray
        'IO.File.WriteAllBytes(screenshotPath, imageBytes)

        HtmlStr1 = "<html><head><meta charset=""utf-8""><style type=""text/css"">"
        HtmlStr1 &= stylesheet
        HtmlStr1 &= "</style></head><body>"
        HtmlStr1 &= "<h2>" + __("How to use the plugin") + "</h2>"
        HtmlStr1 &= "<h3>" + __("Description of the menu icons and their functionality.") + "</h3>"
        HtmlStr1 &= "<p>" + __("Once the extension is installed, a new menu 'BibleGet I/O' will appear. Here is a screenshot with the buttons on the ribbon area of the new menu item:") + "</p><br /><br />"
        HtmlStr1 &= "<img src=""" + screenshotPath + """ style=""width:95%;border:1px solid Gray;"" alt=""Screenshot.jpg"" /><br /><br />"
        HtmlStr1 &= "<p>" + __("There are two ways of inserting a bible quote into a document.")
        HtmlStr1 &= " "
        HtmlStr1 &= __("The first way is by using the input window.")
        HtmlStr1 &= " "
        HtmlStr1 &= String.Format(__("If you click on the menu item ''{0}'', an input window will open where you can input your query and choose the version or versions you would like to take the quote from."), strfmt1)
        HtmlStr1 &= " "
        HtmlStr1 &= __("This list of versions is updated from the available versions on the BibleGet server, but since the information is stored locally it may be necessary to renew the server information when new versions are added to the BibleGet server database.")
        HtmlStr1 &= " "
        HtmlStr1 &= String.Format(__("In order to renew the information from the BibleGet server, click on the ''{0}'' menu item, and then click on the button ''{1}''."), strfmt2, strfmt3)
        HtmlStr1 &= " "
        HtmlStr1 &= String.Format(__("When you choose a version or multiple versions to quote from, this choice is automatically saved as a preference, and will be pre-selected the next time you open the ''{0}'' menu item."), strfmt4)
        HtmlStr1 &= "<br /><br />"
        HtmlStr1 &= "<img src=""" + screenshotPath1 + """ alt=""Screenshot1.jpg"" /><br /><br />"
        HtmlStr1 &= String.Format(__("The second way is by writing your desired quote directly in the document, and then selecting it and choosing the menu item ''{0}''. The selected text will be substituted by the Bible Quote retrieved from the BibleGet server."), strfmt5)
        HtmlStr1 &= " "
        HtmlStr1 &= String.Format(__("The versions previously selected in the ''{0}'' window will be used, so you must have selected your preferred versions at least once from the ''{0}'' window."), strfmt6)
        HtmlStr1 &= "</p><br /><br />"
        HtmlStr1 &= "<img src=""" + screenshotPath2 + """ alt=""Screenshot2.jpg"" /><br /><br />"
        HtmlStr1 &= "<p>"
        HtmlStr1 &= __("Formatting preferences can be set using the 'Preferences' window. You can choose the desired font for the Bible quotes as well as the desired line-spacing, and you can choose separate formatting (font size, font color, font style) for the book / chapter, for the verse numbers, and for the verse text. Preferences are saved automatically.")
        HtmlStr1 &= "</p><br /><br />"
        HtmlStr1 &= "<img src=""" + screenshotPath3 + """ alt=""Screenshot3.jpg"" /><br /><br />"
        HtmlStr1 &= "<p>"
        HtmlStr1 &= __("After the 'Help' menu item that opens up this same help window, the last three menu items are:")
        HtmlStr1 &= "</p>"
        HtmlStr1 &= "<ul>"
        HtmlStr1 &= "<li><img src=""" + packagepath + emailBase64 + """ alt=""email.png"" />"
        HtmlStr1 &= " '"
        HtmlStr1 &= __("Send feedback")
        HtmlStr1 &= "': <span>"
        HtmlStr1 &= __("This will open up your system's default email application with the bibleget.io@gmail.com feedback address already filled in.")
        HtmlStr1 &= "</span></li>"
        HtmlStr1 &= "<li><img src=""" + packagepath + paypalBase64 + """ alt=""paypal.png"" />"
        HtmlStr1 &= " '"
        HtmlStr1 &= __("Contribute")
        HtmlStr1 &= "': <span>"
        HtmlStr1 &= __("This will open a Paypal page in the system's default browser where you can make a donation to contribute to the project. Even just €1 can help to cover the expenses of this project. Just the server costs €120 a year.")
        HtmlStr1 &= "</span></li>"
        HtmlStr1 &= "<li><img src=""" + packagepath + infoBase64 + """ alt=""info.png"" />"
        HtmlStr1 &= " '"
        HtmlStr1 &= __("Information on the BibleGet I/O Project")
        HtmlStr1 &= "': <span>"
        HtmlStr1 &= __("This opens a dialog window with some information on the project and it's plugins, on the author and contributors, and on the current locally stored information about the versions and languages that the BibleGet server supports.")
        HtmlStr1 &= "</span></li>"
        HtmlStr1 &= "</ul>"
        HtmlStr1 &= "</body></html>"

        Dim strfmt7 As String = __("Biblical Books and Abbreviations")

        HtmlStr2 = "<html>"
        HtmlStr2 &= "<head><meta charset=""utf-8""><style type=""text/css"">"
        HtmlStr2 &= stylesheet
        HtmlStr2 &= "</style></head>"
        HtmlStr2 &= "<body>"
        HtmlStr2 &= "<h2>" + __("How to formulate a bible query") + "</h2>"
        HtmlStr2 &= "<p>"
        HtmlStr2 &= __("The queries for bible quotes must be formulated using standard notation for bible citation.")
        HtmlStr2 &= " "
        HtmlStr2 &= __("This can be either the english notation (as explained here: https://en.wikipedia.org/wiki/Bible_citation), or the european notation as explained here below.")
        HtmlStr2 &= "</p>"
        HtmlStr2 &= "<p>"
        HtmlStr2 &= __("A basic query consists of at least two elements: the bible book and the chapter.")
        HtmlStr2 &= " "
        HtmlStr2 &= __("The bible book can be written out in full, or in an abbreviated form.")
        HtmlStr2 &= " "
        HtmlStr2 &= String.Format(__("The BibleGet engine recognizes the names of the books of the bible in {0} different languages: {1}"), booksLangs, booksStr) & ". "
        HtmlStr2 &= " "
        HtmlStr2 &= String.Format(__("See the list of valid books and abbreviations in the section {0}."), "<span class=""internal-link"" id=""to-bookabbrevs"">" + strfmt7 + "</span>")
        HtmlStr2 &= " "
        HtmlStr2 &= __("For example, the query ""Matthew 1"" means the book of Matthew (or better the gospel according to Matthew) at chapter 1.")
        HtmlStr2 &= " "
        HtmlStr2 &= __("This can also be written as ""Mt 1"".")
        HtmlStr2 &= "</p>"
        HtmlStr2 &= "<p>" + __("Different combinations of books, chapters, and verses can be formed using the comma delimiter and the dot delimiter (in european notation, in english notation instead a colon is used instead of a comma and a comma is used instead of a dot):") + "</p>"
        HtmlStr2 &= "<ul>"
        HtmlStr2 &= "<li>" + __(""","": the comma is the chapter-verse delimiter. ""Matthew 1,5"" means the book (gospel) of Matthew, chapter 1, verse 5. (In English notation: ""Matthew 1:5"".)") + "</li>"
        HtmlStr2 &= "<li>" + __("""."": the dot is a delimiter between verses. ""Matthew 1,5.7"" means the book (gospel) of Matthew, chapter 1, verses 5 and 7. (In English notation: ""Matthew 1:5,7"".)") + "</li>"
        HtmlStr2 &= "<li>" + __("""-"": the dash is a range delimiter, which can be used in a variety of ways:")
        HtmlStr2 &= "<ol>"
        HtmlStr2 &= "<li>" + __("For a range of chapters: ""Matthew 1-2"" means the gospel according to Matthew, from chapter 1 to chapter 2.") + "</li>"
        HtmlStr2 &= "<li>" + __("For a range of verses within the same chapter: ""Matthew 1,1-5"" means the gospel according to Matthew, chapter 1, from verse 1 to verse 5. (In English notation: ""Matthew 1:1-5"".)") + "</li>"
        HtmlStr2 &= "<li>" + __("For a range of verses that span over different chapters: ""Matthew 1,5-2,13"" means the gospel according to Matthew, from chapter 1, verse 5 to chapter 2, verse 13. (In English notation: ""Matthew 1:5-2:13"".)") + "</li>"
        HtmlStr2 &= "</ol>"
        HtmlStr2 &= "</ul>"
        HtmlStr2 &= "<p>" + __("Different combinations of these delimiters can form fairly complex queries, for example ""Mt1,1-3.5.7-9"" means the gospel according to Matthew, chapter 1, verses 1 to 3, verse 5, and verses 7 to 9. (In English notation: ""Mt1:1-3,5,7-9"".)") + "</p>"
        HtmlStr2 &= "<p>" + __("Multiple queries can be combined together using a semi-colon "";"".")
        HtmlStr2 &= " "
        HtmlStr2 &= __("If the query following the semi-colon refers to the same book as the preceding query, it is not necessary to indicate the book a second time.")
        HtmlStr2 &= " "
        HtmlStr2 &= __("For example, ""Matthew 1,1;2,13"" means the gospel according to Matthew, chapter 1 verse 1 and chapter 2 verse 13. (In English notation: ""Matthew 1:1;2:13"".)")
        HtmlStr2 &= " "
        HtmlStr2 &= __("Here is an example of multiple complex queries combined into a single querystring: ""Genesis 1,3-5.7.9-11.13;2,4-9.11-13;Apocalypse 3,10.12-14"". (In English notation: ""Genesis 1:3-5,7,9-11,13;2:4-9,11-13;Apocalypse 3:10,12-14"").") + "</p>"
        HtmlStr2 &= "<p>" + __("It doesn't matter whether or not you use a space between the book and the chapter, the querystring will be interpreted just the same.")
        HtmlStr2 &= __("It is also indifferent whether you use uppercase or lowercase letters, the querystring will be interpreted just the same.")
        HtmlStr2 &= "</p>"
        HtmlStr2 &= "</body>"
        HtmlStr2 &= "</html>"


        HtmlStr3 = "<html>"
        HtmlStr3 &= "<head><meta charset=""utf-8""><style type=""text/css"">"
        HtmlStr3 &= stylesheet
        HtmlStr3 &= "</style></head>"
        HtmlStr3 &= "<body>"
        HtmlStr3 &= "<h2>" + __("Biblical Books and Abbreviations") + "</h2>"
        HtmlStr3 &= "<p>" + __("Here is a list of valid books and their corresponding abbreviations, either of which can be used in the querystrings.")
        HtmlStr3 &= " "
        HtmlStr3 &= __("The abbreviations do not always correspond with those proposed by the various editions of the Bible, because they would conflict with those proposed by other editions.")
        HtmlStr3 &= " "
        HtmlStr3 &= __("For example some english editions propose ""Gn"" as an abbreviation for ""Genesis"", while some italian editions propose ""Gn"" as an abbreviation for ""Giona"" (= ""Jonah"").")
        HtmlStr3 &= " "
        HtmlStr3 &= __("Therefore you will not always be able to use the abbreviations proposed by any single edition of the Bible, you must use the abbreviations that are recognized by the BibleGet engine as listed in the following table:")
        HtmlStr3 &= "</p><br /><br />"

        HtmlStr3Table = "<div id=""tablecontainer""><table cellspacing='0'>"
        HtmlStr3Table &= "<caption>{0}</caption>"
        HtmlStr3Table &= "<tr><th style=""width:70%;"">" + __("BOOK") + "</th><th style=""width:30%;"">" + __("ABBREVIATION") + "</th></tr>"
        HtmlStr3Table &= "{1}"
        HtmlStr3Table &= "</table></div>"

        HtmlStr3Closing = "</body>"
        HtmlStr3Closing &= "</html>"

        SetPreviewDocument(__("Help"))
        'End Using
    End Sub


    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        SetPreviewDocument(e.Node.Text)
    End Sub

    Private Sub SetPreviewDocument(ByVal node As String)
        Dim previewDocument As String = String.Empty
        Select Case node
            Case __("Help")
                previewDocument = HtmlStr0
            Case __("Usage of the Plugin")
                previewDocument = HtmlStr1
            Case __("Formulation of the Queries")
                previewDocument = HtmlStr2
            Case __("Biblical Books and Abbreviations")
                Dim curLangIsoCode As String = BibleGetAddIn.locale.TwoLetterISOLanguageName
                Dim curLangDisplayName As String = New CultureInfo(curLangIsoCode).DisplayName
                curLang = localizeLanguage(curLangDisplayName).ToUpper(CultureInfo.CurrentCulture)
                previewDocument = HtmlStr3
                previewDocument &= String.Format(HtmlStr3Table, curLang, booksAndAbbreviations.Item(curLang))
                previewDocument &= HtmlStr3Closing
            Case Else
                If langsLocalized.Contains(node) Then
                    curLang = node
                    previewDocument = HtmlStr3
                    previewDocument &= String.Format(HtmlStr3Table, curLang, booksAndAbbreviations.Item(curLang))
                    previewDocument &= HtmlStr3Closing
                End If
        End Select

        If WebBrowser1.Document Is Nothing Then
            WebBrowser1.DocumentText = previewDocument
        Else
            WebBrowser1.Document.Write(String.Empty)
            WebBrowser1.Document.Write(previewDocument)
            WebBrowser1.Refresh()
        End If

    End Sub

    Private Sub TreeView1_MouseMove(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseMove

        If TreeView1.HitTest(e.Location).Node IsNot Nothing Then
            Dim nde As Windows.Forms.TreeNode = TreeView1.HitTest(e.Location).Node
            If nde IsNot lastNode Then
                TreeView1.BeginUpdate()
                If lastNode IsNot Nothing Then lastNode.BackColor = Drawing.Color.Empty
                nde.BackColor = Drawing.Color.Yellow
                TreeView1.EndUpdate()
                lastNode = nde
            End If
        End If

    End Sub

    Private Shared Function ImageToBase64(ByVal image As Drawing.Image) As String
        Using ms As MemoryStream = New MemoryStream()
            image.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            Dim imageBytes() As Byte = ms.ToArray
            Return Convert.ToBase64String(imageBytes)
        End Using
    End Function

    Private Sub RetrieveInfoFromDB(ByVal bNode As TreeNode)
        Dim bibleGetDB As New BibleGetDatabase
        If bibleGetDB.IsInitialized Then
            Using conn As New SQLiteConnection(bibleGetDB.connectionStr)
                If conn IsNot Nothing Then
                    conn.Open()
                    Using sqlQuery As New SQLiteCommand(conn)
                        'Dim queryString As String = "SELECT VERSIONS FROM METADATA WHERE ID=0"
                        Dim queryString2 As String = "SELECT LANGUAGES FROM METADATA WHERE ID=0"
                        'sqlQuery.CommandText = queryString
                        'Dim versionsString As String = sqlQuery.ExecuteScalar()
                        sqlQuery.CommandText = queryString2
                        Dim langsSupported As String = sqlQuery.ExecuteScalar

                        langsObj = JArray.Parse(langsSupported)
                        booksLangs = langsObj.Count
                        For Each jsonValue As JValue In langsObj
                            langsLocalized.Add(localizeLanguage(jsonValue.ToString(CultureInfo.CurrentCulture)))
                        Next
                        langsLocalized.Sort()
                        booksStr = String.Join(", ", langsLocalized)
                        For Each title In langsLocalized
                            Dim treeNode As New TreeNode(title)
                            treeNode.NodeFont = New Drawing.Font("Garamond", 10, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
                            bNode.Nodes.Add(treeNode)
                        Next
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & String.Join(",", langsLocalized))

                        Dim bibleBooksTemp As New List(Of JArray)
                        Dim bbBooks As String
                        For i As Integer = 0 To 72
                            sqlQuery.CommandText = "SELECT BIBLEBOOKS" & i.ToString(CultureInfo.InvariantCulture) & " FROM METADATA WHERE ID=0"
                            bbBooks = sqlQuery.ExecuteScalar
                            Dim bibleBooksObj As JArray = JArray.Parse(bbBooks)
                            bibleBooksTemp.Add(bibleBooksObj)
                        Next

                        'Dim booksForCurLang As New List(Of String())
                        'booksAndAbbreviations = New Dictionary(Of String, String)
                        Dim buildStr As String
                        For y As Integer = 0 To (langsObj.Count - 1)
                            curLang = String.Empty
                            If langsObj.Value(Of String)(y) IsNot Nothing Then curLang = localizeLanguage(langsObj.Value(Of String)(y)).ToUpper(CultureInfo.CurrentCulture)
                            buildStr = String.Empty
                            For n As Integer = 0 To 72
                                Dim styleStr As String = String.Empty
                                If langsObj.Value(Of String)(y) Is "TAMIL" Or langsObj.Value(Of String)(y) Is "KOREAN" Then
                                    styleStr = " style=""font-family:'Arial Unicode MS';"""
                                End If
                                Dim curBook As JArray = bibleBooksTemp.Item(n)
                                Dim curBookCurLang As JArray = JArray.Parse(curBook.Item(y).ToString)
                                Dim str1 As String = String.Empty
                                If curBookCurLang.Value(Of String)(0) IsNot Nothing Then str1 = curBookCurLang.Value(Of String)(0)
                                Dim str2 As String = String.Empty
                                If curBookCurLang.Value(Of String)(1) IsNot Nothing Then str2 = curBookCurLang.Value(Of String)(1)
                                buildStr += "<tr><td" + styleStr + ">" + str1 + "</td><td" + styleStr + ">" + str2 + "</td></tr>"
                            Next
                            booksAndAbbreviations.Add(curLang, buildStr)
                        Next

                    End Using

                Else
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "we seem to have a null connection... arghhh!")
                End If
            End Using
        End If

    End Sub

    Private Function localizeLanguage(ByVal language As String) As String
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Attempting to localize language <" & language & ">")
        language = language.ToUpper(CultureInfo.CurrentCulture)
        Dim langCode As String = String.Empty
        If langcodes.TryGetValue(language, langCode) Then
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ">> localization is now taking place...")
            Dim myCulture As CultureInfo = New CultureInfo(langCode, False)
            Return myCulture.DisplayName.ToUpper(CultureInfo.CurrentCulture)
        Else
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ">> Oops, localization does not seem to have been successful. Returning original language string.")
            Return language
        End If
        Return Nothing
    End Function

    Private Sub BuildLangCodes()
        'ISO language codes supported by Microsoft, taken from https://msdn.microsoft.com/it-it/goglobal/bb896001.aspx
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
    End Sub

    Private Function GetMimeType(ByVal i As Image) As String
        Dim imgguid As Guid = i.RawFormat.Guid
        For Each codec As ImageCodecInfo In ImageCodecInfo.GetImageDecoders()
            If codec.FormatID = imgguid Then Return codec.MimeType
        Next
        Return "image/unknown"
    End Function


End Class

Friend NotInheritable Class NativeMethods
    <DllImport("urlmon.dll")> _
    <PreserveSig> _
    Public Shared Function CoInternetSetFeatureEnabled(FeatureEntry As Integer, <MarshalAs(UnmanagedType.U4)> dSFlags As Integer, <MarshalAs(UnmanagedType.U1)> eEnable As Boolean) As <MarshalAs(UnmanagedType.[Error])> Integer
    End Function

    Private Sub New()
    End Sub
End Class