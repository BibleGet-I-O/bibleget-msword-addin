Imports System.Globalization


Public Class ThisAddIn

    Public Shared RM As Resources.ResourceManager = New Resources.ResourceManager("BibleGetIO.BibleGetResource", System.Reflection.Assembly.GetExecutingAssembly())
    Public Shared locale As CultureInfo = CultureInfo.CurrentCulture
    'Public Shared helpFile As String

    Public Shared Function __(ByVal myStr As String) As String
        Dim myTranslation As String = RM.GetString(myStr, locale)
        If Not String.IsNullOrEmpty(myTranslation) Then
            Return myTranslation
        Else
            Return myStr
        End If
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'Dim myKey As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Word\Addins\BibleGetIO"
        'Dim myLoc As String = My.Computer.Registry.GetValue(myKey, "Manifest", Nothing)
        'Dim installPath As String = myLoc.Split("|")(0)
        'Dim installPathBase As String = System.IO.Path.GetDirectoryName(installPath)
        'UNDONE: now using a Windows Form as Help
        'helpFile = System.IO.Path.Combine(installPathBase, "bibleget-io.chm")
        'Diagnostics.Debug.WriteLine("helpFile path = " + helpFile)
        'Diagnostics.Debug.WriteLine("AddIn path = " + installPath)
        'Dim lang As String = culture1.TwoLetterISOLanguageName
        'System.Diagnostics.Debug.WriteLine(String.Format("The current culture is {0}", lang))
        'System.Diagnostics.Debug.WriteLine(RM.BaseName)

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Private Sub Application_DocumentBeforeSave(ByVal Doc As Word.Document, ByRef SaveAsUI As Boolean, _
    '    ByRef Cancel As Boolean) Handles Application.DocumentBeforeSave
    '    'Doc.Paragraphs(1).Range.InsertParagraphBefore()
    '    'Doc.Paragraphs(1).Range.Text = "This text was added by using code."
    'End Sub

    'Private Sub Application_Startup() Handles Application.Startup
    '    ''Dim oCult As String = Application.Language.ToString
    '    ''Dim culture2 As CultureInfo = Thread.CurrentThread.CurrentCulture
    '    'Dim culture1 As CultureInfo = CultureInfo.CurrentCulture
    '    'Dim lang As String = culture1.TwoLetterISOLanguageName
    '    'System.Diagnostics.Debug.WriteLine(String.Format("The current culture is {0}", lang))
    '    'RM = New Resources.ResourceManager("BibleGetIO.BibleGetResource", System.Reflection.Assembly.GetExecutingAssembly())
    '    'System.Diagnostics.Debug.WriteLine(RM.BaseName)
    '    'Dim greeting As String = RM.GetString("About this plugin", CultureInfo.CurrentCulture)
    '    'System.Diagnostics.Debug.WriteLine(greeting)
    'End Sub

End Class
