'/**
' *
' * @author Lwangaman
' */
Imports Newtonsoft.Json.Linq

Public Class Indexes

    Private bibleGetDB As BibleGetDatabase
    Private versionsabbrev As List(Of String) = Nothing
    Private VersionIndexes As Dictionary(Of String, VersionIDX)
    Private DEBUG_MODE As Boolean

    Public Sub New()
        DEBUG_MODE = My.Settings.DEBUG_MODE
        bibleGetDB = New BibleGetDatabase()
        If versionsabbrev Is Nothing Then
            Dim versions As String = bibleGetDB.getMetaData("VERSIONS")
            Dim jsbooks As JObject = JObject.Parse(versions)
            versionsabbrev = jsbooks.Properties().Select(Function(p) p.Name).ToList
        End If

        'Dim len As Integer = versionsabbrev.Count

        VersionIndexes = New Dictionary(Of String, VersionIDX)
        For Each s As String In versionsabbrev
            VersionIndexes.Add(s, New VersionIDX(s))
        Next

    End Sub


    Public Function IsValidVersion(ByVal version As String) As Boolean
        Return versionsabbrev.Contains(version)
        'version = version.ToUpper
        'For Each s As String In versionsabbrev
        '    If s.Equals(version) Then
        '        Return True
        '    End If
        'Next
        'Return False
    End Function

    Public Function IsValidChapter(ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Boolean
        Dim flag As Boolean = True
        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & String.Join(",", selectedVersions))
        If selectedVersions Is Nothing Then Return False
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            If idx = -1 Then
                flag = False
            ElseIf VersionIndexes(version).chapter_limit()(idx) < chapter Then
                flag = False
            End If
        Next
        Return flag
    End Function


    Public Function IsValidVerse(ByVal verse As Integer, ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Boolean
        Dim flag As Boolean = True
        If selectedVersions Is Nothing Then Return False
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            '//System.out.println("corresponding book index in VersionIndexes for version "+version+", book "+Integer.toString(book)+" is "+Integer.toString(idx));
            If idx = -1 Then
                flag = False
            ElseIf VersionIndexes(version).verse_limit(idx)(chapter - 1) < verse Then
                flag = False
            End If
        Next
        Return flag
    End Function

    Public Function GetChapterLimit(ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Integer()
        If selectedVersions Is Nothing Then
            selectedVersions = New List(Of String)
        End If
        Dim retInt(selectedVersions.Count - 1) As Integer
        Dim count As Integer = 0
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            retInt(count) = VersionIndexes(version).chapter_limit()(idx)
            count += 1
        Next
        '//        System.out.print("value of chapter retInt = ");
        '//        System.out.println(Arrays.toString(retInt));
        Return retInt
    End Function

    Public Function GetVerseLimit(ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Integer()
        If selectedVersions Is Nothing Then
            selectedVersions = New List(Of String)
        End If
        Dim retInt(selectedVersions.Count - 1) As Integer
        Dim count As Integer = 0
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            retInt(count) = VersionIndexes(version).verse_limit(idx)(chapter - 1)
            count += 1
        Next
        '//        System.out.print("value of verse retInt = ");
        '//        System.out.println(Arrays.toString(retInt));
        Return retInt
    End Function


    Private Class VersionIDX

        Private versionIDX As JObject
        Private DEBUG_MODE As Boolean

        Public Sub New(ByVal version As String)
            DEBUG_MODE = My.Settings.DEBUG_MODE
            Dim bibleGetDB As New BibleGetDatabase
            Dim versionIdxStr As String = bibleGetDB.GetMetaData(version + "IDX")
            'JsonReader jsonReader = Json.createReader(new StringReader(versionIdxStr));
            versionIDX = JObject.Parse(versionIdxStr) 'jsonReader.readObject();
        End Sub

        Public Function book_num() As List(Of Integer)
            'JsonArray booknum = versionIDX.getJsonArray("book_num");
            Dim booknum As JToken = versionIDX.SelectToken("book_num")
            Dim len As Integer = booknum.Count
            Dim booknum_array(len - 1) As Integer
            For i As Integer = 0 To booknum_array.GetUpperBound(0)
                booknum_array(i) = booknum.Value(Of Integer)(i)
            Next
            Return booknum_array.ToList
        End Function

        Public Function chapter_limit() As List(Of Integer)
            'JsonArray chapterlimit = versionIDX.getJsonArray("chapter_limit");
            Dim chapterlimit As JToken = versionIDX.SelectToken("chapter_limit")
            Dim len As Integer = chapterlimit.Count
            Dim chapterLimit_array(len - 1) As Integer
            For i As Integer = 0 To chapterLimit_array.GetUpperBound(0)
                chapterLimit_array(i) = chapterlimit.Value(Of Integer)(i)
            Next
            Return chapterLimit_array.ToList
        End Function

        Public Function verse_limit(ByVal book As Integer) As List(Of Integer)
            'JsonArray verselimit_temp = versionIDX.getJsonArray("verse_limit");
            Dim verselimit_temp As JToken = versionIDX.SelectToken("verse_limit")
            '//System.out.println(verselimit_temp.toString());
            '//System.out.println("array has "+Integer.toString(verselimit_temp.size())+" elements which correspond to books...");
            '//System.out.println("requesting array element corresponding to book "+Integer.toString(book+1)+" at index "+Integer.toString(book));

            'JsonArray verselimit = verselimit_temp.getJsonArray(book);
            Dim verselimit As JArray = JArray.FromObject(verselimit_temp(book))
            '//System.out.println(verselimit.toString());
            Dim len As Integer = verselimit.Count
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "verse_limit array has " + len.ToString + " elements which correspond to actual verse limits in the given book " & book)
            '//System.out.println("verse_limit array has "+Integer.toString(len)+" elements which correspond to actual verse limits in the given book");
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & verselimit.ToString)
            Dim verseLimit_array(len - 1) As Integer
            For i As Integer = 0 To verseLimit_array.GetUpperBound(0)
                verseLimit_array(i) = verselimit.Value(Of Integer)(i)
            Next
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "verseLimit_array has now been filled with " & verseLimit_array.Count & " values")
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & String.Join(",", verseLimit_array))
            Return verseLimit_array.ToList
        End Function

    End Class


End Class
