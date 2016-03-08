'/**
' *
' * @author Lwangaman
' */
Imports Newtonsoft.Json.Linq

Public Class Indexes

    Private bibleGetDB As BibleGetDatabase
    Private versionsabbrev As List(Of String) = Nothing
    Private VersionIndexes As Dictionary(Of String, VersionIDX)

    Public Sub New()
        Me.bibleGetDB = New BibleGetDatabase()
        If Me.versionsabbrev Is Nothing Then
            Dim versions As String = bibleGetDB.getMetaData("VERSIONS")
            Dim jsbooks As JObject = JObject.Parse(versions)
            versionsabbrev = jsbooks.Properties().Select(Function(p) p.Name).ToList
        End If

        'Dim len As Integer = versionsabbrev.Count

        Me.VersionIndexes = New Dictionary(Of String, VersionIDX)
        For Each s As String In versionsabbrev
            VersionIndexes.Add(s, New VersionIDX(s))
        Next

    End Sub


    Public Function isValidVersion(ByVal version As String) As Boolean
        Return versionsabbrev.Contains(version)
        'version = version.ToUpper
        'For Each s As String In versionsabbrev
        '    If s.Equals(version) Then
        '        Return True
        '    End If
        'Next
        'Return False
    End Function

    Public Function isValidChapter(ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Boolean
        Dim flag As Boolean = True
        'Diagnostics.Debug.WriteLine(String.Join(",", selectedVersions))
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            If VersionIndexes(version).chapter_limit()(idx) < chapter Then flag = False
        Next
        Return flag
    End Function


    Public Function isValidVerse(ByVal verse As Integer, ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Boolean
        Dim flag As Boolean = True
        For Each version As String In selectedVersions
            Dim idx As Integer = VersionIndexes(version).book_num().IndexOf(book)
            '//System.out.println("corresponding book index in VersionIndexes for version "+version+", book "+Integer.toString(book)+" is "+Integer.toString(idx));
            If VersionIndexes(version).verse_limit(idx)(chapter - 1) < verse Then flag = False
        Next
        Return flag
    End Function

    Public Function getChapterLimit(ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Integer()
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

    Public Function getVerseLimit(ByVal chapter As Integer, ByVal book As Integer, ByVal selectedVersions As List(Of String)) As Integer()
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

        Public Sub New()
            Throw New Exception("Must initialize with desired version.")
        End Sub

        Public Sub New(ByVal version As String)
            Dim bibleGetDB As New BibleGetDatabase
            Dim versionIdxStr As String = bibleGetDB.getMetaData(version + "IDX")
            'JsonReader jsonReader = Json.createReader(new StringReader(versionIdxStr));
            Me.versionIDX = JObject.Parse(versionIdxStr) 'jsonReader.readObject();
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
            'Diagnostics.Debug.WriteLine("verse_limit array has " + len.ToString + " elements which correspond to actual verse limits in the given book " & book)
            '//System.out.println("verse_limit array has "+Integer.toString(len)+" elements which correspond to actual verse limits in the given book");
            'Diagnostics.Debug.WriteLine(verselimit.ToString)
            Dim verseLimit_array(len - 1) As Integer
            For i As Integer = 0 To verseLimit_array.GetUpperBound(0)
                verseLimit_array(i) = verselimit.Value(Of Integer)(i)
            Next
            'Diagnostics.Debug.WriteLine("verseLimit_array has now been filled with " & verseLimit_array.Count & " values")
            'Diagnostics.Debug.WriteLine(String.Join(",", verseLimit_array))
            Return verseLimit_array.ToList
        End Function

    End Class


End Class
