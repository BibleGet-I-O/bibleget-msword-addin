Imports System.Data.SQLite
Imports System.Globalization
Imports Newtonsoft.Json.Linq

Public Class LocalizedBibleBooks

    Private BookAbbreviations As New Dictionary(Of Integer, String)
    Private BookNames As New Dictionary(Of Integer, String)
    Private curLangIsoCode As String
    Private curLangDisplayName As String
    Private curLangIdx As Integer
    Private DEBUG_MODE = My.Settings.DEBUG_MODE

    Public Sub New()
        curLangIsoCode = BibleGetAddIn.locale.TwoLetterISOLanguageName
        curLangDisplayName = New CultureInfo(curLangIsoCode).EnglishName.ToUpper
        Dim bibleGetDB As New BibleGetDatabase
        If bibleGetDB.IsInitialized Then
            Using conn As New SQLiteConnection(bibleGetDB.connectionStr)
                If conn IsNot Nothing Then
                    conn.Open()
                    Using sqlQuery As New SQLiteCommand(conn)
                        sqlQuery.CommandText = "SELECT LANGUAGES FROM METADATA WHERE ID=0"
                        Dim langsSupported As String = sqlQuery.ExecuteScalar
                        'Debug.Print("langsSupported = " & langsSupported)
                        Dim langsObj = JArray.Parse(langsSupported).ToObject(Of String())
                        'Debug.Print("langsObj = ")
                        'Debug.Print(String.Join(",", langsObj))
                        curLangIdx = Array.IndexOf(langsObj, curLangDisplayName)
                        'Debug.Print("curLangIdx = " & curLangIdx.ToString)
                        Dim bbBooks As String
                        Dim bookName As String = ""
                        Dim bookAbbrev As String = ""
                        For i As Integer = 0 To 72
                            sqlQuery.CommandText = "SELECT BIBLEBOOKS" & i.ToString(CultureInfo.InvariantCulture) & " FROM METADATA WHERE ID=0"
                            bbBooks = sqlQuery.ExecuteScalar
                            'Debug.Print("bbBooks = " & bbBooks)
                            Dim bibleBooksObj As JArray = JArray.Parse(bbBooks)
                            'Debug.Print("bibleBooksObj = " & String.Join(",", bibleBooksObj))
                            Dim bibleBooksInCurLang As JArray = bibleBooksObj.Item(curLangIdx)
                            'Debug.Print("bibleBooksInCurLang = " & String.Join(",", bibleBooksInCurLang))
                            bookName = bibleBooksInCurLang.Value(Of String)(0)
                            bookAbbrev = bibleBooksInCurLang.Value(Of String)(1)
                            If bookName.Contains("|") Then
                                bookName = bookName.Split("|").First.Trim
                            End If
                            If bookAbbrev.Contains("|") Then
                                bookAbbrev = bookAbbrev.Split("|").First.Trim
                            End If
                            BookAbbreviations.Add(i, bookAbbrev)
                            BookNames.Add(i, bookName)
                        Next


                    End Using

                Else
                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "we seem to have a null connection... arghhh!")
                End If
            End Using
        End If

    End Sub

    Public Function GetBookByIndex(idx As Integer) As LocalizedBibleBook
        Return New LocalizedBibleBook(BookAbbreviations.Item(idx), BookNames.Item(idx))
    End Function
End Class

Public Class LocalizedBibleBook
    Public Abbrev As String
    Public Fullname As String
    Public Sub New(abb As String, name As String)
        Abbrev = abb
        Fullname = name
    End Sub
End Class
