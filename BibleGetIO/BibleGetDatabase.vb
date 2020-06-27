Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports System.Globalization

Public Class BibleGetDatabase

    Private databaseFile As String = "BibleGetIO.sqlite"
    Private dbFullPath As String = Path.Combine(BibleGetAddIn.ThisAppDataDirectory, databaseFile)
    Private _connectionStr As String
    Private _INITIALIZED As Boolean
    Private DEBUG_MODE As Boolean

    ReadOnly Property connectionStr As String
        Get
            Return _connectionStr
        End Get
    End Property

    ReadOnly Property IsInitialized As Boolean
        Get
            Return _INITIALIZED
        End Get
    End Property

    Public Sub New()
        DEBUG_MODE = My.Settings.DEBUG_MODE
        _INITIALIZED = Initialize()
    End Sub

    Public Function Initialize() As Boolean
        Dim success As Boolean = False

        _connectionStr = "Data Source=" + dbFullPath

        If Not Directory.Exists(BibleGetAddIn.ThisAppDataDirectory) Then
            Try
                Dim dirInfo As DirectoryInfo
                dirInfo = Directory.CreateDirectory(BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(TimeOfDay.ToLongTimeString & " >> Directory created successfully: " & dirInfo.ToString)
            Catch ex As UnauthorizedAccessException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As ArgumentNullException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As ArgumentException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As PathTooLongException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As DirectoryNotFoundException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As IOException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As NotSupportedException
                Diagnostics.Debug.WriteLine("UnauthorizedAccessException caught while trying to create directory: " & BibleGetAddIn.ThisAppDataDirectory)
                Diagnostics.Debug.WriteLine(ex.Message)
            End Try

        End If
        'First check that the database file exists
        If Not exists() Then
            SQLiteConnection.CreateFile(dbFullPath)
        End If

        If Not metadataTableExists() Then
            success = UpdateMetaDataTable(True)
        Else
            success = True
        End If
        Return success
    End Function

    Public Shared Function disconnect(ByVal myConn As SQLiteConnection) As Boolean
        If myConn IsNot Nothing Then
            If myConn.State = ConnectionState.Closed Then
                Return True
            Else
                myConn.Dispose()
                If myConn Is Nothing Then
                    Return True
                Else
                    Return False
                End If
            End If
        End If
        Return False
    End Function

    Public Function exists() As Boolean
        Return File.Exists(dbFullPath)
    End Function

    Private Function metadataTableExists() As Boolean
        Dim AlreadyExisted As Boolean = True
        Dim res As Integer

        Dim queryString As String
        queryString = "CREATE TABLE IF NOT EXISTS METADATA("
        queryString &= "ID INTEGER, "
        For index = 0 To 72
            queryString &= "BIBLEBOOKS" + index.ToString() + " TEXT, "
        Next
        queryString &= "LANGUAGES TEXT, "
        queryString &= "VERSIONS TEXT"
        queryString &= ")"

        'Dim queryString2 As String
        'queryString2 = "CREATE TABLE IF NOT EXISTS TRANSIENTS("
        'queryString2 &= "NAME TEXT PRIMARY KEY, "
        'queryString2 &= "VALUE TEXT"
        'queryString2 &= ")"

        Try
            Using conn As New SQLiteConnection(_connectionStr)
                conn.Open()
                Using sqlQuery As New SQLiteCommand(conn)
                    sqlQuery.CommandText = queryString
                    res = sqlQuery.ExecuteNonQuery()
                    If res = 0 Then
                        queryString = "INSERT INTO METADATA (ID) VALUES (0)"
                        sqlQuery.CommandText = queryString
                        res = sqlQuery.ExecuteNonQuery()
                        AlreadyExisted = False
                    ElseIf res = -1 Then
                        AlreadyExisted = True
                    End If

                    'sqlQuery.CommandText = queryString2
                    'res = sqlQuery.ExecuteNonQuery()
                    'If res = 0 Then
                    '    queryString2 = "INSERT INTO TRANSIENTS (NAME,VALUE) VALUES (UPDATING,FALSE)"
                    '    sqlQuery.CommandText = queryString2
                    '    res = sqlQuery.ExecuteNonQuery()
                    'End If
                End Using
            End Using

        Catch ex As SQLiteException
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
            'If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & ex.Message)
        End Try
        Return AlreadyExisted
    End Function

    Public Function UpdateMetaDataTable() As Boolean
        Return UpdateMetaDataTable(False)
    End Function

    Public Function UpdateMetaDataTable(ByVal create As Boolean) As Boolean
        Dim res As Integer
        Dim queryString As String
        Dim response As String
        Dim success As Boolean = True
        Try
            Using conn As New SQLiteConnection(_connectionStr)
                conn.Open()
                Using sqlQuery As New SQLiteCommand(conn)

                    response = HTTPCaller.GetMetaData("biblebooks")
                    If response IsNot Nothing Then
                        Dim responseObj As JToken = JObject.Parse(response)
                        Dim resultsObj As JToken = responseObj.SelectToken("results")
                        Dim langsObj As JToken = responseObj.SelectToken("languages")
                        'Dim errsObj As JToken = responseObj.SelectToken("errors")
                        If resultsObj Is Nothing Or langsObj Is Nothing Then
                            success = False
                        Else
                            'Dim langsCount As Integer = langsObj.Count
                            For i As Integer = 0 To resultsObj.Count - 1
                                Dim currentBibleBook As String = resultsObj.Item(i).ToString
                                queryString = "UPDATE METADATA SET BIBLEBOOKS" + i.ToString(CultureInfo.InvariantCulture()) + "='" + currentBibleBook + "' WHERE ID=0"
                                sqlQuery.CommandText = queryString
                                res = sqlQuery.ExecuteNonQuery
                                If res <> 1 Then
                                    success = False
                                End If
                            Next

                            Dim langsStr As String = langsObj.ToString
                            queryString = "UPDATE METADATA SET LANGUAGES='" + langsStr + "' WHERE ID=0"
                            sqlQuery.CommandText = queryString
                            res = sqlQuery.ExecuteNonQuery
                            If res <> 1 Then
                                success = False
                            End If

                        End If
                    Else
                        success = False
                    End If

                    response = HTTPCaller.GetMetaData("bibleversions")
                    If response IsNot Nothing Then
                        Dim responseObj As JToken = JObject.Parse(response)
                        'Dim resultsObj As JToken = responseObj.SelectToken("results")
                        'Dim errsObj As JToken = responseObj.SelectToken("errors")
                        Dim validversionsObj As JObject = responseObj.SelectToken("validversions_fullname")
                        If validversionsObj Is Nothing Then
                            success = False
                        Else
                            Dim validversionsStr As String = validversionsObj.ToString
                            queryString = "UPDATE METADATA SET VERSIONS='" + validversionsStr + "' WHERE ID=0"
                            sqlQuery.CommandText = queryString
                            res = sqlQuery.ExecuteNonQuery
                            If res <> 1 Then
                                success = False
                            Else
                                Dim keys() As String = validversionsObj.Properties().Select(Function(p) p.Name).ToArray()
                                Dim versionsStr = Join(keys, ",")

                                response = HTTPCaller.GetMetaData("versionindex&versions=" + versionsStr)
                                If response IsNot Nothing Then
                                    responseObj = JObject.Parse(response)
                                    Dim indexes As JObject = responseObj.SelectToken("indexes")
                                    If indexes Is Nothing Then
                                        success = False
                                    Else
                                        For Each name As String In indexes.Properties().Select(Function(p) p.Name).ToArray()
                                            Dim jsonObj As JObject = New JObject
                                            Dim versionindex As JToken = indexes.SelectToken(name)
                                            If versionindex Is Nothing Then
                                                success = False
                                            Else
                                                jsonObj.Add("book_num", versionindex.SelectToken("book_num"))
                                                jsonObj.Add("chapter_limit", versionindex.SelectToken("chapter_limit"))
                                                jsonObj.Add("verse_limit", versionindex.SelectToken("verse_limit"))
                                                Dim versionindex_str As String = jsonObj.ToString
                                                Dim rslt As Boolean = createColumnIfNotExists(name + "IDX")
                                                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "result of createColumnIfNotExists: " & rslt.ToString)
                                                If rslt Then
                                                    If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Now that we are sure that column " & name & "IDX exists, we shall proceed to update its value")
                                                    queryString = "UPDATE METADATA SET " + name + "IDX='" + versionindex_str + "' WHERE ID=0"
                                                    sqlQuery.CommandText = queryString
                                                    res = sqlQuery.ExecuteNonQuery
                                                    If res <> 1 Then
                                                        success = False
                                                    End If
                                                Else
                                                    success = False
                                                End If

                                            End If
                                        Next
                                    End If
                                Else
                                    success = False
                                End If
                            End If

                        End If


                    Else
                        success = False
                    End If
                End Using
            End Using

        Catch ex As SQLiteException
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
            Return False
        End Try

        Return success
    End Function

    Public Function GetMetaData(ByVal dataOption As String) As String
        Dim result As String = String.Empty
        If dataOption Is Nothing Then dataOption = String.Empty
        dataOption = dataOption.ToUpper(CultureInfo.InvariantCulture)
        If dataOption.StartsWith("BIBLEBOOKS", StringComparison.Ordinal) Or dataOption.Equals("LANGUAGES", StringComparison.Ordinal) Or dataOption.Equals("VERSIONS", StringComparison.Ordinal) Or dataOption.EndsWith("IDX", StringComparison.Ordinal) Then
            Try
                Dim sqlexec As String = "SELECT " + dataOption + " FROM METADATA WHERE ID=0"
                Using conn As New SQLiteConnection(_connectionStr)
                    conn.Open()
                    Using sqlQuery As New SQLiteCommand(conn)
                        sqlQuery.CommandText = sqlexec
                        result = sqlQuery.ExecuteScalar().ToString
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "retrieving metadata from our sqlite database, option requested = " & dataOption & " & result = " & result)
                    End Using
                End Using

            Catch ex As SQLiteException
                If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
            End Try
        End If

        Return result

    End Function

    'Public Sub SetUpdating(ByVal value As Boolean)
    '    Try
    '        Dim res As Integer
    '        Dim sqlexec As String = "UPDATE TRANSIENTS SET VALUE=" & value.ToString & " WHERE NAME=UPDATING"
    '        Using conn As New SQLiteConnection(_connectionStr)
    '            conn.Open()
    '            Using sqlQuery As New SQLiteCommand(conn)
    '                sqlQuery.CommandText = sqlexec
    '                res = sqlQuery.ExecuteNonQuery()
    '            End Using
    '        End Using
    '    Catch ex As SQLiteException
    '        If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & ex.Message)
    '    End Try
    'End Sub

    'Public Function GetUpdating() As String
    '    Dim res As String = "FALSE"
    '    Try
    '        Dim sqlexec As String = "SELECT VALUE FROM TRANSIENTS WHERE NAME=UPDATING"
    '        Using conn As New SQLiteConnection(_connectionStr)
    '            conn.Open()
    '            Using sqlQuery As New SQLiteCommand(conn)
    '                sqlQuery.CommandText = sqlexec
    '                res = sqlQuery.ExecuteScalar()
    '            End Using
    '        End Using
    '    Catch ex As SQLiteException
    '        If Me.DEBUG_MODE Then ThisAddIn.LogInfoToDebug(Me.GetType().FullName & vbTab & ex.Message)
    '    End Try
    '    Return res.ToUpper
    'End Function
    Public Function createColumnIfNotExists(ByVal columnName) As Integer
        Return createColumnIfNotExists(columnName, "METADATA")
    End Function


    Public Function createColumnIfNotExists(ByVal columnName, ByVal tableName) As Boolean
        Dim res As Boolean = False
        Try
            'Dim sqlexec As String = "SELECT VALUE FROM TRANSIENTS WHERE NAME=UPDATING"
            Using conn As New SQLiteConnection(_connectionStr)
                conn.Open()

                Using sqlQuery As New SQLiteCommand(conn)
                    sqlQuery.CommandText = String.Format("PRAGMA table_info({0})", tableName)
                    Dim reader As SQLiteDataReader = sqlQuery.ExecuteReader()
                    Dim columnNames As New List(Of String)
                    While reader.Read()
                        'does column exist?
                        If DEBUG_MODE = True Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & reader("name"))
                        columnNames.Add(reader("name").ToString)
                    End While
                    reader.Close()

                    If Not columnNames.Contains(columnName) Then
                        If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & "Now creating column " + columnName)
                        Dim queryString As String = "ALTER TABLE " & tableName & " ADD COLUMN " + columnName + " TEXT"
                        sqlQuery.CommandText = queryString
                        res = sqlQuery.ExecuteNonQuery
                        If res = 0 Then
                            res = True
                        End If
                    Else : res = True
                    End If
                End Using
            End Using
        Catch ex As SQLiteException
            If DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug([GetType]().FullName & vbTab & ex.Message)
            res = False
        End Try
        Return res
    End Function

End Class

