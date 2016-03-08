Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class BibleGetDatabase

    Private userProfilePath As String
    Private databaseHomeBase As String
    Private databaseHome As String
    Private databaseFile As String
    Private dbFullPath As String
    Public connectionString As String
    Public INITIALIZED As Boolean


    Public Sub New()
        Me.INITIALIZED = Me.initialize()
    End Sub

    Public Function initialize() As Boolean
        Dim success As Boolean = False

        Me.userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Me.databaseHomeBase = "BibleGetMSOfficePlugin" + Path.DirectorySeparatorChar
        Me.databaseHome = Me.getDatabaseHome()
        Me.databaseFile = "BibleGetIO.sqlite"
        Me.dbFullPath = userProfilePath + databaseHome + databaseFile
        Me.connectionString = "Data Source=" + dbFullPath

        'First check that the database file exists
        If Not Me.exists() Then
            SQLiteConnection.CreateFile(Me.dbFullPath)
        End If

        If Not Me.metadataTableExists() Then
            success = Me.updateMetaDataTable()
        Else
            success = True
        End If
        Return success
    End Function

    Public Function connect() As SQLiteConnection
        'Environment.SetEnvironmentVariable("PreLoadSQLite_BaseDirectory", System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location))
        Dim conn As SQLiteConnection
        Try
            conn = New SQLiteConnection(connectionString)
            conn.Open()
            If conn.State = ConnectionState.Open Then
                Return conn
            End If
        Catch ex As Exception
            Diagnostics.Debug.WriteLine(ex.Message)
            Return Nothing
        End Try
        Return Nothing
    End Function

    Public Function disconnect(ByVal myConn As SQLiteConnection) As Boolean
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
        Return False
    End Function


    Private Function getDatabaseHome() As String
        Dim databaseHome As String = ""
        Select Case Environment.OSVersion.Platform
            Case PlatformID.MacOSX
                databaseHome = Path.DirectorySeparatorChar + "Library" + Path.DirectorySeparatorChar + "Application Support" & Path.DirectorySeparatorChar & databaseHomeBase
            Case PlatformID.Unix
                databaseHome = Path.DirectorySeparatorChar & "." & databaseHomeBase
            Case PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.Win32NT, PlatformID.WinCE
                databaseHome = Path.DirectorySeparatorChar & "AppData" & Path.DirectorySeparatorChar & "Roaming" & Path.DirectorySeparatorChar & databaseHomeBase
            Case PlatformID.Xbox
                Return False
        End Select
        Return databaseHome
    End Function

    Public Function exists() As Boolean
        If Not Directory.Exists(userProfilePath + databaseHome) Then
            Directory.CreateDirectory(userProfilePath + databaseHome)
        End If
        Return File.Exists(dbFullPath)
    End Function

    Private Function metadataTableExists() As Boolean
        Dim conn As SQLiteConnection = Me.connect()

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
        Try
            Using conn
                Using sqlQuery As New SQLiteCommand(conn)
                    sqlQuery.CommandText = queryString
                    res = sqlQuery.ExecuteNonQuery()
                    If res = 0 Then
                        queryString = "INSERT INTO METADATA (ID) VALUES (0)"
                        sqlQuery.CommandText = queryString
                        res = sqlQuery.ExecuteNonQuery()
                        Return False
                    ElseIf res = -1 Then
                        Return True
                    End If
                End Using
            End Using

        Catch ex As Exception
            Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Function

    Private Function updateMetaDataTable() As Boolean
        Dim conn As SQLiteConnection = Me.connect()

        Dim res As Integer
        Dim queryString As String
        Dim response As String
        Dim success As Boolean = True
        Try
            Using conn
                Using sqlQuery As New SQLiteCommand(conn)

                    response = HTTPCaller.getMetaData("biblebooks")
                    If response IsNot Nothing Then
                        Dim responseObj As JToken = JObject.Parse(response)
                        Dim resultsObj As JToken = responseObj.SelectToken("results")
                        Dim langsObj As JToken = responseObj.SelectToken("languages")
                        'Dim errsObj As JToken = responseObj.SelectToken("errors")
                        If resultsObj Is Nothing Or langsObj Is Nothing Then
                            success = False
                        Else
                            Dim langsCount As Integer = langsObj.Count
                            For i As Integer = 0 To resultsObj.Count - 1
                                Dim currentBibleBook As String = resultsObj.Item(i).ToString
                                queryString = "UPDATE METADATA SET BIBLEBOOKS" + i.ToString + "='" + currentBibleBook + "' WHERE ID=0"
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

                    response = HTTPCaller.getMetaData("bibleversions")
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

                                response = HTTPCaller.getMetaData("versionindex&versions=" + versionsStr)
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
                                                queryString = "ALTER TABLE METADATA ADD COLUMN " + name + "IDX TEXT"
                                                sqlQuery.CommandText = queryString
                                                res = sqlQuery.ExecuteNonQuery
                                                'Diagnostics.Debug.WriteLine("ALTER TABLE query result: " + res.ToString)
                                                If res = 1 Then
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

        Catch ex As Exception
            Diagnostics.Debug.WriteLine(ex.Message)
            Return False
        End Try

        Return success
    End Function

    Public Function getMetaData(ByVal dataOption As String) As String
        Dim conn As SQLiteConnection = Me.connect()
        Dim result As String = String.Empty
        dataOption = dataOption.ToUpper
        If dataOption.StartsWith("BIBLEBOOKS") Or dataOption.Equals("LANGUAGES") Or dataOption.Equals("VERSIONS") Or dataOption.EndsWith("IDX") Then
            Try
                Dim sqlexec As String = "SELECT " + dataOption + " FROM METADATA WHERE ID=0"
                Using conn
                    Using sqlQuery As New SQLiteCommand(conn)
                        sqlQuery.CommandText = sqlexec
                        result = sqlQuery.ExecuteScalar().ToString
                        'Diagnostics.Debug.WriteLine("retrieving metadata from our sqlite database, option requested = " & dataOption & " & result = " & result)
                    End Using
                End Using

            Catch ex As Exception
                Diagnostics.Debug.WriteLine(ex.Message)
            End Try
        End If

        Return result

    End Function

End Class

