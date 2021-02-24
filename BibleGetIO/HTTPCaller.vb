Imports System.Net
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports System.Diagnostics

Public NotInheritable Class HTTPCaller


    Public Shared Function SendGet(ByVal myQuery As String, ByVal versions As String) As String
        versions = Uri.EscapeDataString(versions)
        myQuery = Uri.EscapeDataString(myQuery)

        Dim url As String = BibleGetAddIn.BGET_ENDPOINT & "?query=" + myQuery + "&version=" + versions + "&return=json&appid=msword&pluginversion=" + My.Application.Info.Version.ToString
        Return GetResponse(New Uri(url))
    End Function

    Public Shared Function GetResponse(ByVal url As String) As String
        Dim uri As New Uri(url)
        Return GetResponse(uri)
    End Function

    Public Shared Function GetResponse(ByVal uri As Uri) As String
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Try
            request = CType(WebRequest.Create(uri), HttpWebRequest)
            request.UserAgent = "Windows NT .Net Client"
            response = CType(request.GetResponse(), HttpWebResponse)
            If response IsNot Nothing And response.StatusCode = HttpStatusCode.OK Then
                Dim dataStream As Stream = response.GetResponseStream()
                Dim reader As New StreamReader(dataStream, Encoding.UTF8)
                Dim responseFromServer As String = reader.ReadToEnd()
                response.Close()
                Return responseFromServer
            Else
                If My.Settings.DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("HTTPCaller.vb" & vbTab & "Error contacting server. HTTP Status Code: " & response.StatusDescription)
            End If
        Catch e As WebException
            Diagnostics.Debug.WriteLine(e.Message)

            If e.Status = WebExceptionStatus.ProtocolError Then
                Diagnostics.Debug.WriteLine("Status Code : {0}", CType(e.Response, HttpWebResponse).StatusCode)
                Diagnostics.Debug.WriteLine("Status Description : {0}", CType(e.Response, HttpWebResponse).StatusDescription)
            End If

        Catch e As Exception
            Diagnostics.Debug.WriteLine(e.Message)


        End Try

        Return Nothing
    End Function

    Public Shared Function GetMetaData(ByVal query As String) As String
        Dim url As String = BibleGetAddIn.BGET_METADATA_ENDPOINT & "?query=" + query
        Dim response As String = GetResponse(New Uri(url))
        If response IsNot Nothing Then
            Return response
        End If
        Return Nothing
    End Function

    Public Shared Function GetCurrentVersion() As Version
        Dim retVersion As Version = My.Application.Info.Version 'initialize as known current version
        'Dim url As String = "https://bibleget.io/?wpdm_api=API_KEY&package_id=596&task=getpackageversion"
        Dim url As String = "https://sourceforge.net/projects/bibleget/best_release.json"
        Dim response As String = GetResponse(New Uri(url))
        If response IsNot Nothing Then
            'Debug.WriteLine(response)
            Dim BestRelease As JObject = JObject.Parse(response)
            Dim fileName As String = BestRelease.SelectToken("$.platform_releases.windows.filename").Value(Of String)()
            'Console.WriteLine("Filename of current best version from sourceforge = " & fileName)
            'Debug.WriteLine("Filename of current best version from sourceforge = " & fileName)
            Dim detectedVersion As String = fileName.Substring(2, 7) 'skip /v and get the version info such as 3.0.1.2 (7 characters)
            'Console.WriteLine("Version detected from filename = " & detectedVersion)
            'Debug.WriteLine("Version detected from filename = " & detectedVersion)
            If My.Settings.DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("HTTPCaller.vb" & vbTab & "onlineVersion = " & detectedVersion)
            retVersion = New Version(detectedVersion)
        End If
        Return retVersion
    End Function

End Class
