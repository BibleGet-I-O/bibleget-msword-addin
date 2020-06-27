Imports System.Net
Imports System.IO


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
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Try
            request = CType(WebRequest.Create(uri), HttpWebRequest)
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
        Dim url As String = "https://bibleget.io/?wpdm_api=API_KEY&package_id=596&task=getpackageversion"
        Dim response As String = GetResponse(New Uri(url))
        If response IsNot Nothing Then
            response = response.Replace("""", "")
            If My.Settings.DEBUG_MODE Then BibleGetAddIn.LogInfoToDebug("HTTPCaller.vb" & vbTab & "onlineVersion = " & response)
            retVersion = New Version(response)
        End If
        Return retVersion
    End Function

End Class
