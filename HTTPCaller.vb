Imports System.Net
Imports System.IO


Public Class HTTPCaller

    Public Shared Function sendGet(ByVal myQuery As String, ByVal versions As String) As String
        Try
            versions = System.Uri.EscapeDataString(versions)
            myQuery = System.Uri.EscapeDataString(myQuery)
        Catch ex As Exception
            Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        Dim url As String = "http://query.bibleget.io/index.php?query=" + myQuery + "&version=" + versions + "&return=json&appid=msword&pluginversion=" + My.Application.Info.Version.ToString
        Return getResponse(url)
    End Function

    Public Shared Function getResponse(ByVal url As String) As String
        Dim request As WebRequest = WebRequest.Create(url)
        Dim response As WebResponse = request.GetResponse()
        If CType(response, HttpWebResponse).StatusCode = HttpStatusCode.OK Then
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As New StreamReader(dataStream, Encoding.UTF8)
            Dim responseFromServer As String = reader.ReadToEnd()
            Return responseFromServer
        Else
            Diagnostics.Debug.WriteLine("Error contacting server. HTTP Status Code: " & CType(response, HttpWebResponse).StatusDescription)
        End If
        Return Nothing
    End Function

    Public Shared Function getMetaData(ByVal query As String) As String
        Dim url As String = "http://query.bibleget.io/metadata.php?query=" + query
        Dim response As String = getResponse(url)
        If response IsNot Nothing Then
            Return response
        End If
        Return Nothing
    End Function


End Class
