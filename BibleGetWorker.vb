Imports System.Net

Public Class BibleGetWorker
    Private myCommand As String
    Private myQueryString As String
    Private myWebResponse As WebResponse

    Public Sub New(ByVal strCommand As String, ByVal strQueryString As String)
        Me.myCommand = strCommand
        Me.myQueryString = strQueryString
    End Sub

    Public Sub New(ByVal strCommand As String, ByVal xWebResponse As WebResponse)
        Me.myCommand = strCommand
        Me.myWebResponse = xWebResponse
    End Sub

    Public ReadOnly Property Command() As String
        Get
            Return myCommand
        End Get
    End Property

    Public ReadOnly Property QueryString() As String
        Get
            Return myQueryString
        End Get
    End Property

    Public ReadOnly Property WebResponse() As WebResponse
        Get
            Return myWebResponse
        End Get
    End Property

End Class
