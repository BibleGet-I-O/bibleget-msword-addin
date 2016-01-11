Public Class BibleVersion
    Private myAbbrev As String
    Private myFullname As String
    Private myYear As String
    Private myLang As String

    Public Sub New(ByVal strAbbrev As String, ByVal strFullname As String, ByVal strYear As String, ByVal strLang As String)
        Me.myAbbrev = strAbbrev
        Me.myFullname = strFullname
        Me.myYear = strYear
        Me.myLang = strLang
    End Sub 'NewNew

    Public ReadOnly Property Abbrev() As String
        Get
            Return myAbbrev
        End Get
    End Property

    Public ReadOnly Property Fullname() As String
        Get
            Return myFullname
        End Get
    End Property

    Public ReadOnly Property Year() As String
        Get
            Return myYear
        End Get
    End Property


    Public ReadOnly Property Lang() As String
        Get
            Return myLang
        End Get
    End Property
End Class
