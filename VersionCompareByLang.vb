Imports System.Collections

Public Class VersionCompareByLang

    Implements IComparer

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Return New CaseInsensitiveComparer().Compare(x.Lang, y.Lang)
    End Function

End Class
