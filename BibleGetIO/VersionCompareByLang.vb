Imports System.Collections
Imports System.Globalization

Public Class VersionCompareByLang

    Implements IComparer

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Return New CaseInsensitiveComparer(CultureInfo.CurrentCulture).Compare(x.Lang, y.Lang)
    End Function

End Class
