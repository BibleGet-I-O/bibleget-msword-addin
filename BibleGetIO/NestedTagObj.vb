Imports System.Text.RegularExpressions

Public Class NestedTagObj

    Private remainingText As String
    Private NABREfmt As String = "(.*?)<((speaker|sm|i|pr|po)[f|l|s|i|3]{0,1}[f|l]{0,1})>(.*?)</\2>"
    Private NABREfmtMatch As Match
    Public Before As String
    Public Contents As String
    Public After As String
    Public Tag As String

    Public Sub New(ByVal formattingTagContents As String)
        remainingText = formattingTagContents
        For Each NABREfmtMatch As Match In Regex.Matches(formattingTagContents, NABREfmt, RegexOptions.Singleline)
            If NABREfmtMatch.Groups(2).Value IsNot Nothing And NABREfmtMatch.Groups(2).Value IsNot String.Empty Then
                Tag = NABREfmtMatch.Groups(2).Value
                If NABREfmtMatch.Groups(1).Value IsNot Nothing And NABREfmtMatch.Groups(1).Value IsNot String.Empty Then
                    Before = NABREfmtMatch.Groups(1).Value
                    remainingText = remainingText.Replace(Before, "")
                End If

                Contents = NABREfmtMatch.Groups(4).Value
                After = remainingText.Replace("<" & Tag & ">" & Contents & "</" & Tag & ">", "")
            End If
        Next
    End Sub
End Class
