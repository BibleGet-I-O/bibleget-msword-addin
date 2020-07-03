Public Class SelectSingleBibleVersion
    Private Sub SelectSingleBibleVersion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InsertQuoteDialog.LoadBibleVersions(ListView1)

    End Sub

    Private Sub SetBibleVersionBtn_Click(sender As Object, e As EventArgs) Handles SetBibleVersionBtn.Click
        Dim bibleVersion As String
        Dim selectedItems As Windows.Forms.ListView.SelectedListViewItemCollection = ListView1.SelectedItems
        If ListView1.SelectedItems.Count = 0 Then
            bibleVersion = "NABRE"
        Else
            bibleVersion = InsertQuoteDialog.listItems(selectedItems.Item(0).Index)
        End If
        'BibleGetRibbon.BibleVersionForSearch
    End Sub
End Class