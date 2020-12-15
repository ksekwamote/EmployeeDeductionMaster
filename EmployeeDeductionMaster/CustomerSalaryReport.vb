Public Class CustomerSalaryReport

    Dim ds, dsaff As New DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ds.Clear()
        dsaff.Clear()
        ds = getSalariesByCustomerID(tbcustomerid.Text.Trim)
        For count As Integer = 0 To ds.Tables(0).Rows.Count
            dsaff = getAfforbilities(tbcustomerid.Text.Trim)
            If dsaff.Tables(0).Rows.Count > 0 Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dsaff.Tables(0).Rows(0).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count -1).Cells(1).Value = 
            DataGridView1.Rows(DataGridView1.Rows.Count -1).Cells(2).Value = 
DataGridView1.Rows(DataGridView1.Rows.Count -1).Cells(3).Value = 
            End If
        Next
    End Sub
End Class