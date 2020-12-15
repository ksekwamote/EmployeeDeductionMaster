Public Class UpdateReceiptNo

    Private Sub loadReceipts(ByVal PayerID As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select ReceiptNo,PayerID,Amount,TransactionDate,CapturedBy,CapturedTime,BatchName,Status from Receipts where PayerID = '" & PayerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceipts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbreceiptno.Text.Trim = "" Or tbbatchname.Text.Trim = "" Then
            MessageBox.Show("Enter all the details first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
                updateReceiptNumberForReceipts(tbreceiptno.Text.Trim, tbbatchname.Text.Trim)
                updateReceiptNumberForReceiptsBreakdown(tbreceiptno.Text.Trim, tbbatchname.Text.Trim)
                MessageBox.Show("Done updating receipt numner", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        loadReceipts(tbpayerid.Text.Trim)
    End Sub

    Private Sub DataGridView1_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbreceiptno.Text = DataGridView1.CurrentRow.Cells(0).Value
            tbbatchname.Text = DataGridView1.CurrentRow.Cells(6).Value
        End If
    End Sub
End Class