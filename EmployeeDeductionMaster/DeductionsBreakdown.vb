Public Class DeductionsBreakdown

    Private Sub DeductionsBreakdown_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub loadActualDeductions(ByVal customerid As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select CustomerID,ActualDeduction as Actual,'' as Expected,ForMonth,TransactionDate,'' Comment from ActualDeductions where CustomerID = '" & customerid & "'"

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
                For count As Integer = 0 To dt.Rows.Count - 1
                    dt.Rows(count).Item(2) = getExpectedDeductionsPerProductInDatabase(dt.Rows(count).Item(0).ToString, Date.Parse(dt.Rows(count).Item(3).ToString))
                    If Decimal.Parse(dt.Rows(count).Item(1).ToString) = Decimal.Parse(dt.Rows(count).Item(2).ToString) Then
                        dt.Rows(count).Item(5) = "Received As Expected"
                    Else
                        dt.Rows(count).Item(5) = ""
                    End If
                Next
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
            MessageBox.Show(ex.Message, "loadActualDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadActualDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadExpectedDeductions(ByVal customerid As String, ByVal formonth As Date)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select CustomerID,Amount,ForMonth,Products from ExpectedDeductionsPerProduct inner join Products on ExpectedDeductionsPerProduct.ProductID = Products.AutoID where CustomerID = '" & customerid & "' and ForMonth = '" & formonth & "'"

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
                DataGridView2.DataSource = bs
                BindingNavigator2.BindingSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadExpectedDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadActualDeductions(TextBox1.Text.Trim)
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            loadExpectedDeductions(DataGridView1.CurrentRow.Cells(0).Value, Date.Parse(DataGridView1.CurrentRow.Cells(3).Value))
            tbexpected.Text = getExpectedDeductionsPerProductInDatabase(DataGridView1.CurrentRow.Cells(0).Value, Date.Parse(DataGridView1.CurrentRow.Cells(3).Value))
            tbactual.Text = DataGridView1.CurrentRow.Cells(1).Value
        End If
    End Sub
End Class