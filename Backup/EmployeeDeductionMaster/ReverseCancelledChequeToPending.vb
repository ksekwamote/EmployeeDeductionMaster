Public Class ReverseCancelledChequeToRequested

    Private Sub ReverseCancelledChequeToPending_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCancelledCheques(tbcustomerid.Text.Trim)
    End Sub

    Private Sub loadCancelledCheques(ByVal customerid As String)
        DataGridView1.DataSource = Nothing

        Dim qry As String = ""
        If customerid = "" Then
            qry = "select * from ProductsCheques where Status = 'Cancelled'"
        Else
            qry = "select * from ProductsCheques where Status = 'Cancelled' and CustomerID = '" & customerid & "'"
        End If
        Dim ds As New DataSet
        Dim cmd As New SqlClient.SqlCommand(qry, con)
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
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCancelledCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadCancelledCheques(tbcustomerid.Text.Trim)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure to reverse cheque to pending ?") = DialogResult.Yes Then
                updateChequeStatus("Requested", Integer.Parse(DataGridView1.CurrentRow.Cells(0).Value), Date.Now)
                wrapper.saveAction("Cheque Reversed To Requested", Date.Now, logedinname, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value)
                MessageBox.Show("Cheque Reversed To Requested", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                tbcustomerid.Text = ""
                loadCancelledCheques(tbcustomerid.Text.Trim)
            End If
        End If
    End Sub

End Class