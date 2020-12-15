Public Class ApproveReceiptBatch

    Dim batchname As String

    Private Sub ApproveReceiptBatch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadReceipts(branchid, logedinname)
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.ReadOnly = True
        Next
        For Each col As DataGridViewColumn In DataGridView2.Columns
            col.ReadOnly = True
        Next
    End Sub

    Private Sub loadReceipts(ByVal CompanyBranchID As Integer, ByVal capturedby As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'Batch' and CompanyBranchID = '" & CompanyBranchID & "' and CapturedBy = '" & capturedby & "'"

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

                tbbatchtotal.Text = 0.0
                For Each row As DataGridViewRow In DataGridView1.Rows
                    tbbatchtotal.Text = Math.Round(Decimal.Parse(tbbatchtotal.Text.Trim) + Decimal.Parse(row.Cells(1).Value), 2)
                Next
                Label1.Text = getCurrency(Math.Round(Decimal.Parse(tbbatchtotal.Text.Trim), 2))
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

    Private Sub loadReceiptsBreakdown(ByVal batchname As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,ReceiptNo,ReceiptType,TransactionDate,Products,BatchName,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Batch' and BatchName = '" & batchname & "'"

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

                tbtotalreceipt.Text = 0.0
                For Each row As DataGridViewRow In DataGridView2.Rows
                    tbtotalreceipt.Text = Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim) + Decimal.Parse(row.Cells(1).Value), 2)
                Next
                Label2.Text = getCurrency(Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2))
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            batchname = DataGridView1.CurrentRow.Cells(7).Value
            loadReceiptsBreakdown(DataGridView1.CurrentRow.Cells(7).Value)
        End If
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = tbtotalreceipt.Text
    End Sub

    Private Sub tbbatchtotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbbatchtotal.TextChanged
        Label1.Text = tbbatchtotal.Text
    End Sub

    Private Sub btdecline_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btdecline.Click
        If DataGridView2.Rows.Count > 0 Then
            If decision("Are you sure you want to DECLINE the receipt ?") = DialogResult.Yes Then
                updateReceiptStatusByBatchName("Declined", batchname)
                updateReceiptBreakdownStatusByBatchName("Declined", batchname)
                loadReceipts(branchid, logedinname)

                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing

                tbtotalreceipt.Text = "0.00"
                wrapper.saveAction("Receipt Declined", Date.Now, logedinname, batchname)
            End If
        End If
    End Sub

    Private Sub btconfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btconfirm.Click
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to Confirm All the receipt as a batch for approval ?") = DialogResult.Yes Then

                Dim vrl As New ReceiptAuthorization
                vrl.MdiParent = mform
                vrl.arb = Me
                vrl.whichside = "Batch"
                vrl.Show()
            End If
        End If
    End Sub

    Public Sub proceedToConfirm(ByVal authorizer As String)
        For Each row As DataGridViewRow In DataGridView1.Rows
            updateReceiptStatusByBatchName("New", row.Cells(7).Value)
            updateReceiptBreakdownStatusByBatchName("New", row.Cells(7).Value)
            updateReceiptCapturedByDetails(authorizer, Date.Now, row.Cells(7).Value)
        Next
        loadReceipts(branchid, logedinname)

        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing

        tbbatchtotal.Text = "0.00"
        wrapper.saveAction("Receipt Confirmed For Approval", Date.Now, logedinname, batchname)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class