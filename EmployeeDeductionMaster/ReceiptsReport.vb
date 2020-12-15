Public Class ReceiptsReport

    Dim receiptstatus As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadReceipts(tbpayerid.Text.Trim, "Per Payer")
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
    End Sub

    Private Sub loadReceipts(ByVal PayerID As String, ByVal howmany As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = ""
        If howmany = "All" Then
            qry = "select PayerID,Amount,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName,Receipts.Status from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID"
        Else
            qry = "select PayerID,Amount,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName,Receipts.Status from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where PayerID = '" & PayerID & "'"
        End If

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        loadReceipts(tbpayerid.Text.Trim, "All")
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            loadReceiptsBreakdown(DataGridView1.CurrentRow.Cells(6).Value)
            tbbatchname.Text = DataGridView1.CurrentRow.Cells(6).Value
            receiptstatus = DataGridView1.CurrentRow.Cells(7).Value
            If receiptstatus = "Approved" Then
                Button3.Enabled = True
                Button4.Enabled = False
            ElseIf receiptstatus = "Declined" Then
                Button3.Enabled = False
                Button4.Enabled = True
            End If
        End If
    End Sub

    Private Sub loadReceiptsBreakdown(ByVal batchname As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,ReceiptNo,ReceiptType,TransactionDate,Products,BatchName,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where BatchName = '" & batchname & "'"

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

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = tbtotalreceipt.Text
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If tbbatchname.Text.Trim = "" Then
            MessageBox.Show("Select the receipt first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If receiptstatus = "Approved" Or receiptstatus = "Banking" Then
                If decision("Are you sure to decline the receipt ?") = DialogResult.Yes Then
                    updateReceiptStatusByBatchName("Declined", tbbatchname.Text.Trim)
                    updateReceiptBreakdownStatusByBatchName("Declined", tbbatchname.Text.Trim)

                    tbbatchname.Text = ""
                    tbtotalreceipt.Text = 0
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    DataGridView2.DataSource = Nothing
                    BindingNavigator2.BindingSource = Nothing
                End If
            Else
                MessageBox.Show("Cannot decline a receipt which is not in the status of Approved/ Banking", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub ReceiptsReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If tbbatchname.Text.Trim = "" Then
            MessageBox.Show("Select the receipt first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If receiptstatus = "Declined" Then
                If decision("Are you sure to revserse declined receipt ?") = DialogResult.Yes Then
                    updateReceiptStatusByBatchName("New", tbbatchname.Text.Trim)
                    updateReceiptBreakdownStatusByBatchName("New", tbbatchname.Text.Trim)

                    tbbatchname.Text = ""
                    tbtotalreceipt.Text = 0
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    DataGridView2.DataSource = Nothing
                    BindingNavigator2.BindingSource = Nothing
                End If
            Else
                MessageBox.Show("Cannot decline a receipt which is not in the status of Declined", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If tbbatchname.Text.Trim = "" Then
            MessageBox.Show("Select the receipt first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If receiptstatus = "Approved" Then
                If decision("Are you sure to reverse receipt to approval ?") = DialogResult.Yes Then
                    updateReceiptStatusByBatchName("New", tbbatchname.Text.Trim)
                    updateReceiptBreakdownStatusByBatchName("New", tbbatchname.Text.Trim)

                    tbbatchname.Text = ""
                    tbtotalreceipt.Text = 0
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    DataGridView2.DataSource = Nothing
                    BindingNavigator2.BindingSource = Nothing
                End If
            Else
                MessageBox.Show("Cannot reverse receipt which is not in the status of Approved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DataGridView2.Rows.Count > 0 Then
            If decision("Are you sure you want to Re-Print Receipt?") = DialogResult.Yes Then
                Dim vrl As New RePrintReceipts
                vrl.MdiParent = mform
                vrl.BatchName = tbbatchname.Text.Trim
                vrl.rr = Me
                vrl.Show()
            End If
        Else
            MessageBox.Show("Select Batch To Reverse Back To Approval", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class