Imports DgvFilterPopup

Public Class TransferReceiptBetweenPaymentModes

    Dim batchname As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadReceipts(tbpayerid.Text.Trim, "Per Payer")
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
    End Sub

    Private Sub loadReceipts(ByVal PayerID As String, ByVal howmany As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = ""
        If howmany = "All" Then
            qry = "select PayerID,Amount,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName,Receipts.Status,CompanyName,Receipts.BookID,AutoID,Receipts.ModeID from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID inner join CompanyBooks on Receipts.BookID = CompanyBooks.BookID"
        Else
            qry = "select PayerID,Amount,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName,Receipts.Status,CompanyName,Receipts.BookID,AutoID,Receipts.ModeID from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID inner join CompanyBooks on Receipts.BookID = CompanyBooks.BookID where PayerID = '" & PayerID & "'"
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
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView1)
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
            tbreceiptid.Text = DataGridView1.CurrentRow.Cells(10).Value
            batchname = DataGridView1.CurrentRow.Cells(6).Value
            tbpaymentmodefrom.Text = DataGridView1.CurrentRow.Cells(3).Value
            tbpaymentmodeidfrom.Text = DataGridView1.CurrentRow.Cells(11).Value
        End If
    End Sub

    Private Sub loadReceiptsBreakdown(ByVal batchname As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,ReceiptNo,ReceiptType,TransactionDate,Products,BatchName,Comment,CompanyName,ReceiptsBreakdown.BookID,ReceiptsBreakdown.ModeID,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID  where BatchName = '" & batchname & "'"

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
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView2)

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
        If tbpaymentmodeidto.Text.Trim = "" Then
            MessageBox.Show("Select the company book", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbreceiptid.Text.Trim = "" Then
                MessageBox.Show("Select the receipt first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to transfer the receipt ?") = DialogResult.Yes Then
                    updateReceiptModeID(Integer.Parse(tbpaymentmodeidto.Text.Trim), Integer.Parse(tbreceiptid.Text.Trim))
                    updateReceiptBreakdownModeID(Integer.Parse(tbpaymentmodeidto.Text.Trim), batchname)
                    Close()
                End If
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        tbpaymentmodeto.Text = ""
        tbpaymentmodeidto.Text = ""

        Dim vrl As New PaymentModesForTransfer
        vrl.MdiParent = mform
        vrl.trbpm = Me
        vrl.Show()
    End Sub

End Class