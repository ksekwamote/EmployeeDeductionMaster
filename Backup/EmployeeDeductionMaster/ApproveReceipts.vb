Public Class ApproveReceipts

    Dim batchname As String
    Dim bookid, modeid As Integer

    Public whichreport As String

    Private Sub ApproveReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCompanyBatches(branchid, whichreport, logedinname)
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.ReadOnly = True
        Next

        For Each col As DataGridViewColumn In DataGridView2.Columns
            col.ReadOnly = True
        Next

        For Each col As DataGridViewColumn In DataGridView3.Columns
            col.ReadOnly = True
        Next

    End Sub

    Private Sub loadCompanyBatches(ByVal CompanyBranchID As Integer, ByVal whichone As String, ByVal logedin As String)
        DataGridView3.DataSource = Nothing
        Dim qry As String = ""

        If whichone = "New Receipts" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'New' and CompanyBranchID = '" & CompanyBranchID & "' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName order by ReceiptsBreakdown.BookID"
        ElseIf whichone = "My Captured Receipts" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'New' and CompanyBranchID = '" & CompanyBranchID & "'  and BatchName in (select BatchName from Receipts where Status = 'New' and CapturedBy = '" & logedinname & "') group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName order by ReceiptsBreakdown.BookID"
            'qry = "select ReceiptsBreakdown.BookID,CompanyName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID where ReceiptsBreakdown.Status = 'New' and CompanyBranchID = '" & CompanyBranchID & "' and BatchName in (select BatchName from Receipts where Status = 'New' and CapturedBy = '" & logedinname & "') group by ReceiptsBreakdown.BookID,CompanyName"
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
                DataGridView3.DataSource = bs
                BindingNavigator3.BindingSource = bs
                DataGridView3.Columns(DataGridView3.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            Else
                DataGridView3.DataSource = Nothing
                BindingNavigator3.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCompanyBatches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCompanyBatches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceipts(ByVal ReceiptingBookID As Integer, ByVal CompanyBranchID As Integer, ByVal ModeID As Integer, ByVal whichone As String, ByVal capturedby As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = ""

        If whichone = "New Receipts" Then
            qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'New' and BookID = '" & ReceiptingBookID & "' and CompanyBranchID = '" & CompanyBranchID & "' and ModeOfPayment.ModeID = '" & ModeID & "'"
        ElseIf whichone = "My Captured Receipts" Then
            qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'New' and BookID = '" & ReceiptingBookID & "' and CompanyBranchID = '" & CompanyBranchID & "' and ModeOfPayment.ModeID = '" & ModeID & "' and CapturedBy = '" & capturedby & "'"
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

    Private Sub loadReceiptsBreakdown(ByVal batchname As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,ReceiptNo,ReceiptType,TransactionDate,Products,BatchName,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'New' and BatchName = '" & batchname & "'"

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btdecline.Click
        If DataGridView2.Rows.Count > 0 Then
            If decision("Are you sure you want to DECLINE the receipt ?") = DialogResult.Yes Then
                updateReceiptStatusByBatchName("Declined", batchname)
                updateReceiptBreakdownStatusByBatchName("Declined", batchname)
                loadReceipts(bookid, branchid, modeid, whichreport, logedinname)
                If DataGridView1.Rows.Count > 0 Then
                    DataGridView2.DataSource = Nothing
                    BindingNavigator2.BindingSource = Nothing
                Else
                    loadCompanyBatches(branchid, whichreport, logedinname)
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    DataGridView2.DataSource = Nothing
                    BindingNavigator2.BindingSource = Nothing
                End If
                tbtotalreceipt.Text = "0.00"
                wrapper.saveAction("Receipt Declined", Date.Now, logedinname, batchname)
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btapprove.Click
        If DataGridView2.Rows.Count > 0 Then
            If decision("Are you sure you want to APPROVE the receipt ?") = DialogResult.Yes Then
                'If checkIfSameDayBanking(modeid) = "YES" Then
                'If checkIfThereIsApprovedPointOfSalesForPreviousDate(modeid, Date.Now) = True Then
                'MessageBox.Show("Bank the previous Point Of Sales receipts first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Else
                'approve()
                'End If
                'Else
                'approve()
                'End If
                approve()
            End If
        End If
    End Sub

    Public Sub approve()
        updateReceiptStatusByBatchName("Approved", batchname)
        updateReceiptBreakdownStatusByBatchName("Approved", batchname)
        loadReceipts(bookid, branchid, modeid, whichreport, logedinname)
        If DataGridView1.Rows.Count > 0 Then
            DataGridView2.DataSource = Nothing
            BindingNavigator2.BindingSource = Nothing
        Else
            loadCompanyBatches(branchid, whichreport, logedinname)
            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing
            DataGridView2.DataSource = Nothing
            BindingNavigator2.BindingSource = Nothing
        End If
        tbtotalreceipt.Text = "0.00"
        wrapper.saveAction("Receipt Approved", Date.Now, logedinname, batchname)
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = tbtotalreceipt.Text
    End Sub

    Private Sub DataGridView3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.Click
        If DataGridView3.Rows.Count > 0 Then
            bookid = Integer.Parse(DataGridView3.CurrentRow.Cells(0).Value)
            modeid = Integer.Parse(DataGridView3.CurrentRow.Cells(2).Value)
            loadReceipts(bookid, branchid, modeid, whichreport, logedinname)
            DataGridView2.DataSource = Nothing
            BindingNavigator2.BindingSource = Nothing
            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
        End If
    End Sub

End Class