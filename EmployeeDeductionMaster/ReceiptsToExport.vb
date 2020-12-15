Public Class ReceiptsToExport

    Dim loaded As Boolean = False
    Dim decide As String
    Dim exportdate As Date

    Private Sub BankingReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        decide = "All Receipts"
        Try
            loadCompanyAndMode(decide, logedinname)

            For Each col As DataGridViewColumn In DataGridView1.Columns
                col.ReadOnly = True
            Next
            For Each col As DataGridViewColumn In DataGridView2.Columns
                col.ReadOnly = True
            Next
            For Each col As DataGridViewColumn In DataGridView3.Columns
                col.ReadOnly = True
            Next
        Catch ex As Exception

        End Try
        loaded = True
    End Sub

    Private Sub loadCompanyAndMode(ByVal whichreport As String, ByVal capturedby As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = ""
        If whichreport = "All Receipts" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Banked' group by ReceiptsBreakdown.BookID,CompanyName"
        ElseIf whichreport = "My Captured Receipts" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Banked' and BatchName in (select BatchName from Receipts where Status = 'Banked' and CapturedBy = '" & capturedby & "') group by ReceiptsBreakdown.BookID,CompanyName"
        ElseIf whichreport = "Banking Batches" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,Sum(Amount) as TotalAmount,BankingBatch from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Banked' group by ReceiptsBreakdown.BookID,CompanyName,BankingBatch"
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
            MessageBox.Show(ex.Message, "loadCompanyAndMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCompanyAndMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceipts(ByVal ReceiptingBookID As Integer, ByVal whichreport As String, ByVal capturedby As String, ByVal bankingbatch As String)
        DataGridView1.DataSource = Nothing

        Dim qry As String = ""
        If whichreport = "All Receipts" Then
            qry = "select PayerID,Amount,ReceiptType,TransactionDate,Reference,Products as ProductName,ReceiptsBreakdown.ProductID,BankingBatch,ReceiptsBreakdown.AutoID,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "'"
        ElseIf whichreport = "My Captured Receipts" Then
            qry = "select PayerID,Amount,ReceiptType,TransactionDate,Reference,Products as ProductName,ReceiptsBreakdown.ProductID,BankingBatch,ReceiptsBreakdown.AutoID,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "' and BatchName in (select BatchName from Receipts where Status = 'Banked' and CapturedBy = '" & capturedby & "')"
        ElseIf whichreport = "Banking Batches" Then
            qry = "select PayerID,Amount,ReceiptType,TransactionDate,Reference,Products as ProductName,ReceiptsBreakdown.ProductID,BankingBatch,ReceiptsBreakdown.AutoID,Comment,'' as EndColumn from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "' and BankingBatch = '" & bankingbatch & "'"
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

                tbtotalreceipt.Text = 0.0
                For Each row As DataGridViewRow In DataGridView1.Rows
                    tbtotalreceipt.Text = Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim) + Decimal.Parse(row.Cells(1).Value), 2)
                    'MessageBox.Show(tbtotalreceipt.Text, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Next
                Label2.Text = getCurrency(Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2))
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

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to export this list to update systems ?") = DialogResult.Yes Then
                If decision(Label2.Text.Trim & " is the correct amount to be exported ?") = DialogResult.Yes Then
                    If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                        ExportToExcel(DataGridView1, "")
                        exportdate = Date.Now
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            updateReceiptStatusByBankingBatch("Exported", row.Cells(7).Value)
                            updateReceiptBreakdownStatusByBankingBatch("Exported", row.Cells(7).Value)
                            updateReceiptBreakdownExportDetailsByAutoID(logedinname, Date.Now, Integer.Parse(row.Cells(8).Value))
                            wrapper.saveAction("Receipt Exported ", exportdate, logedinname, row.Cells(7).Value)
                        Next
                        MessageBox.Show("Done Updating Details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        loadCompanyAndMode(decide, logedinname)

                        DataGridView1.DataSource = Nothing
                        BindingNavigator1.BindingSource = Nothing

                        DataGridView3.DataSource = Nothing
                        BindingNavigator3.BindingSource = Nothing
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub loadReference(ByVal ReceiptingBookID As String, ByVal whichreport As String, ByVal capturedby As String, ByVal bankingbatch As String)
        DataGridView3.DataSource = Nothing

        Dim qry As String = ""
        If whichreport = "All Receipts" Then
            qry = "select Reference,SUM(Amount) from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "' group by Reference"
        ElseIf whichreport = "My Captured Receipts" Then
            qry = "select Reference,SUM(Amount) from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "' and BatchName in (select BatchName from Receipts where Status = 'Banked' and CapturedBy = '" & capturedby & "') group by Reference"
        ElseIf whichreport = "Banking Batches" Then
            qry = "select Reference,SUM(Amount) from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID where ReceiptsBreakdown.Status = 'Banked' and BookID = '" & ReceiptingBookID & "' and BankingBatch = '" & bankingbatch & "' group by Reference"
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
            MessageBox.Show(ex.Message, "loadReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReference()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub


    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        If DataGridView2.Rows.Count > 0 Then
            If decide = "All Receipts" Then
                loadReceipts(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, "")
                loadReference(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, "")
            ElseIf decide = "My Captured Receipts" Then
                loadReceipts(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, "")
                loadReference(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, "")
            ElseIf decide = "Banking Batches" Then
                loadReceipts(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, DataGridView2.CurrentRow.Cells(3).Value)
                loadReference(Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), decide, logedinname, DataGridView2.CurrentRow.Cells(3).Value)
            End If
        End If
    End Sub

    Private Sub cball_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cball.CheckedChanged
        If cball.Checked = True Then
            cball.Checked = True
            cbmy.Checked = False
            cbbatches.Checked = False
            decide = "All Receipts"
        ElseIf cball.Checked = False Then
            cball.Checked = False
            cbmy.Checked = False
            cbbatches.Checked = False
            'decide = "My Captured Receipts"
        End If

        If loaded = True Then

            loadCompanyAndMode(decide, logedinname)

            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing

            DataGridView3.DataSource = Nothing
            BindingNavigator3.BindingSource = Nothing
            Label2.Text = "0.00"
        End If
    End Sub

    Private Sub cbmy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbmy.CheckedChanged
        If cbmy.Checked = True Then
            cball.Checked = False
            cbmy.Checked = True
            cbbatches.Checked = False
            decide = "My Captured Receipts"
        ElseIf cbmy.Checked = False Then
            cball.Checked = False
            cbmy.Checked = False
            cbbatches.Checked = False
            'decide = "All Receipts"
        End If

        If loaded = True Then

            loadCompanyAndMode(decide, logedinname)

            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing

            DataGridView3.DataSource = Nothing
            BindingNavigator3.BindingSource = Nothing
            Label2.Text = "0.00"
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If decision("Are you sure to proceed ?") = DialogResult.Yes Then
            ExportToExcel(DataGridView1, "")
        End If
    End Sub

    Private Sub cbbatches_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbbatches.CheckedChanged
        If cbbatches.Checked = True Then
            cball.Checked = False
            cbmy.Checked = False
            cbbatches.Checked = True
            decide = "Banking Batches"
        ElseIf cbbatches.Checked = False Then
            cball.Checked = False
            cbmy.Checked = False
            cbbatches.Checked = False
            'decide = "My Captured Receipts"
        End If

        If loaded = True Then

            loadCompanyAndMode(decide, logedinname)

            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing

            DataGridView3.DataSource = Nothing
            BindingNavigator3.BindingSource = Nothing
            Label2.Text = "0.00"
        End If

    End Sub
End Class