Public Class BankingReceipts

    Dim bankingbatch, mode, company As String
    Dim modeid As Integer
    Dim decided, per, confirmedby As String

    Private Sub BankingReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadApprovers(branchid, decided)
        loadCompanyAndMode(branchid, decided, per, confirmedby)

        cbcompany.Text = getCompanyBranch(branchid) & " receipts"
        decided = "Company"
        confirmedby = "Book"
    End Sub

    Private Sub loadApprovers(ByVal CompanyBranchID As Integer, ByVal decide As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = ""
        If decide = "Company" Then
            qry = "select DISTINCT(WhoDidIt) from UserActions where WhatIsDone = 'Receipt Approved' and CustomerID in (select BatchName from Receipts where Status = 'Approved' and CompanyBranchID = '" & CompanyBranchID & "')"
        ElseIf decide = "All" Then
            qry = "select DISTINCT(WhoDidIt) from UserActions where WhatIsDone = 'Receipt Approved' and CustomerID in (select BatchName from Receipts where Status = 'Approved')"
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

            For Each col As DataGridViewColumn In DataGridView2.Columns
                col.ReadOnly = True
            Next

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            'MessageBox.Show(ex.Message, "loadApprovers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'MessageBox.Show(ex.StackTrace, "loadApprovers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadCompanyAndMode(ByVal CompanyBranchID As Integer, ByVal decide As String, ByVal per As String, ByVal approvedby As String)
        DataGridView3.DataSource = Nothing
        Dim qry As String = ""
        If decide = "Company" Then
            If per = "Book" Then
                qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Approved' and CompanyBranchID = '" & CompanyBranchID & "' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName"
            ElseIf per = "ApproverAndBook" Then
                qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Approved' and CompanyBranchID = '" & CompanyBranchID & "' and BatchName in (select CustomerID from UserActions where WhatIsDone = 'Receipt Approved' and WhoDidIt = '" & approvedby & "') group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName"
            End If
        ElseIf decide = "All" Then
            If per = "Book" Then
                qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Approved' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName"
            ElseIf per = "ApproverAndBook" Then
                qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount from ReceiptsBreakdown inner join Products on ReceiptsBreakdown.ProductID = Products.AutoID inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Approved' and BatchName in (select CustomerID from UserActions where WhatIsDone = 'Receipt Approved' and WhoDidIt = '" & approvedby & "') group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName"
            End If
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

            For Each col As DataGridViewColumn In DataGridView3.Columns
                col.ReadOnly = True
            Next

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            'MessageBox.Show(ex.Message, "loadCompanyAndMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'MessageBox.Show(ex.StackTrace, "loadCompanyAndMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceipts(ByVal modeid As Integer, ByVal ReceiptingBookID As Integer, ByVal CompanyBranchID As Integer, ByVal decide As String, ByVal per As String, ByVal approvedby As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = ""
        If decide = "Company" Then
            If per = "Book" Then
                qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'Approved' and Receipts.ModeID = '" & modeid & "' and  BookID = '" & ReceiptingBookID & "' and CompanyBranchID = '" & CompanyBranchID & "'"
            ElseIf per = "ApproverAndBook" Then
                qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'Approved' and Receipts.ModeID = '" & modeid & "' and  BookID = '" & ReceiptingBookID & "' and CompanyBranchID = '" & CompanyBranchID & "' and BatchName in (select CustomerID from UserActions where WhatIsDone = 'Receipt Approved' and WhoDidIt = '" & approvedby & "')"
            End If
        ElseIf decide = "All" Then
            If per = "Book" Then
                qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'Approved' and Receipts.ModeID = '" & modeid & "' and  BookID = '" & ReceiptingBookID & "'"
            ElseIf per = "ApproverAndBook" Then
                qry = "select PayerID,Amount,ReceiptNo,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status = 'Approved' and Receipts.ModeID = '" & modeid & "' and  BookID = '" & ReceiptingBookID & "' and BatchName in (select CustomerID from UserActions where WhatIsDone = 'Receipt Approved' and WhoDidIt = '" & approvedby & "')"
            End If
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

            For Each col As DataGridViewColumn In DataGridView1.Columns
                col.ReadOnly = True
            Next

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btbanking.Click
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to bank the receipts?") = DialogResult.Yes Then
                If decision("Total Receipt to be banked is " & Label2.Text & "?") = DialogResult.Yes Then
                    bankingbatch = company & ":" & mode & ":" & Date.Now

                    If checkIfConfirmationNotNeeded(modeid) = "YES" Then
                        If decision("Is the Transaction Date correct ? ! " & DateTimePicker1.Value.Day & " " & MonthName(DateTimePicker1.Value.Month) & " " & DateTimePicker1.Value.Year & " !") = DialogResult.Yes Then
                            For Each row As DataGridViewRow In DataGridView1.Rows
                                updateReceiptBankingDetails(row.Cells(7).Value, "Banked", Date.Now, logedinname, bankingbatch)
                                updateReceiptBreakdownBankingDetails(row.Cells(7).Value, "Banked", bankingbatch)
                                updateReceiptReferenceForNoConfirmations(row.Cells(2).Value, row.Cells(7).Value)
                                updateReceiptBreakdownReferenceForNoConfirmations(row.Cells(2).Value, row.Cells(7).Value)
                            Next
                            wrapper.saveAction("Total Receipt Confirmed To Be Banked", Date.Now, logedinname, bankingbatch)
                            'MessageBox.Show("Done Updating Details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            MessageBox.Show("Nothing Banked", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    Else
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            updateReceiptBankingDetails(row.Cells(7).Value, "Banking", Date.Now, logedinname, bankingbatch)
                            updateReceiptBreakdownBankingDetails(row.Cells(7).Value, "Banking", bankingbatch)
                        Next
                        wrapper.saveAction("Total Receipt Confirmed To Be Banked", Date.Now, logedinname, bankingbatch)
                        'MessageBox.Show("Done Updating Details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        Dim vrl As New BankingReport
                        vrl.MdiParent = mform
                        vrl.br = Me
                        vrl.BankingBatch = bankingbatch
                        vrl.Show()
                    End If

                    loadCompanyAndMode(branchid, decided, per, confirmedby)
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    tbtotalreceipt.Text = "0.0"
                End If
            End If
        End If
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = tbtotalreceipt.Text
    End Sub

    Private Sub DataGridView3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.Click
        If DataGridView3.Rows.Count > 0 Then
            loadReceipts(Integer.Parse(DataGridView3.CurrentRow.Cells(2).Value), Integer.Parse(DataGridView3.CurrentRow.Cells(0).Value), branchid, decided, per, confirmedby)
            GroupBox1.Enabled = True
            modeid = Integer.Parse(DataGridView3.CurrentRow.Cells(2).Value)
            mode = DataGridView3.CurrentRow.Cells(3).Value
            company = DataGridView3.CurrentRow.Cells(1).Value
            If checkIfConfirmationNotNeeded(modeid) = "YES" Then
                Label1.Enabled = True
                DateTimePicker1.Enabled = True
            Else
                Label1.Enabled = False
                DateTimePicker1.Enabled = False
            End If
        End If
    End Sub

    Private Sub cbcompany_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbcompany.CheckedChanged
        If cbcompany.Checked = True Then
            cbcompany.Checked = True
            cball.Checked = False
            decided = "Company"
        ElseIf cbcompany.Checked = False Then
            cbcompany.Checked = False
            cball.Checked = True
            decided = "All"
        End If
        loadCompanyAndMode(branchid, decided, per, confirmedby)
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
    End Sub

    Private Sub cball_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cball.CheckedChanged
        If cball.Checked = True Then
            cbcompany.Checked = False
            cball.Checked = True
            decided = "All"
        ElseIf cball.Checked = False Then
            cbcompany.Checked = True
            cball.Checked = False
            decided = "Company"
        End If
        loadCompanyAndMode(branchid, decided, per, confirmedby)
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub cbapproverandbook_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbapproverandbook.CheckedChanged
        If cbapproverandbook.Checked = True Then
            cbapproverandbook.Checked = True
            cbbook.Checked = False
            per = "ApproverAndBook"
            GroupBox2.Enabled = True
        ElseIf cbapproverandbook.Checked = False Then
            cbbook.Checked = True
            cbapproverandbook.Checked = False
            per = "Book"
            GroupBox2.Enabled = False
        End If

        loadApprovers(branchid, decided)

        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
        DataGridView3.DataSource = Nothing
        BindingNavigator3.BindingSource = Nothing
    End Sub

    Private Sub cbbook_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbbook.CheckedChanged
        If cbbook.Checked = True Then
            cbapproverandbook.Checked = False
            cbbook.Checked = True
            per = "Book"
            GroupBox2.Enabled = False
        ElseIf cbbook.Checked = False Then
            cbbook.Checked = False
            cbapproverandbook.Checked = True
            per = "ApproverAndBook"
            GroupBox2.Enabled = True
        End If

        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing

        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
        DataGridView3.DataSource = Nothing
        BindingNavigator3.BindingSource = Nothing
    End Sub

    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        If DataGridView2.Rows.Count > 0 Then
            confirmedby = DataGridView2.CurrentRow.Cells(0).Value
            loadCompanyAndMode(branchid, decided, per, confirmedby)
            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing
        Else
            confirmedby = ""
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure to cancel the receipt !") = DialogResult.Yes Then
                updateReceiptStatusByBatchName("Declined", DataGridView1.CurrentRow.Cells(7).Value)
                updateReceiptBreakdownStatusByBatchName("Declined", DataGridView1.CurrentRow.Cells(7).Value)
                wrapper.saveAction("Receipt Declined", Date.Now, logedinname, DataGridView1.CurrentRow.Cells(7).Value)

                loadCompanyAndMode(branchid, decided, per, confirmedby)
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
                tbtotalreceipt.Text = "0.0"

            End If
        End If
    End Sub
End Class