Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class ConfirmBanking

    Dim bankingbatch As String
    Public choice As String

    Private Sub BankingReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadReceiptsBankingBatchesAll(choice)

        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.ReadOnly = True
        Next

        For Each col As DataGridViewColumn In DataGridView2.Columns
            col.ReadOnly = True
        Next
    End Sub

    Private Sub loadReceiptsBankingBatches(ByVal whichone As String, ByVal BookID As Integer)
        DataGridView2.DataSource = Nothing
        Dim qry As String = ""
        If whichone = "All" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount,BankingBatch from ReceiptsBreakdown inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status not in ('New','Declined') and CompanyBooks.BookID = '" & BookID & "' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,BankingBatch"
        ElseIf whichone = "Banking" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount,BankingBatch from ReceiptsBreakdown inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Banking' and CompanyBooks.BookID = '" & BookID & "' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,BankingBatch"
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
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView2)
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsBankingBatches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsBankingBatches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceiptsBankingBatchesAll(ByVal whichone As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = ""
        If whichone = "All" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount,BankingBatch from ReceiptsBreakdown inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status not in ('New','Declined') group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,BankingBatch"
        ElseIf whichone = "Banking" Then
            qry = "select ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,Sum(Amount) as TotalAmount,BankingBatch from ReceiptsBreakdown inner join CompanyBooks on ReceiptsBreakdown.BookID = CompanyBooks.BookID inner join ModeOfPayment on ReceiptsBreakdown.ModeID = ModeOfPayment.ModeID where ReceiptsBreakdown.Status = 'Banking' group by ReceiptsBreakdown.BookID,CompanyName,ReceiptsBreakdown.ModeID,ModeName,BankingBatch"
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
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView2)
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsBankingBatchesAll()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsBankingBatchesAll()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceipts(ByVal bankingbatch As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select PayerID,Amount,TransactionDate,ModeName,CapturedBy,CapturedTime,BatchName,Receipts.Status from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where BankingBatch = '" & bankingbatch & "'"

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataGridView1.Rows.Count > 0 Then
            If tbreference.Text.Trim = "" Then
                MessageBox.Show("Enter reference first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure you want to confirm banking?") = DialogResult.Yes Then
                    If decision("Total Receipt Banked is " & Label2.Text & "?") = DialogResult.Yes Then
                        If decision("Is the Transaction Date correct ? ! " & dtptransactiondate.Value.Day & " " & MonthName(dtptransactiondate.Value.Month) & " " & dtptransactiondate.Value.Year & " !") = DialogResult.Yes Then
                            updateReceiptReference(tbreference.Text.Trim, bankingbatch, dtptransactiondate.Value)
                            updateReceiptBreakdownReference(tbreference.Text.Trim, bankingbatch, dtptransactiondate.Value)

                            wrapper.saveAction("Total Receipt Confirmed Banked", Date.Now, logedinname, bankingbatch)
                            MessageBox.Show("Done Updating Details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            If tbbookid.Text.Trim = "" Then
                                loadReceiptsBankingBatchesAll(choice)
                            Else
                                loadReceiptsBankingBatches(choice, Integer.Parse(tbbookid.Text.Trim))
                            End If
                            DataGridView1.DataSource = Nothing
                            BindingNavigator1.BindingSource = Nothing
                            tbreference.Text = ""
                            tbtotalreceipt.Text = "0.00"
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = tbtotalreceipt.Text
    End Sub

    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        If DataGridView2.Rows.Count > 0 Then
            bankingbatch = DataGridView2.CurrentRow.Cells(5).Value
            loadReceipts(DataGridView2.CurrentRow.Cells(5).Value)
        End If
    End Sub

    Private Sub tbbookid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbbookid.TextChanged
        If tbbookid.Text.Trim = "" Then

        Else
            loadReceiptsBankingBatches(choice, Integer.Parse(tbbookid.Text.Trim))
            DataGridView1.DataSource = Nothing
            BindingNavigator1.BindingSource = Nothing
            tbreference.Text = ""
            tbtotalreceipt.Text = "0.00"

            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim vrl As New CompanyBooksList
        vrl.MdiParent = mform
        vrl.whichform = "confirm banking"
        vrl.cb = Me
        vrl.Show()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        loadReceiptsBankingBatchesAll(choice)
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
        tbreference.Text = ""
        tbtotalreceipt.Text = "0.00"

        tbbookid.Text = ""
        tbbook.Text = ""

        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to DECLINE banking?") = DialogResult.Yes Then
                If decision("Total Receipt Banked is " & Label2.Text & "?") = DialogResult.Yes Then
                    updateReceiptStatusByBankingBatch("Declined", bankingbatch)
                    updateReceiptBreakdownStatusByBankingBatch("Declined", bankingbatch)

                    wrapper.saveAction("Batch Declined", Date.Now, logedinname, bankingbatch)
                    MessageBox.Show("Done Declining", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If tbbookid.Text.Trim = "" Then
                        loadReceiptsBankingBatchesAll(choice)
                    Else
                        loadReceiptsBankingBatches(choice, Integer.Parse(tbbookid.Text.Trim))
                    End If
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    tbreference.Text = ""
                    tbtotalreceipt.Text = "0.00"
                End If
            End If
        Else
            MessageBox.Show("Select Batch To Decline", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to REVERSE BACK TO APPROVAL?") = DialogResult.Yes Then
                If decision("Total Receipt Banked is " & Label2.Text & "?") = DialogResult.Yes Then
                    updateReceiptStatusByBankingBatch("New", bankingbatch)
                    updateReceiptBreakdownStatusByBankingBatch("NEw", bankingbatch)

                    wrapper.saveAction("Batch Reversed To Approval", Date.Now, logedinname, bankingbatch)
                    MessageBox.Show("Done Reversing Back To Approval", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If tbbookid.Text.Trim = "" Then
                        loadReceiptsBankingBatchesAll(choice)
                    Else
                        loadReceiptsBankingBatches(choice, Integer.Parse(tbbookid.Text.Trim))
                    End If
                    DataGridView1.DataSource = Nothing
                    BindingNavigator1.BindingSource = Nothing
                    tbreference.Text = ""
                    tbtotalreceipt.Text = "0.00"
                End If
            End If
        Else
            MessageBox.Show("Select Batch To Reverse Back To Approval", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class