Public Class PotentialRefundsToStop

    Dim dsrmrefunds As New DataSet
    Dim index As Integer = 0
    Dim go As Boolean

    Private Function getRMRefunds(ByVal paymentdate2 As Date, ByVal con As SqlClient.SqlConnection) As DataSet
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,CustomerID,Amount,RefundBy,Reference,PaymentDate,Status from CustomerRefunds where Status not in ('New','Declined','Reversed') and PaymentDate <= '" & paymentdate2 & "' and AutoID in (select RefundID from CustomerRefundsStoped)"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Private Function getRMRefundsByEmployerID(ByVal paymentdate2 As Date, ByVal EmployerCode As String, ByVal con As SqlClient.SqlConnection) As DataSet
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,CustomerID,Amount,RefundBy,Reference,PaymentDate,Status from CustomerRefunds where Status not in ('New','Declined','Reversed') and PaymentDate <= '" & paymentdate2 & "' and AutoID in (select RefundID from CustomerRefundsStoped) and CustomerID in (select Account from CustomerMaster where EmployerCode = '" & EmployerCode & "')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefundsByEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefundsByEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Private Function getRMRefundsByCustomerID(ByVal CustomerID As String, ByVal con As SqlClient.SqlConnection) As DataSet
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,CustomerID,Amount,RefundBy,Reference,PaymentDate,Status from CustomerRefunds where Status not in ('New','Declined','Reversed') and CustomerID = '" & CustomerID & "' and AutoID in (select RefundID from CustomerRefundsStoped)"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getRMRefundsByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getRMRefundsByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Private Function getMLMRefunds(ByVal transactiondate1 As Date, ByVal transactiondate2 As Date, ByVal con As SqlClient.SqlConnection) As DataSet
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select LedgerEntry.AutoID,LedgerEntry.CustomerID,LedgerEntry.Amount,PayBy,LedgerEntry.DepositReference,LedgerEntry.TransactionDate,Refunds.Status from LedgerEntry inner join Refunds on LedgerEntry.AutoItemID = Refunds.AutoID where Reference = 'RE' and Refunds.Status = 'Posted' and LedgerEntry.TransactionDate between '" & transactiondate1 & "' and '" & transactiondate2 & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function


    Private Function getMLMRefundsByCustomerID(ByVal CustomerID As String, ByVal con As SqlClient.SqlConnection) As DataSet
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select LedgerEntry.AutoID,LedgerEntry.CustomerID,LedgerEntry.Amount,PayBy,LedgerEntry.DepositReference,LedgerEntry.TransactionDate,Refunds.Status from LedgerEntry inner join Refunds on LedgerEntry.AutoItemID = Refunds.AutoID where Reference = 'RE' and Refunds.Status = 'Posted' and LedgerEntry.CustomerID = '" & CustomerID & "'"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getMLMRefundsByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getMLMRefundsByCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        'Button1.Enabled = False
        ''Orange
        dsrmrefunds = getRMRefunds(DateTimePicker2.Value, conRMOrange)

        ProgressBar1.Maximum = dsrmrefunds.Tables(0).Rows.Count
        For count As Integer = 0 To dsrmrefunds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            index = DataGridView1.Rows.Count - 1
            If checkStoppedRefund(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString, Decimal.Parse(dsrmrefunds.Tables(0).Rows(count).Item(2).ToString), "RM Orange", dsrmrefunds.Tables(0).Rows(count).Item(1).ToString) = False Then
                DataGridView1.Rows(index).Cells(0).Value = dsrmrefunds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(index).Cells(1).Value = dsrmrefunds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(index).Cells(2).Value = dsrmrefunds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(index).Cells(3).Value = dsrmrefunds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(index).Cells(4).Value = dsrmrefunds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(index).Cells(5).Value = dsrmrefunds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(index).Cells(6).Value = "RM Orange"
                DataGridView1.Rows(index).Cells(7).Value = "N/A"
                DataGridView1.Rows(index).Cells(8).Value = getAllTotalBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(9).Value = -getAllTotalReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(10).Value = getAllTotalRechargesAsReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(11).Value = getBalancesFromRM(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(12).Value = dsrmrefunds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(index).Cells(13).Value = getSuggestedDeduction(Integer.Parse(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString), conRMOrange)
                DataGridView1.Rows(index).Cells(14).Value = getEmployerName(DataGridView1.Rows(index).Cells(1).Value)
            Else
                DataGridView1.Rows.RemoveAt(index)
            End If
            ProgressBar1.Value = count
        Next
        ProgressBar1.Value = 0

        'Be Mobile
        dsrmrefunds.Clear()
        dsrmrefunds = getRMRefunds(DateTimePicker2.Value, conRMBeMobile)

        ProgressBar1.Maximum = dsrmrefunds.Tables(0).Rows.Count
        For count As Integer = 0 To dsrmrefunds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            index = DataGridView1.Rows.Count - 1
            If checkStoppedRefund(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString, Decimal.Parse(dsrmrefunds.Tables(0).Rows(count).Item(2).ToString), "RM BeMobile", dsrmrefunds.Tables(0).Rows(count).Item(1).ToString) = False Then
                DataGridView1.Rows(index).Cells(0).Value = dsrmrefunds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(index).Cells(1).Value = dsrmrefunds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(index).Cells(2).Value = dsrmrefunds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(index).Cells(3).Value = dsrmrefunds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(index).Cells(4).Value = dsrmrefunds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(index).Cells(5).Value = dsrmrefunds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(index).Cells(6).Value = "RM BeMobile"
                DataGridView1.Rows(index).Cells(7).Value = "N/A"
                DataGridView1.Rows(index).Cells(8).Value = getAllTotalBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(9).Value = -getAllTotalReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(10).Value = getAllTotalRechargesAsReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(11).Value = getBalancesFromRM(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(12).Value = dsrmrefunds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(index).Cells(13).Value = getSuggestedDeduction(Integer.Parse(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString), conRMBeMobile)
                DataGridView1.Rows(index).Cells(14).Value = getEmployerName(DataGridView1.Rows(index).Cells(1).Value)
            Else
                DataGridView1.Rows.RemoveAt(index)
            End If
            ProgressBar1.Value = count
        Next
        ProgressBar1.Value = 0

        ''MLM
        'dsrmrefunds.Clear()
        'dsrmrefunds = getMLMRefunds(DateTimePicker1.Value, conMLM)
        'ProgressBar1.Maximum = dsrmrefunds.Tables(0).Rows.Count
        'For count As Integer = 0 To dsrmrefunds.Tables(0).Rows.Count - 1
        '    DataGridView1.Rows.Add()
        '    index = DataGridView1.Rows.Count - 1
        '    If checkStoppedRefund(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString, Decimal.Parse(dsrmrefunds.Tables(0).Rows(count).Item(2).ToString), Date.Parse(dsrmrefunds.Tables(0).Rows(count).Item(5).ToString), dsrmrefunds.Tables(0).Rows(count).Item(6).ToString, dsrmrefunds.Tables(0).Rows(count).Item(1).ToString) = False Then
        '        DataGridView1.Rows(index).Cells(0).Value = dsrmrefunds.Tables(0).Rows(count).Item(0).ToString
        '        DataGridView1.Rows(index).Cells(1).Value = dsrmrefunds.Tables(0).Rows(count).Item(1).ToString
        '        DataGridView1.Rows(index).Cells(2).Value = dsrmrefunds.Tables(0).Rows(count).Item(2).ToString
        '        DataGridView1.Rows(index).Cells(3).Value = dsrmrefunds.Tables(0).Rows(count).Item(3).ToString
        '        DataGridView1.Rows(index).Cells(4).Value = dsrmrefunds.Tables(0).Rows(count).Item(4).ToString
        '        DataGridView1.Rows(index).Cells(5).Value = dsrmrefunds.Tables(0).Rows(count).Item(5).ToString
        '        DataGridView1.Rows(index).Cells(6).Value = "MLM"
        '        DataGridView1.Rows(index).Cells(7).Value = getAllCustomerBalance(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, conMLM)
        '        DataGridView1.Rows(index).Cells(8).Value = "N/A"
        '        DataGridView1.Rows(index).Cells(9).Value = "N/A"
        '        DataGridView1.Rows(index).Cells(10).Value = "N/A"
        '        DataGridView1.Rows(index).Cells(11).Value = "N/A"
        '        DataGridView1.Rows(index).Cells(12).Value = dsrmrefunds.Tables(0).Rows(count).Item(6).ToString
        '    Else
        '        DataGridView1.Rows.RemoveAt(index)
        '    End If
        '    ProgressBar1.Value = count
        'Next
        'ProgressBar1.Value = 0

        Label2.Text = DataGridView1.Rows.Count
        MessageBox.Show("Done loading", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        'If decision("YES = Display Ledger, NO = Stop/ Amend Deduction") = DialogResult.Yes Then
        '    If DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value = "RM Orange" Then
        '        Dim vrl As New RMCustomer_Ledger
        '        vrl.MdiParent = mform
        '        vrl.Show()
        '        vrl.TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
        '        vrl.ledgerOnExcel(conRMOrange)
        '    ElseIf DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value = "RM BeMobile" Then
        '        Dim vrl As New RMCustomer_Ledger
        '        vrl.MdiParent = mform
        '        vrl.Show()
        '        vrl.TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
        '        vrl.ledgerOnExcel(conRMBeMobile)
        '    ElseIf DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value = "MLM" Then
        '        Dim vrl As New MLMCustomerLedger
        '        vrl.MdiParent = mform
        '        vrl.TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
        '        vrl.ledger()
        '        vrl.Show()
        '    End If
        'Else
        '    Dim vrl As New UpdateLodgingReport
        '    vrl.MdiParent = mform
        '    vrl.TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
        '    vrl.Label5.Text = "Refund Amount : " & getCurrency(Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value))
        '    vrl.refundamount = Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value)
        '    vrl.refundid = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
        '    vrl.depositdate = Date.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value)
        '    vrl.fromrefundlist = True
        '    vrl.prts = Me
        '    vrl.index = DataGridView1.CurrentRow.Index
        '    vrl.Label6.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value
        '    vrl.Label8.Text = "Suggested Amount : " & Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(13).Value)
        '    vrl.Show()
        'End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        go = True
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(14).Value = "" Then
                go = False
            Else

            End If
        Next

        If go = True Then
            If decision("After exporting, the details will be removed from the system: Are you sure to continue") = DialogResult.Yes Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    saveStoppedRefunds(Integer.Parse(row.Cells(0).Value), Decimal.Parse(row.Cells(2).Value), Date.Now.Date, "Stopped", row.Cells(6).Value, row.Cells(1).Value)
                Next
            End If
            ExportToExcel(DataGridView1, "")
            MessageBox.Show("Done exporting", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        Else
            MessageBox.Show("Capture customers in the list that do not have Employer Details, that shows they are not existing ", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub loadEmployers(ByVal con As SqlClient.SqlConnection)
        Dim qry As String = "select * from Employer"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Emp")
            If ds.Tables(0).Rows.Count > 0 Then
                Me.cbemployer.DataSource = ds
                Me.cbemployer.DisplayMember = "Emp.Employer_Name"
                Me.cbemployer.ValueMember = "Emp.Code"
                Me.cbemployer.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub PotentialRefundsToStop_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadEmployers(conRMOrange)
        tbemployername.Text = ""
        tbemployerid.Text = ""

        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.ReadOnly = True
        Next

    End Sub

    Private Sub cbemployer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbemployer.SelectedIndexChanged
        tbemployername.Text = cbemployer.Text
        tbemployerid.Text = cbemployer.SelectedValue.ToString
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
        'Button1.Enabled = False
        ''Orange
        dsrmrefunds = getRMRefundsByEmployerID(DateTimePicker2.Value, tbemployerid.Text.Trim, conRMOrange)

        ProgressBar1.Maximum = dsrmrefunds.Tables(0).Rows.Count
        For count As Integer = 0 To dsrmrefunds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            index = DataGridView1.Rows.Count - 1
            If checkStoppedRefund(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString, Decimal.Parse(dsrmrefunds.Tables(0).Rows(count).Item(2).ToString), "RM Orange", dsrmrefunds.Tables(0).Rows(count).Item(1).ToString) = False Then
                DataGridView1.Rows(index).Cells(0).Value = dsrmrefunds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(index).Cells(1).Value = dsrmrefunds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(index).Cells(2).Value = dsrmrefunds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(index).Cells(3).Value = dsrmrefunds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(index).Cells(4).Value = dsrmrefunds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(index).Cells(5).Value = dsrmrefunds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(index).Cells(6).Value = "RM Orange"
                DataGridView1.Rows(index).Cells(7).Value = "N/A"
                DataGridView1.Rows(index).Cells(8).Value = getAllTotalBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(9).Value = -getAllTotalReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(10).Value = getAllTotalRechargesAsReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(11).Value = getBalancesFromRM(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMOrange)
                DataGridView1.Rows(index).Cells(12).Value = dsrmrefunds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(index).Cells(13).Value = getSuggestedDeduction(Integer.Parse(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString), conRMOrange)
                DataGridView1.Rows(index).Cells(14).Value = getEmployerName(DataGridView1.Rows(index).Cells(1).Value)
            Else
                DataGridView1.Rows.RemoveAt(index)
            End If
            ProgressBar1.Value = count
        Next
        ProgressBar1.Value = 0

        'Be Mobile
        dsrmrefunds.Clear()
        dsrmrefunds = getRMRefundsByEmployerID(DateTimePicker2.Value, tbemployerid.Text.Trim, conRMBeMobile)

        ProgressBar1.Maximum = dsrmrefunds.Tables(0).Rows.Count
        For count As Integer = 0 To dsrmrefunds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            index = DataGridView1.Rows.Count - 1
            If checkStoppedRefund(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString, Decimal.Parse(dsrmrefunds.Tables(0).Rows(count).Item(2).ToString), "RM BeMobile", dsrmrefunds.Tables(0).Rows(count).Item(1).ToString) = False Then
                DataGridView1.Rows(index).Cells(0).Value = dsrmrefunds.Tables(0).Rows(count).Item(0).ToString
                DataGridView1.Rows(index).Cells(1).Value = dsrmrefunds.Tables(0).Rows(count).Item(1).ToString
                DataGridView1.Rows(index).Cells(2).Value = dsrmrefunds.Tables(0).Rows(count).Item(2).ToString
                DataGridView1.Rows(index).Cells(3).Value = dsrmrefunds.Tables(0).Rows(count).Item(3).ToString
                DataGridView1.Rows(index).Cells(4).Value = dsrmrefunds.Tables(0).Rows(count).Item(4).ToString
                DataGridView1.Rows(index).Cells(5).Value = dsrmrefunds.Tables(0).Rows(count).Item(5).ToString
                DataGridView1.Rows(index).Cells(6).Value = "RM BeMobile"
                DataGridView1.Rows(index).Cells(7).Value = "N/A"
                DataGridView1.Rows(index).Cells(8).Value = getAllTotalBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(9).Value = -getAllTotalReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(10).Value = getAllTotalRechargesAsReceipts(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(11).Value = getBalancesFromRM(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile) + getAllPaidBilling(dsrmrefunds.Tables(0).Rows(count).Item(1).ToString, getLastDateOfMonth(Date.Now), conRMBeMobile)
                DataGridView1.Rows(index).Cells(12).Value = dsrmrefunds.Tables(0).Rows(count).Item(6).ToString
                DataGridView1.Rows(index).Cells(13).Value = getSuggestedDeduction(Integer.Parse(dsrmrefunds.Tables(0).Rows(count).Item(0).ToString), conRMBeMobile)
                DataGridView1.Rows(index).Cells(14).Value = getEmployerName(DataGridView1.Rows(index).Cells(1).Value)
            Else
                DataGridView1.Rows.RemoveAt(index)
            End If
            ProgressBar1.Value = count
        Next
        ProgressBar1.Value = 0

        Label2.Text = DataGridView1.Rows.Count
        MessageBox.Show("Done loading", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        ProgressBar1.Maximum = DataGridView1.Rows.Count
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(14).Value = "" Then
                row.Cells(14).Value = getEmployerName(row.Cells(1).Value)
            End If
            ProgressBar1.Value = row.Index
        Next
        ProgressBar1.Value = 0
        MessageBox.Show("Done refreshing", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class