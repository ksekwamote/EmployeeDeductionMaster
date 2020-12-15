Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class FileConfirmation

    Dim count, discount As Integer
    Dim typestr, typestr1 As String
    Dim ds As New DataSet
    Dim systemname As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcustomerid.Text.Trim = "" Then
            MessageBox.Show("Enter Customer ID", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            DataGridView1.Rows.Clear()
            loadTotal()
            DataGridView2.DataSource = Nothing
            BindingNavigator2.BindingSource = Nothing
            DataGridView3.DataSource = Nothing
            BindingNavigator3.BindingSource = Nothing
        End If
    End Sub

    Public Sub loadTotal()
        'Advances
        typestr = "Refund"
        count = getCountForFilling(conMLMAdvances, "select Count(AutoID) from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conMLMAdvances, "select Count(AutoID) from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Advances"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Loan"
        count = getCountForFilling(conMLMAdvances, "select Count(AutoLoanID) from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conMLMAdvances, "select Count(AutoLoanID) from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Loan"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Advances"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        'Educational
        typestr = "Refund"
        count = getCountForFilling(conMLMEducational, "select Count(AutoID) from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conMLMEducational, "select Count(AutoID) from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Educational"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Loan"
        count = getCountForFilling(conMLMEducational, "select Count(AutoLoanID) from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conMLMEducational, "select Count(AutoLoanID) from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Loan"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Educational"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        'Orange
        typestr = "Refund"
        count = getCountForFilling(conRMOrange, "select Count(AutoID) from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMOrange, "select Count(AutoID) from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Deal"
        count = getCountForFilling(conRMOrange, "select Count(DealNo) from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMOrange, "select Count(DealNo) from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Deal"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Voucher"
        count = getCountForFilling(conRMOrange, "select Count(EntryNo) from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMOrange, "select Count(EntryNo) from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Voucher"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        'BeMobile
        typestr = "Refund"
        count = getCountForFilling(conRMBeMobile, "select Count(AutoID) from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMBeMobile, "select Count(AutoID) from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Deal"
        count = getCountForFilling(conRMBeMobile, "select Count(DealNo) from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMBeMobile, "select Count(DealNo) from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Deal"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Voucher"
        count = getCountForFilling(conRMBeMobile, "select Count(EntryNo) from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conRMBeMobile, "select Count(EntryNo) from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Voucher"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        'FPM
        typestr = "Refund"
        count = getCountForFilling(conFPM, "select Count(AutoID) from CustomerRefunds where Status = 'Approved' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conFPM, "select Count(AutoID) from CustomerRefunds where Status = 'Approved' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "BenefitSchedule"
        count = getCountForFilling(conFPM, "select Count(AutoID) from BenefitScheduleMessages where Status = 'Complete' and CompletedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and PayerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conFPM, "select Count(AutoID) from BenefitScheduleMessages where Status = 'Complete' and CompletedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and PayerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "BenefitSchedule"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        typestr = "Policy"
        count = getCountForFilling(conFPM, "select Count(Policy.PolicyNumber) from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where PolicyStatus = 'Active' and ApprovalDate >= '" & DateTimePicker1.Value & "' and Policy.PolicyNumber not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and InitiatedByIDNumber = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(conFPM, "select Count(Policy.PolicyNumber) from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where PolicyStatus = 'Active' and ApprovalDate >= '" & DateTimePicker1.Value & "' and Policy.PolicyNumber in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and InitiatedByIDNumber = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Policy"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If

        'EDM
        typestr = "EDM"
        count = getCountForFilling(con, "select Count(AutoID) from SettlementBalances where Status = 'New' and CapturedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        discount = getCountForFilling(con, "select Count(AutoID) from SettlementBalances where Status = 'New' and CapturedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'")
        If count > 0 Or discount > 0 Then
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = "Settlement"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "EDM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = count
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = discount
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            systemname = DataGridView1.CurrentRow.Cells(1).Value
            typestr = DataGridView1.CurrentRow.Cells(0).Value
            typestr1 = typestr
            loadSpecificDataToConfirm()
            loadSpecificDataConfirmed()
        End If
    End Sub

    Public Sub loadSpecificDataToConfirm()
        If systemname = "Advances" Then
            If typestr = "Refund" Then
                loadData(conMLMAdvances, "select * from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Loan" Then
                loadData(conMLMAdvances, "select * from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "Educational" Then
            If typestr = "Refund" Then
                loadData(conMLMEducational, "select * from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Loan" Then
                loadData(conMLMEducational, "select * from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "Orange" Then
            If typestr = "Refund" Then
                loadData(conRMOrange, "select * from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Deal" Then
                loadData(conRMOrange, "select * from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Voucher" Then
                loadData(conRMOrange, "select * from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "BeMobile" Then
            If typestr = "Refund" Then
                loadData(conRMBeMobile, "select * from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Deal" Then
                loadData(conRMBeMobile, "select * from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Voucher" Then
                loadData(conRMBeMobile, "select * from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "FPM" Then
            If typestr = "Refund" Then
                loadData(conFPM, "select * from CustomerRefunds where Status = 'Approved' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "BenefitSchedule" Then
                loadData(conFPM, "select * from BenefitScheduleMessages where Status = 'Complete' and CompletedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and PayerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Policy" Then
                loadData(conFPM, "select Policy.PolicyNumber,PayerID,AmountDeducted,InitiatedByIDNumber as PrincipalID,BeneficiaryID,DeductionDate,CreationDate,CreatedBy,ApprovedBy,ApprovalDate from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where PolicyStatus = 'Active' and ApprovalDate >= '" & DateTimePicker1.Value & "' and Policy.PolicyNumber not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and InitiatedByIDNumber = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "EDM" Then
            If typestr = "Settlement" Then
                loadData(con, "select * from SettlementBalances where Status = 'New' and CapturedDate >= '" & DateTimePicker1.Value & "' and AutoID not in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView2, BindingNavigator2)
            End If
        End If
    End Sub

    Private Sub loadData(ByVal con As SqlClient.SqlConnection, ByVal str As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = str
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
                bn.BindingSource = bs
                dgv.DataSource = bs
                dgv.Columns(dgv.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(dgv)
            Else
                dgv.DataSource = Nothing
                bn.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadData()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        If DataGridView2.Rows.Count > 0 Then
            If decision("Is the document filed ?") = DialogResult.Yes Then
                If systemname = "Advances" Then
                    If typestr = "Refund" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conMLMAdvances)
                    ElseIf typestr = "Loan" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conMLMAdvances)
                    End If
                ElseIf systemname = "Educational" Then
                    If typestr = "Refund" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conMLMEducational)
                    ElseIf typestr = "Loan" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conMLMEducational)
                    End If
                ElseIf systemname = "Orange" Then
                    If typestr = "Refund" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMOrange)
                    ElseIf typestr = "Deal" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMOrange)
                    ElseIf typestr = "Voucher" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMOrange)
                    End If
                ElseIf systemname = "BeMobile" Then
                    If typestr = "Refund" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMBeMobile)
                    ElseIf typestr = "Deal" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMBeMobile)
                    ElseIf typestr = "Voucher" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conRMBeMobile)
                    End If
                ElseIf systemname = "FPM" Then
                    If typestr = "Refund" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conFPM)
                    ElseIf typestr = "BenefitSchedule" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conFPM)
                    ElseIf typestr = "Policy" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", conFPM)
                    End If
                ElseIf systemname = "EDM" Then
                    If typestr = "Settlement" Then
                        saveFillingConfirmation(typestr, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value), "Active", con)
                    End If
                End If
                MessageBox.Show("Done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.Rows.Clear()
                loadTotal()
                typestr = typestr1
                loadSpecificDataToConfirm()
                loadSpecificDataConfirmed()
            End If
        End If
    End Sub

    Public Sub loadSpecificDataConfirmed()
        If systemname = "Advances" Then
            If typestr = "Refund" Then
                loadData(conMLMAdvances, "select * from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Loan" Then
                loadData(conMLMAdvances, "select * from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        ElseIf systemname = "Educational" Then
            If typestr = "Refund" Then
                loadData(conMLMEducational, "select * from Refunds where Status = 'Posted' and ApprovalDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Loan" Then
                loadData(conMLMEducational, "select * from LoanDetails where LoanStatus = 'Active' and CreationDate >= '" & DateTimePicker1.Value & "' and AutoLoanID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        ElseIf systemname = "Orange" Then
            If typestr = "Refund" Then
                loadData(conRMOrange, "select * from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Deal" Then
                loadData(conRMOrange, "select * from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Voucher" Then
                loadData(conRMOrange, "select * from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        ElseIf systemname = "BeMobile" Then
            If typestr = "Refund" Then
                loadData(conRMBeMobile, "select * from CustomerRefunds where Status = 'Posted' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Deal" Then
                loadData(conRMBeMobile, "select * from CustomerDeals where DealStatus not in ('Cancelled','Default') and DealActivationDate >= '" & DateTimePicker1.Value & "' and DealNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and Account = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Voucher" Then
                loadData(conRMBeMobile, "select * from MSPReceivedVouchers where Status = 'Collected' and CollectedDate >= '" & DateTimePicker1.Value & "' and EntryNo in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        ElseIf systemname = "FPM" Then
            If typestr = "Refund" Then
                loadData(conFPM, "select * from CustomerRefunds where Status = 'Approved' and ApprovedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "BenefitSchedule" Then
                loadData(conFPM, "select * from BenefitScheduleMessages where Status = 'Complete' and CompletedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and PayerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            ElseIf typestr = "Policy" Then
                loadData(conFPM, "select Policy.PolicyNumber,PayerID,AmountDeducted,InitiatedByIDNumber as PrincipalID,BeneficiaryID,DeductionDate,CreationDate,CreatedBy,ApprovedBy,ApprovalDate from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where PolicyStatus = 'Active' and ApprovalDate >= '" & DateTimePicker1.Value & "' and Policy.PolicyNumber in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and InitiatedByIDNumber = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        ElseIf systemname = "EDM" Then
            If typestr = "Settlement" Then
                loadData(con, "select * from SettlementBalances where Status = 'New' and CapturedDate >= '" & DateTimePicker1.Value & "' and AutoID in (select TypeID from FillingConfirmation where Type = '" & typestr & "' and Status = 'Active') and CustomerID = '" & tbcustomerid.Text.Trim & "'", DataGridView3, BindingNavigator3)
            End If
        End If
    End Sub

End Class