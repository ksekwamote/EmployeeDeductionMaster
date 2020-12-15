Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class CustomerFiles

    Dim count, discount As Integer
    Dim typestr, typestr1 As String
    Dim ds As New DataSet
    Dim systemname As String
    Dim sCustomerID, sType As String
    Dim sTypeID As Integer

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcustomerid.Text.Trim = "" Then
            MessageBox.Show("Enter Customer ID", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            DataGridView1.Rows.Clear()
            loadTotal()
            DataGridView2.DataSource = Nothing
            BindingNavigator2.BindingSource = Nothing
        End If
    End Sub

    Public Sub loadTotal()
        'Advances
        typestr = "Refund"
        ds.Clear()
        ds = getDataset("select AutoID from Refunds where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMAdvances)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMAdvances)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Advances"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Loan"
        ds.Clear()
        ds = getDataset("select AutoLoanID from LoanDetails where AutoLoanID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", conMLMAdvances)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMAdvances)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Loan"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Advances"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        'Educational
        typestr = "Refund"
        ds.Clear()
        ds = getDataset("select AutoID from Refunds where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMEducational)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMEducational)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Educational"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Loan"
        ds.Clear()
        ds = getDataset("select AutoLoanID from LoanDetails where AutoLoanID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and Customer_ID = '" & tbcustomerid.Text.Trim & "'", conMLMEducational)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conMLMEducational)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Loan"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Educational"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        'Orange
        typestr = "Refund"
        ds.Clear()
        ds = getDataset("select AutoID from CustomerRefunds where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Deal"
        ds.Clear()
        ds = getDataset("select DealNo from CustomerDeals where DealNo in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and Account = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Deal"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Voucher"
        ds.Clear()
        ds = getDataset("select EntryNo from MSPReceivedVouchers where EntryNo in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMOrange)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Voucher"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "Orange"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        'BeMobile
        typestr = "Refund"
        ds.Clear()
        ds = getDataset("select AutoID from CustomerRefunds where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Deal"
        ds.Clear()
        ds = getDataset("select DealNo from CustomerDeals where DealNo in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and Account = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Deal"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Voucher"
        ds.Clear()
        ds = getDataset("select EntryNo from MSPReceivedVouchers where EntryNo in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conRMBeMobile)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Voucher"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "BeMobile"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        'FPM
        typestr = "Refund"
        ds.Clear()
        ds = getDataset("select AutoID from CustomerRefunds where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", conFPM)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conFPM)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Refund"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "BenefitSchedule"
        ds.Clear()
        ds = getDataset("select AutoID from BenefitScheduleMessages where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and PayerID = '" & tbcustomerid.Text.Trim & "'", conFPM)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conFPM)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "BenefitSchedule"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        typestr = "Policy"
        ds.Clear()
        ds = getDataset("select Policy.PolicyNumber from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where Policy.PolicyNumber in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and InitiatedByIDNumber = '" & tbcustomerid.Text.Trim & "'", conFPM)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", conFPM)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Policy"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "FPM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next

        'EDM
        typestr = "EDM"
        ds.Clear()
        ds = getDataset("select AutoID from SettlementBalances where AutoID in (select TypeID from FillingDocuments where Type = '" & typestr & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "') and CustomerID = '" & tbcustomerid.Text.Trim & "'", con)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = getStringValue("select FillingNumber from FillingDocuments where Type = '" & typestr & "' and TypeID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString) & "' and Status = 'Active' and CustomerID = '" & tbcustomerid.Text.Trim & "'", con)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "Settlement"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "EDM"
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ds.Tables(0).Rows(count).Item(0).ToString
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = tbcustomerid.Text.Trim
        Next
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            systemname = DataGridView1.CurrentRow.Cells(2).Value
            typestr = DataGridView1.CurrentRow.Cells(1).Value
            typestr1 = typestr
            loadSpecificDataToConfirm(Integer.Parse(DataGridView1.CurrentRow.Cells(3).Value))

            sCustomerID = DataGridView1.CurrentRow.Cells(4).Value
            sType = DataGridView1.CurrentRow.Cells(1).Value
            sTypeID = Integer.Parse(DataGridView1.CurrentRow.Cells(3).Value)
        End If
    End Sub

    Public Sub loadSpecificDataToConfirm(ByVal AutoID As Integer)
        If systemname = "Advances" Then
            If typestr = "Refund" Then
                loadData(conMLMAdvances, "select * from Refunds where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Loan" Then
                loadData(conMLMAdvances, "select * from LoanDetails where AutoLoanID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "Educational" Then
            If typestr = "Refund" Then
                loadData(conMLMEducational, "select * from Refunds where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Loan" Then
                loadData(conMLMEducational, "select * from LoanDetails where AutoLoanID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "Orange" Then
            If typestr = "Refund" Then
                loadData(conRMOrange, "select * from CustomerRefunds where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Deal" Then
                loadData(conRMOrange, "select * from CustomerDeals where DealNo = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Voucher" Then
                loadData(conRMOrange, "select * from MSPReceivedVouchers where EntryNo = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "BeMobile" Then
            If typestr = "Refund" Then
                loadData(conRMBeMobile, "select * from CustomerRefunds where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Deal" Then
                loadData(conRMBeMobile, "select * from CustomerDeals where DealNo = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Voucher" Then
                loadData(conRMBeMobile, "select * from MSPReceivedVouchers where EntryNo = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "FPM" Then
            If typestr = "Refund" Then
                loadData(conFPM, "select * from CustomerRefunds where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "BenefitSchedule" Then
                loadData(conFPM, "select * from BenefitScheduleMessages where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            ElseIf typestr = "Policy" Then
                loadData(conFPM, "select Policy.PolicyNumber,PayerID,AmountDeducted,InitiatedByIDNumber as PrincipalID,BeneficiaryID,DeductionDate,CreationDate,CreatedBy,ApprovedBy,ApprovalDate from Policy inner join PolicyPaymentMethod on Policy.PolicyNumber = PolicyPaymentMethod.PolicyNumber where Policy.PolicyNumber = '" & AutoID & "'", DataGridView2, BindingNavigator2)
            End If
        ElseIf systemname = "EDM" Then
            If typestr = "Settlement" Then
                loadData(con, "select * from SettlementBalances where AutoID = '" & AutoID & "'", DataGridView2, BindingNavigator2)
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

End Class