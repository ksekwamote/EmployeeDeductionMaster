Public Class PaymentProcessing

    Dim ds, dsreference As New DataSet
    Dim index As Integer
    Dim totalamount As Decimal
    Dim add As Boolean
    Dim customerid As String
    Dim bank, branchcode, account, payeename As String

    Private Sub ChequesProcessing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        shoeHeadings()
    End Sub

    Public Sub shoeHeadings()
        ds = getProducts()
        For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = ds.Tables(0).Rows(index).Item(1).ToString()
            col.Name = ds.Tables(0).Rows(index).Item(0).ToString()
            DataGridView1.Columns.Add(col)
            col.ReadOnly = True
        Next

        Dim col1 As New DataGridViewTextBoxColumn()
        col1.HeaderText = "TotalAmount"
        col1.Name = "TotalAmount"
        col1.DefaultCellStyle.BackColor = Color.LightGreen
        col1.ReadOnly = True
        DataGridView1.Columns.Add(col1)

        Dim col2 As New DataGridViewTextBoxColumn()
        col2.HeaderText = "ReferenceNumber"
        col2.Name = "ReferenceNumber"
        col2.ReadOnly = True
        DataGridView1.Columns.Add(col2)

        Dim col3 As New DataGridViewTextBoxColumn()
        col3.HeaderText = "CustomerID"
        col3.Name = "CustomerID"
        col3.ReadOnly = True
        DataGridView1.Columns.Add(col3)

        Dim col4 As New DataGridViewTextBoxColumn()
        col4.HeaderText = "CustomerName"
        col4.Name = "CustomerName"
        col4.ReadOnly = True
        DataGridView1.Columns.Add(col4)

        Dim col5 As New DataGridViewTextBoxColumn()
        col5.HeaderText = "BankName"
        col5.Name = "BankName"
        col5.ReadOnly = True
        DataGridView1.Columns.Add(col5)

        Dim col6 As New DataGridViewTextBoxColumn()
        col6.HeaderText = "BranchCode"
        col6.Name = "BranchCode"
        col6.ReadOnly = True
        DataGridView1.Columns.Add(col6)

        Dim col7 As New DataGridViewTextBoxColumn()
        col7.HeaderText = "Account1"
        col7.Name = "Account1"
        col7.ReadOnly = True
        DataGridView1.Columns.Add(col7)

        Dim col8 As New DataGridViewTextBoxColumn()
        col8.HeaderText = "Account2"
        col8.Name = "Account2"
        col8.ReadOnly = True
        DataGridView1.Columns.Add(col8)

        ds.Clear()
    End Sub

    Function getProducts() As DataSet
        Dim qry As String = "select * from Products where UseInCheques = 'Yes' order by AllocationOrder ASC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Sub getChequesFromSystems(ByVal customerid As String)
        DataGridView1.Rows.Add()
        index = DataGridView1.Rows.Count - 1
        totalamount = 0
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If (col.Index = (DataGridView1.Columns.Count - 1)) Or (col.Index = (DataGridView1.Columns.Count - 2)) Or (col.Index = (DataGridView1.Columns.Count - 3)) Or (col.Index = (DataGridView1.Columns.Count - 4)) Or (col.Index = (DataGridView1.Columns.Count - 5)) Or (col.Index = (DataGridView1.Columns.Count - 6)) Or (col.Index = (DataGridView1.Columns.Count - 7)) Or (col.Index = (DataGridView1.Columns.Count - 8)) Then
                If (col.Index = (DataGridView1.Columns.Count - 8)) Then
                    DataGridView1.Rows(index).Cells(col.Index).Value = totalamount
                ElseIf (col.Index = (DataGridView1.Columns.Count - 6)) Then
                    DataGridView1.Rows(index).Cells(col.Index).Value = customerid
                ElseIf (col.Index = (DataGridView1.Columns.Count - 5)) Then
                    DataGridView1.Rows(index).Cells(col.Index).Value = getCustomerName(customerid)
                ElseIf (col.Index = (DataGridView1.Columns.Count - 4)) Then

                ElseIf (col.Index = (DataGridView1.Columns.Count - 3)) Then

                ElseIf (col.Index = (DataGridView1.Columns.Count - 2)) Then

                ElseIf (col.Index = (DataGridView1.Columns.Count - 1)) Then

                End If
            Else
                'MessageBox.Show(Integer.Parse(col.Name) & " = " & getProductSystem(Integer.Parse(col.Name)), "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If getProductSystem(Integer.Parse(col.Name)) = "FPM Funeral" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getFPMFuneralRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getFPMFuneralRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getFPMFuneralRefundPayByReference(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "FPM Membership" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getFPMMembershipRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getFPMMembershipRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getFPMMembeshipRefundPayByReference(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "FPM Savings" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getFPMSavingsRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getFPMSavingsRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getFPMSavingsRefundPayByReference(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conFPM)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Quick" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getMLMRLoan(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getMLMRLoan(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getMLMRLoanPayByReference(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Quick Refund" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getMLMRefund(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getMLMRefund(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getMLMRefundPayByReference(customerid, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Loan" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getMLMRLoan(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getMLMRLoan(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getMLMRLoanPayByReference(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Educational" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getMLMRLoan(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMEducational), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getMLMRLoan(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMEducational)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getMLMRLoanPayByReference(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMEducational)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Loan Refund" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getMLMRefund(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getMLMRefund(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getMLMRefundPayByReference(customerid, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "RM Orange" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getRMRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMOrange), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getRMRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMOrange)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getRMRefundPayByReference(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMOrange)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                ElseIf getProductSystem(Integer.Parse(col.Name)) = "RM BeMobile" Then
                    If checkSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), getRMRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMBeMobile), tbpayby.Text.Trim) = True Then
                        DataGridView1.Rows(index).Cells(col.Index).Value = 0.0
                        totalamount = totalamount + 0.0
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    Else
                        DataGridView1.Rows(index).Cells(col.Index).Value = getRMRefund(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMBeMobile)
                        totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                        dsreference.Clear()
                        dsreference = getRMRefundPayByReference(customerid, DateTimePicker1.Value, tbpayby.Text.Trim, conRMBeMobile)
                        If dsreference.Tables(0).Rows.Count > 0 Then
                            bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                            branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                            account = dsreference.Tables(0).Rows(0).Item(2).ToString
                            payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                        Else
                            bank = "N/A"
                            branchcode = "N/A"
                            account = "N/A"
                            payeename = "N/A"
                        End If
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                Else
                    DataGridView1.Rows(index).Cells(col.Index).Value = getSystemCheque(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), tbpayby.Text.Trim)
                    totalamount = totalamount + Decimal.Parse(DataGridView1.Rows(index).Cells(col.Index).Value)
                    dsreference.Clear()
                    dsreference = getSystemChequePayByReference(customerid, DateTimePicker1.Value, Integer.Parse(col.Name), tbpayby.Text.Trim)
                    If dsreference.Tables(0).Rows.Count > 0 Then
                        bank = dsreference.Tables(0).Rows(0).Item(0).ToString
                        branchcode = dsreference.Tables(0).Rows(0).Item(1).ToString
                        account = dsreference.Tables(0).Rows(0).Item(2).ToString
                        payeename = dsreference.Tables(0).Rows(0).Item(3).ToString
                    Else
                        bank = "N/A"
                        branchcode = "N/A"
                        account = "N/A"
                        payeename = "N/A"
                    End If
                    If DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "" Or DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = "N/A" Then
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 4).Value = bank
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 3).Value = branchcode
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 2).Value = account
                        DataGridView1.Rows(index).Cells(DataGridView1.Columns.Count - 1).Value = payeename
                    End If
                End If
            End If
        Next

        If Decimal.Parse(DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(DataGridView1.Columns.Count - 8).Value) > 0 Then

        Else
            DataGridView1.Rows.RemoveAt(DataGridView1.Rows.Count - 1)
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Please select the pay by", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbcustomerid.Text.Trim = "" Then
                MessageBox.Show("Enter Customer id first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                loadAllTheRequest(tbcustomerid.Text.Trim)
            End If
        End If
    End Sub

    Public Sub loadAllTheRequest(ByVal idnumber As String)
        DataGridView1.Rows.Clear()

        ''FPM Funeral Refund
        ds = getFPMFuneralRefundCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, conFPM, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''FPM Membership Refund
        ds = getFPMMembershipRefundCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, conFPM, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''FPM Savings Refund
        ds = getFPMSavingsRefundCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, conFPM, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''MLM QUICK REFUND
        ds = getMLMRefundCustomer(DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''MLM NORMAL LOAN REFUND
        ds = getMLMRefundCustomer(DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''MLM QUICK LOAN
        ds = getMLMRLoanCustomer(DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(11).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''MLM NORMAL LOAN
        ds = getMLMRLoanCustomer(DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(11).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''RM Orange Refund
        ds = getRMRefundCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, conRMOrange, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''RM BeMobile Refund
        ds = getRMRefundCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, conRMBeMobile, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        ''System Refund
        ds = getSystemChequeCustomer(DateTimePicker1.Value, tbpayby.Text.Trim, idnumber)
        For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
            DataGridView4.Rows.Add()
            index = DataGridView4.Rows.Count - 1
            customerid = ds.Tables(0).Rows(count).Item(1).ToString

            add = True
            For count1 As Integer = 0 To ds.Tables(0).Rows.Count - 1
                If DataGridView4.Rows(index).Cells(0).Value = customerid Then
                    add = False
                Else

                End If
            Next

            If add = True Then
                getChequesFromSystems(customerid)
            End If
        Next

        GroupBox1.Enabled = True
        loadCheques(DateTimePicker1.Value, "Requested", tbpayby.Text.Trim, DataGridView2, BindingNavigator1)
        loadCheques(DateTimePicker1.Value, "Written", tbpayby.Text.Trim, DataGridView3, BindingNavigator2)

        DataGridView4.Rows.Clear()
        MessageBox.Show("Loaded", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 8).Value) > 0 Then
                If checkRequestedCheques(DateTimePicker1.Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value) = False Then
                    If DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 5).Value = "" Then
                        MessageBox.Show("Customer does not exist in the system. Please add first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If decision("Are you sure to request cheque for the customer ?") = DialogResult.Yes Then
                            For Each col As DataGridViewColumn In DataGridView1.Columns
                                If (col.Index = (DataGridView1.Columns.Count - 1)) Or (col.Index = (DataGridView1.Columns.Count - 2)) Or (col.Index = (DataGridView1.Columns.Count - 3)) Or (col.Index = (DataGridView1.Columns.Count - 4)) Or (col.Index = (DataGridView1.Columns.Count - 5)) Or (col.Index = (DataGridView1.Columns.Count - 6)) Or (col.Index = (DataGridView1.Columns.Count - 7)) Or (col.Index = (DataGridView1.Columns.Count - 8)) Then

                                Else
                                    If checkIFProductHasSystem(Integer.Parse(col.Name)) = True Then
                                        If Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(col.Index).Value) > 0 Then
                                            If tbpayby.Text.Trim = "Cheque" Then
                                                saveRefund(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, Integer.Parse(col.Name), Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(col.Index).Value), DateTimePicker1.Value, logedinname, Date.Now, "", "Requested", tbpayby.Text.Trim, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 4).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 3).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 2).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 1).Value)
                                            Else
                                                saveRefund(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, Integer.Parse(col.Name), Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(col.Index).Value), DateTimePicker1.Value, logedinname, Date.Now, "N/A", "Written", tbpayby.Text.Trim, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 4).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 3).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 2).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 1).Value)
                                            End If
                                            If getProductSystem(Integer.Parse(col.Name)) = "MLM Loan" Then
                                                ds.Clear()
                                                ds = getMLMRLoanDetails(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, DateTimePicker1.Value, "Normal", tbpayby.Text.Trim, conMLMAdvances)
                                                saveBalance(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString), Decimal.Parse(ds.Tables(0).Rows(0).Item(1).ToString), Decimal.Parse(ds.Tables(0).Rows(0).Item(2).ToString), DateTimePicker1.Value, "Normal", "New")
                                            ElseIf getProductSystem(Integer.Parse(col.Name)) = "MLM Quick" Then
                                                ds.Clear()
                                                ds = getMLMRLoanDetails(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, DateTimePicker1.Value, "Quick", tbpayby.Text.Trim, conMLMAdvances)
                                                saveBalance(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString), Decimal.Parse(ds.Tables(0).Rows(0).Item(1).ToString), Decimal.Parse(ds.Tables(0).Rows(0).Item(2).ToString), DateTimePicker1.Value, "Quick", "New")
                                            End If
                                        End If
                                    Else
                                        If Decimal.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(col.Index).Value) > 0 Then
                                            If tbpayby.Text.Trim = "Cheque" Then
                                                updateSystemRefundStatusToRequested(Integer.Parse(col.Name), DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, DateTimePicker1.Value, "Requested")
                                            Else
                                                updateSystemRefundChequeNumber(Integer.Parse(col.Name), DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, DateTimePicker1.Value, "N/A")
                                                updateSystemRefundStatusToRequested(Integer.Parse(col.Name), DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(DataGridView1.Columns.Count - 6).Value, DateTimePicker1.Value, "Written")
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            wrapper.saveAction("Cheque Approved And Requested", Date.Now, logedinname, DataGridView1.Rows(0).Cells(DataGridView1.Columns.Count - 5).Value)
                            MessageBox.Show("Cheque Requested", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            loadAllTheRequest(tbcustomerid.Text.Trim)
                            loadCheques(DateTimePicker1.Value, "Requested", tbpayby.Text.Trim, DataGridView2, BindingNavigator1)
                            loadCheques(DateTimePicker1.Value, "Written", tbpayby.Text.Trim, DataGridView3, BindingNavigator2)
                        End If
                    End If
                Else
                    MessageBox.Show("Write the requested cheque first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Private Sub loadCheques(ByVal paymentdate As Date, ByVal Status As String, ByVal payby As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select CustomerID,SUM(Amount) as TotalAmount,PaymentDate from ProductsCheques where PaymentDate = '" & paymentdate & "' and Status = '" & Status & "' and PayBy = '" & payby & "' group by CustomerID,PaymentDate"
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
                'filterManager = New DgvFilterManager(DataGridView1)
            Else
                dgv.DataSource = Nothing
                bn.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbpayby.Text = ComboBox1.Text
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

End Class