Public Class SalesProducts
    Dim dscustomers, dsproducts, dscustpro As New DataSet
    Dim batchname As String

    Dim adddown As Boolean = False
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, Label3.Text & "between " & DateTimePicker1.Value & " and " & DateTimePicker1.Value)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If adddown = True Then
            If DataGridView1.Rows.Count > 0 Then
                If decision("Pay Commission for the above products ?") = DialogResult.Yes Then
                    Button2.Enabled = False
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        DataGridView2.Rows.Add()
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0).Value = row.Cells(0).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(1).Value = row.Cells(1).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(2).Value = row.Cells(2).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(3).Value = row.Cells(3).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(4).Value = row.Cells(4).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(5).Value = row.Cells(5).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(6).Value = row.Cells(6).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(7).Value = row.Cells(7).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(8).Value = row.Cells(8).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(9).Value = row.Cells(9).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(10).Value = row.Cells(10).Value
                        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(11).Value = row.Cells(11).Value
                    Next
                    DataGridView1.Rows.Clear()
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If decision("Are you sure to proceed ?") = DialogResult.Yes Then
            Button3.Enabled = False
            batchname = "Claim:" & Date.Now
            For Each row As DataGridViewRow In DataGridView2.Rows
                If cbproducts.Checked = True Then
                    saveSalesCommission(row.Cells(0).Value, row.Cells(1).Value, row.Cells(2).Value, Date.Parse(row.Cells(3).Value), Date.Parse(row.Cells(4).Value), row.Cells(5).Value, Date.Parse(row.Cells(6).Value), Decimal.Parse(row.Cells(7).Value), row.Cells(8).Value, row.Cells(9).Value, Integer.Parse(row.Cells(10).Value), "Claimed", batchname, Date.Now, logedinname, row.Cells(11).Value)
                ElseIf cbcontacted.Checked = True Then
                    saveSalesCommission(row.Cells(0).Value, row.Cells(1).Value, row.Cells(2).Value, Date.Parse(row.Cells(3).Value), Date.Now, row.Cells(5).Value, Date.Now, 0.00, row.Cells(8).Value, row.Cells(9).Value, 0, "Claimed", batchname, Date.Now, logedinname, row.Cells(11).Value)
                End If
            Next
            If cbproducts.Checked = True Then
                ExportToExcel(DataGridView2, Label3.Text & "between " & DateTimePicker1.Value & " and " & DateTimePicker1.Value)
            ElseIf cbcontacted.Checked = True Then
                ExportToExcel(DataGridView2, Label3.Text & "between " & DateTimePicker1.Value & " and " & DateTimePicker1.Value)
            End If
            MessageBox.Show("Done ", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        adddown = True
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

        If cbproducts.Checked = True Then
            Label3.Text = "New Products Commission"
            dsproducts.Clear()
            dsproducts = getSalesProductsSystems()

            dscustomers.Clear()
            dscustomers = getSalesCustomers(DateTimePicker2.Value)
            For countcust As Integer = 0 To dscustomers.Tables(0).Rows.Count - 1
                For countpro As Integer = 0 To dsproducts.Tables(0).Rows.Count - 1
                    If dsproducts.Tables(0).Rows(countpro).Item(1).ToString = "MLM Loan" Then
                        dscustpro.Clear()
                        dscustpro = getDataset("select CreationDate,Principal_Amount,AutoLoanID from LoanDetails where CreationDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and CreationDate >= '" & Date.Parse(dscustomers.Tables(0).Rows(countcust).Item(2).ToString) & "' and LoanStatus = 'Active' and LoanDescription = 'Normal' and Customer_ID = '" & dscustomers.Tables(0).Rows(countcust).Item(0).ToString & "'", conMLMAdvances)
                        For countitem As Integer = 0 To dscustpro.Tables(0).Rows.Count - 1
                            If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Normal Loan"
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustpro.Tables(0).Rows(countitem).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "MLM Loan"
                            End If
                        Next
                    End If
                    If dsproducts.Tables(0).Rows(countpro).Item(1).ToString = "MLM Quick" Then
                        dscustpro.Clear()
                        dscustpro = getDataset("select CreationDate,Principal_Amount,AutoLoanID from LoanDetails where CreationDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and CreationDate >= '" & Date.Parse(dscustomers.Tables(0).Rows(countcust).Item(2).ToString) & "' and LoanStatus = 'Active' and LoanDescription = 'Quick' and Customer_ID = '" & dscustomers.Tables(0).Rows(countcust).Item(0).ToString & "'", conMLMAdvances)
                        For countitem As Integer = 0 To dscustpro.Tables(0).Rows.Count - 1
                            If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Quick Loan"
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustpro.Tables(0).Rows(countitem).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "MLM Quick"
                            End If
                        Next
                    End If
                    If dsproducts.Tables(0).Rows(countpro).Item(1).ToString = "RM Orange" Then
                        dscustpro.Clear()
                        dscustpro = getDataset("select CreationDate,DealAmount,DealNo from CustomerDeals inner join DealTemplate on CustomerDeals.DTCode = DealTemplate.DTCode where CreationDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and CreationDate >= '" & Date.Parse(dscustomers.Tables(0).Rows(countcust).Item(2).ToString) & "' and DealStatus in ('Stopped','Active','Received') and Account = '" & dscustomers.Tables(0).Rows(countcust).Item(0).ToString & "'", conRMOrange)
                        For countitem As Integer = 0 To dscustpro.Tables(0).Rows.Count - 1
                            If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Orange Deal"
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustpro.Tables(0).Rows(countitem).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "RM Orange"
                            End If
                        Next
                    End If
                    If dsproducts.Tables(0).Rows(countpro).Item(1).ToString = "RM BeMobile" Then
                        dscustpro.Clear()
                        dscustpro = getDataset("select CreationDate,DealAmount,DealNo from CustomerDeals inner join DealTemplate on CustomerDeals.DTCode = DealTemplate.DTCode where CreationDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and CreationDate >= '" & Date.Parse(dscustomers.Tables(0).Rows(countcust).Item(2).ToString) & "' and DealStatus in ('Stopped','Active','Received') and Account = '" & dscustomers.Tables(0).Rows(countcust).Item(0).ToString & "'", conRMBeMobile)
                        For countitem As Integer = 0 To dscustpro.Tables(0).Rows.Count - 1
                            If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Bemobile Deal"
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustpro.Tables(0).Rows(countitem).Item(1).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "RM BeMobile"
                            End If
                        Next
                    End If
                    If dsproducts.Tables(0).Rows(countpro).Item(1).ToString = "FPM Funeral" Then
                        dscustpro.Clear()
                        dscustpro = getDataset("select CreationDate,InitiatedByIDNumber,PolicyNumber from Policy where CreationDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and CreationDate >= '" & Date.Parse(dscustomers.Tables(0).Rows(countcust).Item(2).ToString) & "' and PolicyStatus in ('Active') and InitiatedByIDNumber = '" & dscustomers.Tables(0).Rows(countcust).Item(0).ToString & "'", conFPM)
                        For countitem As Integer = 0 To dscustpro.Tables(0).Rows.Count - 1
                            If getMembershipFeeByGroupID(GetGroupIDByCustomerID(dscustomers.Tables(0).Rows(countcust).Item(0).ToString, conFPM), conFPM) > 0 Then
                                If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                    DataGridView1.Rows.Add()
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Membership"
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = getTotalPremiumPerPolicyNumberPrincipalIDandCoverStartDate(DateTimePicker1.Value, DateTimePicker2.Value, dscustomers.Tables(0).Rows(countcust).Item(0).ToString, Integer.Parse(dscustpro.Tables(0).Rows(countitem).Item(2).ToString), conFPM) + getTotalPremiumPerPolicyNumberPrincipalIDandBenefitStartDate(DateTimePicker1.Value, DateTimePicker2.Value, dscustomers.Tables(0).Rows(countcust).Item(0).ToString, Integer.Parse(dscustpro.Tables(0).Rows(countitem).Item(2).ToString), conFPM) + getMembershipFeeByGroupID(GetGroupIDByCustomerID(dscustomers.Tables(0).Rows(countcust).Item(0).ToString, conFPM), conFPM)
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "FPM Funeral"
                                End If
                            Else
                                If getValueBooleanOtherSystems("select * from SalesCommission where ProductAutoID = '" & dscustpro.Tables(0).Rows(countitem).Item(2).ToString & "' and Status = 'Claimed' and SystemName = '" & dsproducts.Tables(0).Rows(countpro).Item(1).ToString & "'", con) = False Then
                                    DataGridView1.Rows.Add()
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(0).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Funeral"
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustpro.Tables(0).Rows(countitem).Item(0).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = getTotalPremiumPerPolicyNumberPrincipalIDandCoverStartDate(DateTimePicker1.Value, DateTimePicker2.Value, dscustomers.Tables(0).Rows(countcust).Item(0).ToString, Integer.Parse(dscustpro.Tables(0).Rows(countitem).Item(2).ToString), conFPM) + getTotalPremiumPerPolicyNumberPrincipalIDandBenefitStartDate(DateTimePicker1.Value, DateTimePicker2.Value, dscustomers.Tables(0).Rows(countcust).Item(0).ToString, Integer.Parse(dscustpro.Tables(0).Rows(countitem).Item(2).ToString), conFPM) + getMembershipFeeByGroupID(GetGroupIDByCustomerID(dscustomers.Tables(0).Rows(countcust).Item(0).ToString, conFPM), conFPM)
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustpro.Tables(0).Rows(countitem).Item(2).ToString
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "FPM Funeral"
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        ElseIf cbcontacted.Checked = True Then
            Label3.Text = "New Customer Commission"
            dscustomers.Clear()
            dscustomers = getDataset("select * from SalesCustomers where ContactedDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "'", con)

            For countcust As Integer = 0 To dscustomers.Tables(0).Rows.Count - 1
                If getValueBooleanOtherSystems("select * from SalesCommission where Status = 'Claimed' and Product = 'Customer' and CustomerID = '" & dscustomers.Tables(0).Rows(countcust).Item(1).ToString & "'", con) = False Then
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ""
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = "Customer"
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ""
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = ""
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = getStringValue("select Fullname from SalesAgents where IDNumber = '" & dscustomers.Tables(0).Rows(countcust).Item(5).ToString & "'", con)
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = ""
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "EDM"
                End If
            Next
        End If
        MessageBox.Show("Done Running", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub cbproducts_CheckedChanged(sender As Object, e As EventArgs) Handles cbproducts.CheckedChanged
        If cbproducts.Checked = True Then
            cbproducts.Checked = True
            cbcontacted.Checked = False
        ElseIf cbproducts.Checked = False Then
            cbproducts.Checked = False
            cbcontacted.Checked = True
        End If
    End Sub

    Private Sub cbcontacted_CheckedChanged(sender As Object, e As EventArgs) Handles cbcontacted.CheckedChanged
        If cbcontacted.Checked = True Then
            cbcontacted.Checked = True
            cbproducts.Checked = False
        ElseIf cbcontacted.Checked = False Then
            cbcontacted.Checked = False
            cbproducts.Checked = True
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub SalesProducts_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        adddown = False
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

        If cbproducts.Checked = True Then
            Label3.Text = "Products Commission (History)"
            dscustomers.Clear()
            dscustomers = getDataset("select * from SalesCommission where ContactedDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and Status = 'Claimed' and Product not in ('Customer')", con)

            For countcust As Integer = 0 To dscustomers.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustomers.Tables(0).Rows(countcust).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustomers.Tables(0).Rows(countcust).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustomers.Tables(0).Rows(countcust).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = dscustomers.Tables(0).Rows(countcust).Item(12).ToString
            Next
        ElseIf cbcontacted.Checked = True Then
            Label3.Text = "Customer Commission (History)"
            dscustomers.Clear()
            dscustomers = getDataset("select * from SalesCommission where ContactedDate between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' and Status = 'Claimed' and Product = 'Customer'", con)

            For countcust As Integer = 0 To dscustomers.Tables(0).Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = dscustomers.Tables(0).Rows(countcust).Item(1).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = dscustomers.Tables(0).Rows(countcust).Item(2).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = dscustomers.Tables(0).Rows(countcust).Item(3).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = dscustomers.Tables(0).Rows(countcust).Item(4).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = dscustomers.Tables(0).Rows(countcust).Item(5).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = dscustomers.Tables(0).Rows(countcust).Item(6).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = dscustomers.Tables(0).Rows(countcust).Item(7).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = dscustomers.Tables(0).Rows(countcust).Item(8).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8).Value = dscustomers.Tables(0).Rows(countcust).Item(9).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = dscustomers.Tables(0).Rows(countcust).Item(10).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = dscustomers.Tables(0).Rows(countcust).Item(11).ToString
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = dscustomers.Tables(0).Rows(countcust).Item(12).ToString
            Next
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        DateTimePicker1.Value = getFirstDateOfMonth(DateTimePicker2.Value)
        DateTimePicker2.Value = getLastDateOfMonth(DateTimePicker2.Value)
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        If adddown = True Then
            If DataGridView1.Rows.Count > 0 Then
                If decision("Pay Commission for this product ?") = DialogResult.Yes Then
                    DataGridView2.Rows.Add()
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0).Value = DataGridView1.CurrentRow.Cells(0).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(1).Value = DataGridView1.CurrentRow.Cells(1).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(2).Value = DataGridView1.CurrentRow.Cells(2).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(3).Value = DataGridView1.CurrentRow.Cells(3).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(4).Value = DataGridView1.CurrentRow.Cells(4).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(5).Value = DataGridView1.CurrentRow.Cells(5).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(6).Value = DataGridView1.CurrentRow.Cells(6).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(7).Value = DataGridView1.CurrentRow.Cells(7).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(8).Value = DataGridView1.CurrentRow.Cells(8).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(9).Value = DataGridView1.CurrentRow.Cells(9).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(10).Value = DataGridView1.CurrentRow.Cells(10).Value
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(11).Value = DataGridView1.CurrentRow.Cells(11).Value
                    DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
                End If
            End If
        End If
    End Sub
End Class