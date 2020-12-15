Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource

Public Class SettlementBalance

    Dim Receipts, TotalRechargesAsReceipts, TotalContractualAmount, TotalBilling, PendingAllocations, TotalReceipts, PaidBills, balance As Decimal
    Dim CustomerID As String

    Dim crrpsl As New ReportDocument

    Dim proceed As Boolean
    Dim dscustomer, dssupervisor, dssettlementdetails As New DataSet
    Dim CustomerName As String
    Dim CreationDate, CalcEndDate, ValidityDate As Date

    Public datestring, towhom, re, thisserves, ifyouintend, thitoaccount, pleaseuse, formoreinfo, yours, supervisor, title, email, pleasebeinformed, pleaseensure, product, instalment, totalbalance, instalmentdescription, totalinstalments As String

    'balances
    Dim dsclearancedetails As New DataSet

    Dim dtsettlement As New DataTable
    Dim dtclearance As New DataTable

    Dim datetouse As Date
    Dim firstdateassigned As Boolean
    Dim reporttype As String

    Public Sub viewClearanceLetter()
        Dim a, b, c, d, e, f, g, h, i As String

        CustomerID = tbcustomerid.Text.Trim
        dscustomer = getCustomerDetails(tbcustomerid.Text.Trim)
        dssupervisor = getSupervisorDetails(supervisorid)

        If dscustomer.Tables(0).Rows.Count > 0 Then

            dtclearance.Columns.Add("Text1", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text2", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text3", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text4", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text5", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text6", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text7", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text8", Type.GetType("System.String"))
            dtclearance.Columns.Add("Text9", Type.GetType("System.String"))

            CustomerName = dscustomer.Tables(0).Rows(0).Item(0).ToString

            CreationDate = DateTimePicker1.Value

            dsclearancedetails = getClearanceDetails()

            a = DateTimePicker1.Value.Day & " " & MonthName(DateTimePicker1.Value.Month) & " " & DateTimePicker1.Value.Year
            b = dsclearancedetails.Tables(0).Rows(0).Item(1).ToString
            If getGender(tbcustomerid.Text.Trim) = "Male" Then
                c = dsclearancedetails.Tables(0).Rows(0).Item(2).ToString & " Mr" & " " & CustomerName & " " & dsclearancedetails.Tables(0).Rows(0).Item(3).ToString & " " & CustomerID
                d = dsclearancedetails.Tables(0).Rows(0).Item(4).ToString & " Mr" & " " & CustomerName & " " & dsclearancedetails.Tables(0).Rows(0).Item(5).ToString & " " & MonthName(dtpclearance.Value.Month) & " " & dtpclearance.Value.Year & ": " & dsclearancedetails.Tables(0).Rows(0).Item(6).ToString
            ElseIf getGender(tbcustomerid.Text.Trim) = "Female" Then
                c = dsclearancedetails.Tables(0).Rows(0).Item(2).ToString & " Ms" & " " & CustomerName & " " & dsclearancedetails.Tables(0).Rows(0).Item(3).ToString & " " & CustomerID
                d = dsclearancedetails.Tables(0).Rows(0).Item(4).ToString & " Ms" & " " & CustomerName & " " & dsclearancedetails.Tables(0).Rows(0).Item(5).ToString & " " & MonthName(dtpclearance.Value.Month) & " " & dtpclearance.Value.Year & ": " & dsclearancedetails.Tables(0).Rows(0).Item(6).ToString
            End If
            e = dsclearancedetails.Tables(0).Rows(0).Item(7).ToString
            f = dsclearancedetails.Tables(0).Rows(0).Item(8).ToString
            g = dssupervisor.Tables(0).Rows(0).Item(1).ToString
            h = dssupervisor.Tables(0).Rows(0).Item(2).ToString
            i = dssupervisor.Tables(0).Rows(0).Item(3).ToString

            dtclearance.Rows.Add(a, b, c, d, e, f, g, h, i)

            Dim objRpt As New Clearance_Letter
            objRpt.SetDataSource(dtclearance)
            Me.CrystalReportViewer1.ReportSource = objRpt
            Me.CrystalReportViewer1.DisplayBackgroundEdge = True
            Me.CrystalReportViewer1.DisplayToolbar = True
            Me.CrystalReportViewer1.DisplayGroupTree = True
            Me.CrystalReportViewer1.Refresh()
            Me.Cursor = Windows.Forms.Cursors.Default
            Me.TabControl1.SelectTab(1)
            crrpsl = objRpt
        Else
            MessageBox.Show("Customer does not exist", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub generateSettlementReport()

        CustomerID = tbcustomerid.Text.Trim
        dscustomer = getCustomerDetails(tbcustomerid.Text.Trim)
        dssupervisor = getSupervisorDetails(supervisorid)

        If dscustomer.Tables(0).Rows.Count > 0 Then

            dtsettlement.Columns.Add("datestring", Type.GetType("System.String"))
            dtsettlement.Columns.Add("towhom", Type.GetType("System.String"))
            dtsettlement.Columns.Add("re", Type.GetType("System.String"))
            dtsettlement.Columns.Add("thisserves", Type.GetType("System.String"))
            dtsettlement.Columns.Add("ifyouintend", Type.GetType("System.String"))
            dtsettlement.Columns.Add("thitoaccount", Type.GetType("System.String"))
            dtsettlement.Columns.Add("pleaseuse", Type.GetType("System.String"))
            dtsettlement.Columns.Add("formoreinfo", Type.GetType("System.String"))
            dtsettlement.Columns.Add("yours", Type.GetType("System.String"))
            dtsettlement.Columns.Add("supervisor", Type.GetType("System.String"))
            dtsettlement.Columns.Add("title", Type.GetType("System.String"))
            dtsettlement.Columns.Add("email", Type.GetType("System.String"))
            dtsettlement.Columns.Add("pleasebeinformed", Type.GetType("System.String"))
            dtsettlement.Columns.Add("pleaseensure", Type.GetType("System.String"))
            dtsettlement.Columns.Add("product", Type.GetType("System.String"))
            dtsettlement.Columns.Add("instalment", Type.GetType("System.String"))
            dtsettlement.Columns.Add("totalbalance", Type.GetType("System.String"))
            dtsettlement.Columns.Add("instalmentdescription", Type.GetType("System.String"))
            dtsettlement.Columns.Add("totalinstalment", Type.GetType("System.String"))

            CustomerName = dscustomer.Tables(0).Rows(0).Item(0).ToString

            CreationDate = DateTimePicker1.Value

            If checkIfLoanIsAReducingBalanceLoan(getLastActiveLoanID(CustomerID, "Normal", conMLMAdvances), conMLMAdvances) = True Then
                CalcEndDate = getLastDateOfMonth(CreationDate)
            Else
                If DateTimePicker1.Value.Day > getCutOfDate() Then
                    'add a month to the date
                    Dim dt As DateTime = DateTimePicker1.Value.AddMonths(1)
                    Dim year As Integer = dt.Year
                    Dim month As Integer = dt.Month
                    Dim days As Integer = DateTime.DaysInMonth(year, month)
                    Dim lastDayOfMonth As New DateTime(year, month, days)
                    CalcEndDate = Date.Parse(lastDayOfMonth)

                Else
                    Dim dt As DateTime = DateTimePicker1.Value
                    Dim year As Integer = dt.Year
                    Dim month As Integer = dt.Month
                    Dim days As Integer = DateTime.DaysInMonth(year, month)
                    Dim lastDayOfMonth As New DateTime(year, month, days)
                    CalcEndDate = Date.Parse(lastDayOfMonth)
                End If
            End If

            dssettlementdetails = getSettlementDetails()

            datestring = DateTimePicker1.Value.Day & " " & MonthName(DateTimePicker1.Value.Month) & " " & DateTimePicker1.Value.Year
            towhom = dssettlementdetails.Tables(0).Rows(0).Item(1).ToString
            If getGender(tbcustomerid.Text.Trim) = "Male" Then
                re = dssettlementdetails.Tables(0).Rows(0).Item(2).ToString & " Mr" & " " & CustomerName & " " & dssettlementdetails.Tables(0).Rows(0).Item(3).ToString & " " & CustomerID
                thisserves = dssettlementdetails.Tables(0).Rows(0).Item(4).ToString & " Mr" & " " & CustomerName & " " & dssettlementdetails.Tables(0).Rows(0).Item(5).ToString & " " & dssettlementdetails.Tables(0).Rows(0).Item(6).ToString & " " & MonthName(CalcEndDate.Month) & " " & CalcEndDate.Year & ";"
            ElseIf getGender(tbcustomerid.Text.Trim) = "Female" Then
                re = dssettlementdetails.Tables(0).Rows(0).Item(2).ToString & " Ms" & " " & CustomerName & " " & dssettlementdetails.Tables(0).Rows(0).Item(3).ToString & " " & CustomerID
                thisserves = dssettlementdetails.Tables(0).Rows(0).Item(4).ToString & " Ms" & " " & CustomerName & " " & dssettlementdetails.Tables(0).Rows(0).Item(5).ToString & " " & dssettlementdetails.Tables(0).Rows(0).Item(6).ToString & " " & MonthName(CalcEndDate.Month) & " " & CalcEndDate.Year & ";"
            End If

            ifyouintend = dssettlementdetails.Tables(0).Rows(0).Item(7).ToString
            thitoaccount = dssettlementdetails.Tables(0).Rows(0).Item(8).ToString
            pleaseuse = dssettlementdetails.Tables(0).Rows(0).Item(9).ToString
            formoreinfo = dssettlementdetails.Tables(0).Rows(0).Item(10).ToString
            yours = dssettlementdetails.Tables(0).Rows(0).Item(11).ToString
            supervisor = dssupervisor.Tables(0).Rows(0).Item(1).ToString
            title = dssupervisor.Tables(0).Rows(0).Item(2).ToString
            email = dssupervisor.Tables(0).Rows(0).Item(3).ToString
            pleasebeinformed = dssettlementdetails.Tables(0).Rows(0).Item(15).ToString & " " & CalcEndDate.Day & " " & MonthName(CalcEndDate.Month) & " " & CalcEndDate.Year
            'pleaseensure = dssettlementdetails.Tables(0).Rows(0).Item(18).ToString & " " & dssettlementdetails.Tables(0).Rows(0).Item(20).ToString & " " & MonthName(CalcEndDate.AddMonths(1).Month) & " " & CalcEndDate.AddMonths(1).Year & " " & dssettlementdetails.Tables(0).Rows(0).Item(19).ToString

            Dim vdate As New DateTime(CalcEndDate.AddMonths(1).Year, CalcEndDate.AddMonths(1).Month, Integer.Parse(dssettlementdetails.Tables(0).Rows(0).Item(20).ToString))
            ValidityDate = vdate

            For Each row As DataGridViewRow In DataGridView1.Rows
                If Boolean.Parse(row.Cells(5).Value) = True Then
                    product = row.Cells(1).Value
                    instalment = ": " + getCurrency(Decimal.Parse(row.Cells(3).Value))
                    instalmentdescription = ": " + getCurrency(Decimal.Parse(row.Cells(4).Value)) + " monthly instalment(s)"
                    dtsettlement.Rows.Add(datestring, towhom, re, thisserves, ifyouintend, thitoaccount, pleaseuse, formoreinfo, yours, supervisor, title, email, pleasebeinformed, pleaseensure, product, instalment, totalbalance, instalmentdescription, totalinstalments)
                End If
            Next

            Dim objRpt As New SettlementDetails2
            objRpt.SetDataSource(dtsettlement)
            Me.CrystalReportViewer1.ReportSource = objRpt
            Me.CrystalReportViewer1.DisplayBackgroundEdge = True
            Me.CrystalReportViewer1.DisplayToolbar = True
            Me.CrystalReportViewer1.DisplayGroupTree = True
            Me.CrystalReportViewer1.Refresh()
            Me.Cursor = Windows.Forms.Cursors.Default
            Me.TabControl1.SelectTab(1)
            crrpsl = objRpt
            Button2.Enabled = True
        Else
            MessageBox.Show("Customer does not exist", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub loadProducts()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,Products,HasSystem,'0' as SettlementBalance,'0' as Instalment,HasSystem as Include,'0' as LedgerBalance,'' as Comment1,'' as Comment2,'' as LastDateOfDeduction from Products where UseInSettlement = 'Yes'"

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
                For Each col As DataGridViewColumn In DataGridView1.Columns
                    If col.Index = 3 Or col.Index = 4 Or col.Index = 5 Then
                        If col.Index = 3 Then
                            col.DefaultCellStyle.BackColor = Color.LightGray
                        End If
                    Else
                        col.ReadOnly = True
                    End If
                Next

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Boolean.Parse(row.Cells(2).Value) = True Then
                        DataGridView1(3, row.Index).ReadOnly = True
                    Else
                        DataGridView1(3, row.Index).ReadOnly = False
                    End If
                    row.Cells(5).Value = False
                Next

            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            datetouse = Date.Now
            Button1.Enabled = False
            Button5.Enabled = True
            If checkCustomerInDatabase(tbcustomerid.Text.Trim) = True Then
                Label3.Text = "Warning : "
                If CheckIfCustomerPaysForPeopleInFPM(tbcustomerid.Text.Trim, conFPM) = True Then
                    MessageBox.Show("Please update KYC for Bereavement fund", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Else
                Label3.Text = "Warning : " & "Customer does not exist in the database. Please add the customer before you proceed"
            End If

            loadProducts()
            CustomerID = tbcustomerid.Text.Trim
            Label2.Text = 0.0
            Label4.Text = 0.0
            For Each row As DataGridViewRow In DataGridView1.Rows
                If getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Loan" Then
                    row.Cells(6).Value = getGeneralLedgerBalance(CustomerID, "Normal", conMLMAdvances)
                    If getGeneralLedgerBalance(CustomerID, "Normal", conMLMAdvances) > 0 Then
                        row.Cells(3).Value = getSettlementMLMNormal(CustomerID, DateTimePicker1.Value, conMLMAdvances)
                        row.Cells(4).Value = "0"
                        row.Cells(7).Value = "Has Balance"
                        Label2.Text = Decimal.Parse(Label2.Text.Trim) + getSettlementMLMNormal(CustomerID, DateTimePicker1.Value, conMLMAdvances)
                    Else
                        row.Cells(3).Value = "0"
                        row.Cells(4).Value = "0"
                    End If
                    row.Cells(9).Value = getLastDateOfDeductionMLM(CustomerID, "Normal", conMLMAdvances)
                ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Educational" Then
                    row.Cells(6).Value = getGeneralLedgerBalance(CustomerID, "Normal", conMLMEducational)
                    If getGeneralLedgerBalance(CustomerID, "Normal", conMLMEducational) > 0 Then
                        row.Cells(3).Value = getSettlementMLMNormal(CustomerID, DateTimePicker1.Value, conMLMEducational)
                        row.Cells(4).Value = "0"
                        row.Cells(7).Value = "Has Balance"
                        Label2.Text = Decimal.Parse(Label2.Text.Trim) + getSettlementMLMNormal(CustomerID, DateTimePicker1.Value, conMLMEducational)
                    Else
                        row.Cells(3).Value = "0"
                        row.Cells(4).Value = "0"
                    End If
                    row.Cells(9).Value = getLastDateOfDeductionMLM(CustomerID, "Normal", conMLMEducational)
                ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Quick" Then
                    row.Cells(6).Value = getGeneralLedgerBalance(CustomerID, "Quick", conMLMAdvances)
                    If getGeneralLedgerBalance(CustomerID, "Quick", conMLMAdvances) > 0 Then
                        row.Cells(3).Value = getSettlementMLMQuick(CustomerID, DateTimePicker1.Value, conMLMAdvances)
                        row.Cells(4).Value = "0"
                        row.Cells(7).Value = "Has Balance"
                        Label2.Text = Decimal.Parse(Label2.Text.Trim) + getSettlementMLMQuick(CustomerID, DateTimePicker1.Value, conMLMAdvances)
                    Else
                        row.Cells(3).Value = "0"
                        row.Cells(4).Value = "0"
                    End If
                    row.Cells(9).Value = getLastDateOfDeductionMLM(CustomerID, "Quick", conMLMAdvances)
                ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "RM Orange" Then
                    row.Cells(6).Value = getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMOrange)
                    If getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMOrange) > 0 Then
                        row.Cells(3).Value = getSettlementRM(CustomerID, DateTimePicker1.Value, conRMOrange)
                        row.Cells(4).Value = "0"
                        row.Cells(7).Value = "Has Balance"
                        Label2.Text = Decimal.Parse(Label2.Text.Trim) + getSettlementRM(CustomerID, DateTimePicker1.Value, conRMOrange)
                    Else
                        row.Cells(3).Value = "0"
                        row.Cells(4).Value = "0"
                    End If
                    row.Cells(9).Value = getLastDateOfDeductionRM(CustomerID, conRMOrange)
                ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "RM BeMobile" Then
                    row.Cells(6).Value = getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMBeMobile)
                    If getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMBeMobile) > 0 Then
                        row.Cells(3).Value = getSettlementRM(CustomerID, DateTimePicker1.Value, conRMBeMobile)
                        row.Cells(4).Value = "0"
                        row.Cells(7).Value = "Has Balance"
                        Label2.Text = Decimal.Parse(Label2.Text.Trim) + getSettlementRM(CustomerID, DateTimePicker1.Value, conRMBeMobile)
                    Else
                        row.Cells(3).Value = "0"
                        row.Cells(4).Value = "0"
                    End If
                    row.Cells(9).Value = getLastDateOfDeductionRM(CustomerID, conRMBeMobile)
                ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "FPM" Then
                    'row.Cells(9).Value = getLastDateOfDeductionFPM(CustomerID, conFPM)
                End If

                row.Cells(4).Value = getInstalment(tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value))
                If Decimal.Parse(row.Cells(4).Value) > 0 Then
                    row.Cells(5).Value = True
                End If

                firstdateassigned = False
                If row.Cells(9).Value = "" Then

                Else
                    If firstdateassigned = False Then
                        datetouse = Date.Parse(row.Cells(9).Value).Date
                        firstdateassigned = True
                    Else
                        If datetouse > Date.Parse(row.Cells(9).Value) Then
                            datetouse = Date.Parse(row.Cells(9).Value).Date
                        Else

                        End If
                    End If
                End If

            Next

            dtpclearance.Value = datetouse

            updateBalance("Get")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Public Function getGeneralLedgerBalance(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim str As Decimal
        If loantype = "Quick" Then
            Dim qry As String = "select Sum(Amount) from LedgerEntry where CustomerID = '" & custid & "' and LoanType = '" & loantype & "'"

            Dim cmd As New SqlClient.SqlCommand(qry, con)
            Dim ds As New DataSet
            Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

            Try
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim ctr As Integer = 0
                SQLAdapter.Fill(ds, "LedgerEntry")

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                If ds.Tables(0).Rows.Count > 0 Then
                    If ds.Tables(0).Rows(0).IsNull(0) = True Then
                        str = 0
                    Else
                        str = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                    End If
                Else
                    str = 0
                End If
            Catch ex As Exception
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                MessageBox.Show(ex.Message, "getGeneralLedgerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBox.Show(ex.StackTrace, "getGeneralLedgerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                'ds.Clear()
                con.Close()
            End Try
        ElseIf loantype = "Normal" Then
            Dim qry As String = "select Sum(Amount) from LedgerEntry where CustomerID = '" & custid & "' and LoanType = '" & loantype & "'"

            Dim cmd As New SqlClient.SqlCommand(qry, con)
            Dim ds As New DataSet
            Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

            Try
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim ctr As Integer = 0
                SQLAdapter.Fill(ds, "LedgerEntry")

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                If ds.Tables(0).Rows.Count > 0 Then
                    If ds.Tables(0).Rows(0).IsNull(0) = True Then
                        str = 0
                    Else
                        str = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                    End If
                Else
                    str = 0
                End If
            Catch ex As Exception
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                MessageBox.Show(ex.Message, "getGeneralLedgerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBox.Show(ex.StackTrace, "getGeneralLedgerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                'ds.Clear()
                con.Close()
            End Try
        End If

        Return str
    End Function

    Public Function getSettlementRM(ByVal CustomerID As String, ByVal idate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select SettlememntAmount from SettlementBalances where CustomerID = '" & CustomerID & "' and IssuedDate = '" & idate & "' order by AutoID DESC"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSettlementRM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSettlementRM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getSettlementMLMNormal(ByVal CustomerID As String, ByVal CreationDate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select LoanBalance from SettlementLetters where CustomerID = '" & CustomerID & "' and CreationDate = '" & CreationDate & "' order by AutoID DESC"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSettlementMLMNormal()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSettlementMLMNormal()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getSettlementMLMQuick(ByVal CustomerID As String, ByVal CreationDate As Date, ByVal con As SqlClient.SqlConnection) As Decimal
        Dim qry As String = "select QuickBalance from SettlementLetters where CustomerID = '" & CustomerID & "' and CreationDate = '" & CreationDate & "' order by AutoID DESC"

        Dim perc As Decimal = 0

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = 0
                Else
                    perc = Decimal.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                perc = 0
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getSettlementMLMQuick()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getSettlementMLMQuick()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        updateBalance("Refresh")
    End Sub

    Public Sub updateBalance(ByVal whichbutton As String)
        Try
            Button3.Enabled = True
            Label2.Text = 0
            Label4.Text = 0
            For Each row As DataGridViewRow In DataGridView1.Rows

                If Decimal.Parse(row.Cells(3).Value) > 0 Or Decimal.Parse(row.Cells(4).Value) > 0 Then
                    If whichbutton = "Refresh" Then
                        row.Cells(5).Value = True
                    End If
                End If

                If Boolean.Parse(row.Cells(5).Value) = True Then
                    Label2.Text = Decimal.Parse(Label2.Text.Trim) + Decimal.Parse(row.Cells(3).Value)
                    Label4.Text = Decimal.Parse(Label4.Text.Trim) + Decimal.Parse(row.Cells(4).Value)
                End If

                If Boolean.Parse(row.Cells(2).Value) = True Then
                    If Decimal.Parse(row.Cells(6).Value) > 0 Then
                        If (Decimal.Parse(row.Cells(3).Value) > 0) And (Decimal.Parse(row.Cells(4).Value) > 0) Then
                            row.Cells(8).Value = "Requested"
                        ElseIf (Decimal.Parse(row.Cells(3).Value) > 0) Then
                            row.Cells(8).Value = "Requested But No Instalment"
                            Button3.Enabled = False
                        ElseIf (Decimal.Parse(row.Cells(4).Value) > 0) Then
                            If Boolean.Parse(row.Cells(2).Value) = True Then
                                row.Cells(8).Value = "Not Requested"
                                Button3.Enabled = False
                            Else
                                row.Cells(8).Value = "Requested"
                            End If
                        Else
                            row.Cells(8).Value = "Not Requested"
                            Button3.Enabled = False
                        End If
                    Else
                        row.Cells(8).Value = ""
                    End If
                ElseIf Boolean.Parse(row.Cells(2).Value) = False Then

                    If (Decimal.Parse(row.Cells(3).Value) > 0) And (Decimal.Parse(row.Cells(4).Value) > 0) Then
                        row.Cells(8).Value = "Added"
                    ElseIf (Decimal.Parse(row.Cells(3).Value) > 0) Then
                        row.Cells(8).Value = "Added But No Instalment"
                        Button3.Enabled = False
                    ElseIf (Decimal.Parse(row.Cells(4).Value) > 0) Then
                        row.Cells(8).Value = "Requested"
                    Else
                        row.Cells(8).Value = ""
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Make sure there is not empty cell on the settlement balance column", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        reporttype = "Settlement"
        proceed = True
        updateBalance("Proceed")

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(8).Value = "" Or row.Cells(8).Value = "Requested" Or row.Cells(8).Value = "Added" Then

            Else
                proceed = False
            End If
        Next

        If proceed = True Then
            If checkCustomerInDatabase(tbcustomerid.Text.Trim) = True Then
                If Decimal.Parse(Label2.Text.Trim) > 0 Then
                    Dim vrl As New RequestSettlement
                    vrl.MdiParent = mform
                    vrl.sb = Me
                    vrl.tbcustomerid.Text = tbcustomerid.Text
                    vrl.Label2.Text = getCustomerName(tbcustomerid.Text)
                    vrl.Show()
                Else
                    MessageBox.Show("Balance should be greater than 0", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Customer Does Not Exist", "getGeneralLedgerBalance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Correct the error first to proceed", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub SettlementBalance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'If checkCutOfDate() = False Then
        '    MessageBox.Show("Cut Off Date has not been set yet. Contact Supervisor", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    GroupBox1.Enabled = False
        'Else
        '    If checkIfMonthHasCutOfDate(getFirstDateOfMonth(Date.Now), getLastDateOfMonth(Date.Now)) = True Then
        GroupBox1.Enabled = True
        '    Else
        'MessageBox.Show("Cut Off Date has not been confirmed yet. Contact Supervisor", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'GroupBox1.Enabled = False
        '    End If
        'End If
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Try
            For Each ctrl As Control In CrystalReportViewer1.Controls
                If TypeOf ctrl Is Windows.Forms.ToolStrip Then
                    Dim btnNew As New ToolStripButton
                    btnNew.Text = "Print Report"
                    ''btnNew.Image = Image.FromFile("../../Resources/printer.ico")
                    CType(ctrl, ToolStrip).Items.Add(btnNew)
                    AddHandler btnNew.Click, AddressOf Printdomf
                    'index = 5
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Printdomf(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If decision("Are you sure you want to print the letter and save details?") = DialogResult.Yes Then
                saveBalances()
                crrpsl.PrintToPrinter(1, False, 0, 0)
                MessageBox.Show("Done printing", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        reporttype = "Clearance"
        proceed = True

        For Each row As DataGridViewRow In DataGridView1.Rows
            If (Decimal.Parse(row.Cells(6).Value) > 0) Or (Decimal.Parse(row.Cells(6).Value) < 0) Then
                proceed = False
            Else

            End If
        Next

        If proceed = True Then
            If checkCustomerInDatabase(tbcustomerid.Text.Trim) = True Then
                saveBalances()
                viewClearanceLetter()
            Else
                MessageBox.Show("Customer Does Not Exist", "getClearance()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Customer Has Balance", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub saveBalances()
        Dim ndate As Date
        ndate = Date.Now

        If checkCustomerInDatabase(tbcustomerid.Text.Trim) = True Then
            If reporttype = "Clearance" Then
                saveSettlementBalance(tbcustomerid.Text.Trim, 0, 0, Date.Now, logedinname, Date.Now, Date.Now, "New", "Clear:" + tbcustomerid.Text.Trim + ":" + ndate)
            ElseIf reporttype = "Settlement" Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If (Boolean.Parse(row.Cells(5).Value) = True) Then
                        If Decimal.Parse(row.Cells(3).Value) > 0 Then
                            saveSettlementBalance(tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value), Decimal.Parse(row.Cells(3).Value), Date.Now, logedinname, Date.Now, Date.Now, "New", "Settle:" + tbcustomerid.Text.Trim + ":" + ndate)
                        End If
                    End If
                Next
            End If
        Else
            MessageBox.Show("Customer Does Not Exist", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Sub displayInstalments()
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.Cells(4).Value = getInstalment(tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value))
            If Decimal.Parse(row.Cells(4).Value) > 0 Then
                row.Cells(5).Value = True
            End If
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        displayInstalments()
    End Sub

    Public Function getLastDateOfDeductionMLM(ByVal customerid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select TransactionDate from LedgerEntry where CustomerID = '" & customerid & "' and Reference in (select Reference from TotalReceiptReferences where Reference not in ('RE','REr')) and LoanType = '" & loantype & "' order by TransactionDate DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString).Date
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastDateOfDeductionMLM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastDateOfDeductionMLM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastDateOfDeductionRM(ByVal customerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select TransactionDate from CustomerLedgerEntry where TransactionDescription not in ('Billing','Refund') and CustomerID = '" & customerid & "' order by TransactionDate DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString).Date
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastDateOfDeductionRM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastDateOfDeductionRM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastDateOfDeductionFPM(ByVal payerid As String, ByVal con As SqlClient.SqlConnection) As String
        Dim qry As String = "select TransactionDate from CustomerReceipts where PayerID = '" & payerid & "' and Batch not in ('UNP') and Status in ('Posted','Approved') order by TransactionDate DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) Then
                    perc = ""
                Else
                    perc = Date.Parse(ds.Tables(0).Rows(0).Item(0).ToString).Date
                End If
            Else
                perc = ""
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getLastDateOfDeductionFPM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastDateOfDeductionFPM()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Function checkIfLoanIsAReducingBalanceLoan(ByVal loanid As Integer, ByVal con As SqlClient.SqlConnection) As Boolean
        Dim qry As String = "select Top 1 * from LoanInterestTypes where AutoID in (select LoanInterestID from LoanInterest where LoanID = '" & loanid & "') and LoanMethod = 'Reducing Balance'"
        Dim ds As New DataSet
        Dim perc As Boolean = False
        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0

            SQLAdapter.Fill(ds, "CustomerLedgerEntry")
            If ds.Tables(0).Rows.Count > 0 Then
                perc = True
            Else
                perc = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.StackTrace, "checkIfLoanIsAReducingBalanceLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.Message, "checkIfLoanIsAReducingBalanceLoan()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
        Return perc
    End Function

    Public Function getLastActiveLoanID(ByVal custid As String, ByVal loantype As String, ByVal con As SqlClient.SqlConnection) As Integer
        Dim qry As String = "select AutoLoanID from LoanDetails where Customer_ID = '" & custid & "' and LoanDescription = '" & loantype & "' and LoanStatus in ('Active') order by AutoLoanID DESC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "LoanDetails")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).IsNull(0) = True Then
                    Return 0
                Else
                    Return Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString)
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "getLastActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getLastActiveLoanID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'ds.Clear()
            con.Close()
        End Try
    End Function

End Class