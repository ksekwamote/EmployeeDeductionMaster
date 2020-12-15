Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource

Public Class ReceiptingCustomer

    Dim Receipts, TotalRechargesAsReceipts, TotalContractualAmount, TotalBilling, PendingAllocations, TotalReceipts, PaidBills, balance As Decimal
    Dim CustomerID As String

    Dim crrpl As New ReportDocument

    Dim proceed As Boolean
    Dim dscustomer, dssupervisor, dssettlementdetails As New DataSet
    Dim CustomerName As String
    Dim batchname As String

    Dim status As String

    Dim go As Boolean = False

    Public receiptbookid As Integer

    Dim PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, PaymentMode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount As String

    Private Sub loadProducts(ByVal CustomerID As String)
        DataGridView1.DataSource = Nothing
        ''Dim qry As String = "select AutoID,Products,HasSystem,'0' as LedgerBalance,'0' as Receipt,'' as ReceiptType,'N/A' as Comment from Products where UseInReceipts = 'Yes'"
        Dim qry As String = "select AutoID,Products,HasSystem,'0' as LedgerBalance,'0' as Receipt,'' as ReceiptType,'N/A' as Comment from Products where UseInReceipts = 'Yes' and AutoID in (select ProductID from EmployerProducts where Status = 'Active' and EmployerID in (select EmployerID from CustomerDetails where CustomerID = '" & CustomerID & "'))"
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
                    If col.Index = 3 Or col.Index = 4 Then
                        If col.Index = 3 Then
                            col.DefaultCellStyle.BackColor = Color.LightGray
                            col.ReadOnly = True
                        End If
                    ElseIf col.Index = 6 Then
                        col.ReadOnly = False
                    Else
                        col.ReadOnly = True
                    End If
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
            Button1.Enabled = False
            If checkCustomerInDatabase(tbcustomerid.Text.Trim) = True Then
                Label3.Text = "Warning : "
                If CheckIfCustomerPaysForPeopleInFPM(tbcustomerid.Text.Trim, conFPM) = True Then
                    MessageBox.Show("Please update KYC for Bereavement fund", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If

                Button2.Enabled = True
                Button3.Enabled = True
                loadProducts(tbcustomerid.Text.Trim)
                CustomerID = tbcustomerid.Text.Trim
                tbpayerid.Text = tbcustomerid.Text.Trim
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Loan" Then
                        row.Cells(3).Value = getGeneralLedgerBalance(CustomerID, "Normal", conMLMAdvances)
                    ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Educational" Then
                        row.Cells(3).Value = getGeneralLedgerBalance(CustomerID, "Normal", conMLMEducational)
                    ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "MLM Quick" Then
                        row.Cells(3).Value = getGeneralLedgerBalance(CustomerID, "Quick", conMLMAdvances)
                    ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "RM Orange" Then
                        row.Cells(3).Value = getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMOrange)
                    ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "RM BeMobile" Then
                        row.Cells(3).Value = getBalancesFromRM(CustomerID, DateTimePicker1.Value, conRMBeMobile)
                    ElseIf getProductSystem(Integer.Parse(row.Cells(0).Value)) = "FPM" Then
                        row.Cells(3).Value = getPayerBalance(CustomerID, conFPM)
                    End If
                Next

                tbpayername.Text = getCustomerName(tbcustomerid.Text.Trim)

                If DataGridView1.Rows.Count > 0 Then

                Else
                    Label3.Text = "Warning : " & "Customer Employer (" & getEmployerNameByCustomerID(tbcustomerid.Text.Trim) & ") Does Not Have Products. Contact Back Office"
                End If
            Else
                Label3.Text = "Warning : " & "Customer does not exist in the database. Please add the customer before you proceed"
                GroupBox2.Enabled = False
                Button3.Enabled = False
                Button2.Enabled = False
            End If
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        updateTotalReceipt("Refresh")
    End Sub

    Public Sub updateTotalReceipt(ByVal whichbutton As String)
        Try
            tbtotalreceipt.Text = 0.0
            For Each row As DataGridViewRow In DataGridView1.Rows
                tbtotalreceipt.Text = Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim) + Decimal.Parse(row.Cells(4).Value), 2)
                'MessageBox.Show(tbtotalreceipt.Text, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Next
        Catch ex As Exception
            MessageBox.Show("Make sure there is not empty cell on the settlement balance column", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tbtotalreceipt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtotalreceipt.TextChanged
        Label2.Text = getCurrency(Math.Round(Decimal.Parse(tbtotalreceipt.Text.Trim), 2))
        tbamount.Text = Decimal.Parse(tbtotalreceipt.Text.Trim)
    End Sub

    Public Sub proceedToSaveReceipt(ByVal authorizer As String, ByVal companybranchid As Integer, ByVal StatusStr As String)

        status = StatusStr

        batchname = tbpayerid.Text.Trim & ":" & Date.Now
        If checkIfToModeIsDirectDeposit(Integer.Parse(tbmodeid.Text.Trim)) = True Then

        Else
            savePaymentDetails(tbreceivedfrom.Text.Trim, tbreceivedidnumber.Text.Trim, tbreceivedcellnumber.Text.Trim, tbpayername.Text.Trim, tbpayerid.Text.Trim, tbpayercellnumber.Text.Trim, Date.Now.Date, Decimal.Parse(tbamount.Text.Trim), batchname, tbdescription.Text.Trim, tbreference.Text.Trim)
        End If
        saveReceiptPayerInfor(tbpayerid.Text.Trim, tbpayercellnumber.Text.Trim, tbreceivedidnumber.Text.Trim, tbreceivedfrom.Text.Trim, tbreceivedcellnumber.Text.Trim, tbdescription.Text.Trim, batchname, tbpayername.Text.Trim)
        saveReceipt(tbpayerid.Text.Trim, Decimal.Parse(tbamount.Text.Trim), Integer.Parse(tbmodeid.Text.Trim), dtptransactiondate.Value, authorizer, Date.Now, Date.Now, "", Date.Now, Date.Now, batchname, status, "", "", tbreference.Text.Trim, companybranchid, receiptbookid)
        wrapper.saveAction("Receipt Saved", Date.Now, authorizer, tbpayerid.Text.Trim)
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Decimal.Parse(row.Cells(4).Value) > 0 Then
                saveReceiptBreakdown(tbpayerid.Text.Trim, Decimal.Parse(row.Cells(4).Value), dtptransactiondate.Value, Integer.Parse(row.Cells(0).Value), batchname, "", status, "", row.Cells(5).Value, "", Date.Now, tbreference.Text.Trim, Integer.Parse(tbmodeid.Text.Trim), companybranchid, receiptbookid, row.Cells(6).Value)
            End If
        Next
        MessageBox.Show("Done Saving Receipt", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        If checkIfToModeIsDirectDeposit(Integer.Parse(tbmodeid.Text.Trim)) = True Then
            Close()
        Else
            loadReport()
            Me.TabControl1.SelectTab(2)
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If Decimal.Parse(DataGridView1.CurrentRow.Cells(4).Value) > 0.0 Then
                Dim vrl As New ReceiptType
                vrl.MdiParent = mform
                vrl.rc = Me
                vrl.rowindex = Integer.Parse(DataGridView1.CurrentRow.Index)
                vrl.Show()
            End If
        End If
    End Sub

    Public Sub loadmodes()
        Dim qry As String = "select * from ModeOfPayment"

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
                Me.cbmodename.DataSource = ds
                Me.cbmodename.DisplayMember = "Emp.ModeName"
                Me.cbmodename.ValueMember = "Emp.ModeID"
                Me.cbmodename.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadmodes()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ReceiptingCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadmodes()
        tbmodename.Text = ""
        tbmodeid.Text = ""
        go = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        updateTotalReceipt("Refresh")
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then

            If tbreceivedfrom.Text.Trim = "" Or tbreceivedidnumber.Text.Trim = "" Or tbreceivedcellnumber.Text.Trim = "" Or tbpayercellnumber.Text.Trim = "" Or tbpayerid.Text.Trim = "" Or tbdescription.Text.Trim = "" Or tbamount.Text.Trim = "" Or tbmodename.Text.Trim = "" Or tbmodeid.Text.Trim = "" Or tbreference.Text.Trim = "" Then
                MessageBox.Show("Enter all details first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If Decimal.Parse(tbamount.Text.Trim) > 0 Then

                    go = True
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        If Decimal.Parse(row.Cells(4).Value) > 0 Then
                            If row.Cells(5).Value = "" Then
                                go = False
                            End If
                        End If
                    Next

                    If go = True Then
                        If decision("Is this the Amount Paid, " & getCurrency(Math.Round(Decimal.Parse(tbamount.Text.Trim), 2)) & " ?") = DialogResult.Yes Then
                            If cbbatch.Checked = True Then
                                proceedToSaveReceipt(logedinname, branchid, "Batch")
                            Else
                                Dim vrl As New ReceiptAuthorization
                                vrl.MdiParent = mform
                                vrl.rc = Me
                                vrl.whichside = "Normal"
                                vrl.Show()
                            End If
                        End If
                    ElseIf go = False Then
                        MessageBox.Show("Please ensure that you have put the Receipt Type", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBox.Show("Cannot proceed. Amount should be greater than 0", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Decimal.Parse(tbamount.Text.Trim) > 0 Then
            go = True
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Decimal.Parse(row.Cells(4).Value) > 0 Then
                    If row.Cells(5).Value = "" Then
                        go = False
                    End If
                End If
            Next

            If go = True Then
                If decision("Is this the Amount Paid, " & getCurrency(Math.Round(Decimal.Parse(tbamount.Text.Trim), 2)) & " ?") = DialogResult.Yes Then
                    Me.TabControl1.SelectTab(1)

                    batchname = "Grouping:" & Date.Now
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        If Decimal.Parse(row.Cells(4).Value) > 0 Then
                            If checkIfProductHasAGroup(Integer.Parse(row.Cells(0).Value), getEmployerIDByCustomerID(tbcustomerid.Text.Trim)) = True Then
                                If checkIfCustomerGroupIsForMerging(tbcustomerid.Text.Trim, conFPM) = True Then
                                    saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(3).Value), Decimal.Parse(row.Cells(4).Value), getProductGroupingName(Integer.Parse(row.Cells(0).Value), getEmployerIDByCustomerID(tbcustomerid.Text.Trim)), batchname)
                                ElseIf checkIfCustomerGroupIsForMerging(tbcustomerid.Text.Trim, conFPM) = False Then
                                    saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(3).Value), Decimal.Parse(row.Cells(4).Value), row.Cells(1).Value, batchname)
                                End If
                            Else
                                saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(3).Value), Decimal.Parse(row.Cells(4).Value), row.Cells(1).Value, batchname)
                            End If
                        End If
                    Next
                    loadProductSummary(batchname)
                    deleteGroupingBreakdown(batchname)

                End If
            ElseIf go = False Then
                MessageBox.Show("Please ensure that you have put the Receipt Type", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Cannot proceed. Amount should be greater than 0", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub cbmodename_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbmodename.SelectedIndexChanged
        tbmodename.Text = cbmodename.Text
        tbmodeid.Text = cbmodename.SelectedValue.ToString
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If decision(" Are you sure they are the same ?") = DialogResult.Yes Then
                tbreceivedfrom.Text = tbpayername.Text
                tbreceivedidnumber.Text = tbpayerid.Text
                tbreceivedcellnumber.Text = tbpayercellnumber.Text
                CheckBox1.Checked = True
            Else
                tbreceivedfrom.Text = ""
                tbreceivedidnumber.Text = ""
                tbreceivedcellnumber.Text = ""
                CheckBox1.Checked = False
            End If
        ElseIf CheckBox1.Checked = False Then
            tbreceivedfrom.Text = ""
            tbreceivedidnumber.Text = ""
            tbreceivedcellnumber.Text = ""
            CheckBox1.Checked = False
        End If
    End Sub

    Public Sub loadReport()
        Dim dtcrystalreport As New DataTable

        ' start report
        dtcrystalreport.Columns.Add("PostalAddress", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PhysicalAddress", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ContactAddresses", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromIDNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceivedFromCellNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerIDNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PayerCellNumber", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("Description", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("PaymentMode", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("Amount", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceiptDate", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ReceiptNo", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ProductName", Type.GetType("System.String"))
        dtcrystalreport.Columns.Add("ProductAmount", Type.GetType("System.String"))

        PostalAddress = getPostalAddress()
        PhysicalAddress = getPhysicalAddress()
        ContactAddresses = "Tel = " & getCompanyContactNumber() & " : Fax = " & getCompanyFaxNumber() & " : Email = " & getCompanyEmailAdress()
        ReceivedFromName = tbreceivedfrom.Text.Trim
        ReceivedFromIDNumber = tbreceivedidnumber.Text.Trim
        ReceivedFromCellNumber = tbreceivedcellnumber.Text.Trim
        PayerName = tbpayername.Text.Trim
        PayerIDNumber = tbpayerid.Text.Trim
        PayerCellNumber = tbpayercellnumber.Text.Trim
        Description = tbdescription.Text.Trim
        PaymentMode = tbmodename.Text.Trim
        Amount = getCurrency(Decimal.Parse(tbamount.Text.Trim))
        ReceiptDate = Date.Now.Date
        ReceiptNo = tbreference.Text.Trim

        For Each row As DataGridViewRow In DataGridView2.Rows
            If Decimal.Parse(row.Cells(1).Value) > 0 Then
                PayerName = tbpayername.Text.Trim
                PayerIDNumber = tbpayerid.Text.Trim
                ProductsName = row.Cells(0).Value
                ProductAmount = getCurrency(Decimal.Parse(row.Cells(1).Value))
                dtcrystalreport.Rows.Add(PostalAddress, PhysicalAddress, ContactAddresses, ReceivedFromName, ReceivedFromIDNumber, ReceivedFromCellNumber, PayerName, PayerIDNumber, PayerCellNumber, Description, PaymentMode, Amount, ReceiptDate, ReceiptNo, ProductsName, ProductAmount)
            End If
        Next

        Dim objRpt As New ReceiptReport
        objRpt.SetDataSource(dtcrystalreport)
        Me.CrystalReportViewer1.ReportSource = objRpt
        Me.CrystalReportViewer1.DisplayBackgroundEdge = True
        Me.CrystalReportViewer1.DisplayToolbar = True
        Me.CrystalReportViewer1.DisplayGroupTree = True
        Me.CrystalReportViewer1.Refresh()
        Me.Cursor = Windows.Forms.Cursors.Default
        'Me.TabControl1.SelectTab(1)
        crrpl = objRpt

    End Sub

    Private Sub tbreceivedcellnumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbreceivedcellnumber.TextChanged
        If CheckBox1.Checked = True Then
            tbpayercellnumber.Text = tbreceivedcellnumber.Text.Trim
        End If
    End Sub

    Private Sub tbpayercellnumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbpayercellnumber.TextChanged
        If CheckBox1.Checked = True Then
            tbreceivedcellnumber.Text = tbpayercellnumber.Text.Trim
        End If
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Try
            For Each ctrl As Control In CrystalReportViewer1.Controls
                If TypeOf ctrl Is Windows.Forms.ToolStrip Then
                    Dim btnNew As New ToolStripButton
                    btnNew.Text = "Print Report"
                    ''btnNew.Image = Image.FromFile("../../Resources/printer.ico")
                    CType(ctrl, ToolStrip).Items.Add(btnNew)
                    AddHandler btnNew.Click, AddressOf Printlaf
                    'index = 6
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Printlaf(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            crrpl.PrintToPrinter(1, False, 0, 0)
            MessageBox.Show("Done printing", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.TabControl1.SelectTab(0)
    End Sub

    Private Sub tbmodeid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbmodeid.TextChanged
        Try
            If tbmodeid.Text.Trim = "" Then

            Else
                If go = True Then
                    If checkIfToModeIsDirectDeposit(Integer.Parse(tbmodeid.Text.Trim)) = True Then
                        tbpayercellnumber.Text = "N/A"
                        tbreceivedfrom.Text = tbpayername.Text
                        tbreceivedidnumber.Text = tbpayerid.Text
                        tbreceivedcellnumber.Text = tbpayercellnumber.Text
                        tbdescription.Text = "N/A"
                        tbreference.Text = ""
                        tbreference.Enabled = True
                    Else
                        tbreference.Text = getNextReceipNo()
                        tbreference.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub loadProductSummary(ByVal BatchName As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select GroupingName,SUM(Instalment) as Receipt from GroupingBreakdown where BatchName = '" & BatchName & "'  group by GroupingName"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "DIM")
            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView2.DataSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadProductSummary()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProductSummary()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

End Class