Public Class Add_Requested_Refund

    Dim addcustomer As Boolean

    Private Sub Add_Requested_Refund_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadProductsWithNoSystem()
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""
    End Sub

    Private Sub loadProductsWithNoSystem()
        DataGridView1.DataSource = Nothing

        Dim qry As String = "select AutoID,Products from Products where HasSystem = 0"
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
                DataGridView1.DataSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadProductsWithNoSystem()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbproductid.Text.Trim = "" Or tbrefundamount.Text.Trim = "" Or tbpayby.Text.Trim = "" Or tbbankname.Text.Trim = "" Or tbbranchcode.Text.Trim = "" Or tbaccountnumber.Text.Trim = "" Or tbpayeename.Text.Trim = "" Then
            MessageBox.Show("Enter all details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure to continue ?") = DialogResult.Yes Then
                If checkPendingCheques(DateTimePicker1.Value, tbcustomerid.Text.Trim) = False Then
                    saveRefund(tbcustomerid.Text.Trim, Integer.Parse(tbproductid.Text.Trim), Decimal.Parse(tbrefundamount.Text.Trim), DateTimePicker1.Value, logedinname, Date.Now, "", "New", tbpayby.Text.Trim, tbbankname.Text.Trim, tbbranchcode.Text.Trim, tbaccountnumber.Text.Trim, tbpayeename.Text.Trim)
                    wrapper.saveAction("Refund Saved", Date.Now, logedinname, tbcustomerid.Text.Trim)
                    MessageBox.Show("Refund Saved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Close()
                Else
                    MessageBox.Show("Approve/ Write pending cheques first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'If decision("Do you want to request the cheque ?") = DialogResult.Yes Then
                    '    saveRefund(tbcustomerid.Text.Trim, Integer.Parse(tbproductid.Text.Trim), Decimal.Parse(tbrefundamount.Text.Trim), DateTimePicker1.Value, logedinname, Date.Now, "", "New")
                    '    wrapper.saveAction("Refund Saved", Date.Now, logedinname, tbcustomerid.Text.Trim)
                    '    MessageBox.Show("Refund Saved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    Close()
                    'End If
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        tbcustomername.Text = getCustomerName(tbcustomerid.Text.Trim)
        If tbcustomername.Text.Trim = "" Then
            GroupBox1.Enabled = False
            tbcustomerid.Enabled = True
            tbcustomername.Enabled = True
            addcustomer = False
            If checkChequeCustomerEmployer() = True Then

            Else
                saveEmployers("Cheque Customers", 0.0, 0, "No")
            End If

            tbcustomername.Text = getFPMCustomerName(tbcustomername.Text.Trim, conFPM)
            tbgender.Text = getFPMCustomerGender(tbcustomername.Text.Trim, conFPM)
            If tbcustomername.Text.Trim = "" Then
                tbcustomername.Text = getMLMCustomerName(tbcustomername.Text.Trim, conMLMAdvances)
                tbgender.Text = getMLMCustomerGender(tbcustomername.Text.Trim, conMLMAdvances)
                If tbcustomername.Text.Trim = "" Then
                    tbcustomername.Text = getRMCustomerName(tbcustomername.Text.Trim, conRMOrange)
                    tbgender.Text = getRMCustomerGender(tbcustomername.Text.Trim, conRMOrange)
                    If tbcustomername.Text.Trim = "" Then
                        tbcustomername.Text = getRMCustomerName(tbcustomername.Text.Trim, conRMBeMobile)
                        tbgender.Text = getRMCustomerGender(tbcustomername.Text.Trim, conRMBeMobile)
                    End If
                End If
            End If

            If tbcustomername.Text.Trim = "" Then
                MessageBox.Show("Customer does not exists in all the systems", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If decision("Are you sure the IDNumber is Correct ?") = DialogResult.Yes Then
                    GroupBox1.Enabled = True
                    tbcustomerid.Enabled = False
                    'tbcustomername.Enabled = True
                    addcustomer = True
                    Button3.Enabled = True
                Else
                    GroupBox1.Enabled = False
                    GroupBox2.Enabled = False
                    tbcustomerid.Enabled = False
                    'tbcustomername.Enabled = False
                    addcustomer = False
                    Button3.Enabled = False
                End If
            Else
                GroupBox1.Enabled = True
                tbcustomerid.Enabled = False
                'tbcustomername.Enabled = False
                addcustomer = True
                Button3.Enabled = False
            End If
        Else
            tbemployername.Text = getEmployerNameByCustomerID(tbcustomerid.Text.Trim)
            tbemployerid.Text = getEmployerIDByCustomerID(tbcustomerid.Text.Trim)
            tbgender.Text = getGender(tbcustomerid.Text.Trim)
            GroupBox1.Enabled = True
            tbcustomerid.Enabled = False
            'tbcustomername.Enabled = False
            addcustomer = False
            Button3.Enabled = False
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        GroupBox2.Enabled = True
        GroupBox1.Enabled = False
        tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
        tbproductname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If addcustomer = True Then
            If tbcustomername.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbgender.Text.Trim = "" Then
                MessageBox.Show("Enter Customer Name", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                saveCustomerDetails(tbcustomerid.Text.Trim, tbcustomername.Text.Trim, Integer.Parse(tbemployerid.Text.Trim), tbgender.Text.Trim)
                wrapper.saveAction("Customer Added", Date.Now, logedinname, tbcustomerid.Text.Trim)
                GroupBox1.Enabled = True
                tbcustomerid.Enabled = False
                tbcustomername.Enabled = False
                addcustomer = False
                Button3.Enabled = False
            End If
        End If
    End Sub

    Public Sub loadEmployers()
        Dim qry As String = "select * from Employers"

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
                Me.cbemployer.DisplayMember = "Emp.EmployerName"
                Me.cbemployer.ValueMember = "Emp.AutoID"
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

    Private Sub cbemployer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbemployer.SelectedIndexChanged
        tbemployername.Text = cbemployer.Text
        tbemployerid.Text = cbemployer.SelectedValue.ToString
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbpayby.Text = ComboBox1.Text
    End Sub

    Private Sub cbgender_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbgender.SelectedIndexChanged
        tbgender.Text = cbgender.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If tbcustomername.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbgender.Text.Trim = "" Then
            MessageBox.Show("Enter Customer Name", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateCustomerDetails(tbcustomerid.Text.Trim, tbcustomername.Text.Trim, Integer.Parse(tbemployerid.Text.Trim), tbgender.Text.Trim)
            wrapper.saveAction("Customer Updated", Date.Now, logedinname, tbcustomerid.Text.Trim)
            tbcustomerid.Enabled = False
            tbcustomername.Enabled = False
            addcustomer = False
            Button3.Enabled = False
            Button2.Enabled = False
            Close()
        End If
    End Sub
End Class