Public Class AddCustomer

    Dim addcustomer As Boolean
    Dim PreEmployerName As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        tbcustomername.Text = getCustomerName(tbcustomerid.Text.Trim)
        If tbcustomername.Text.Trim = "" Then
            tbcustomerid.Enabled = True
            tbcustomername.Enabled = True
            AddCustomer = False
            If checkChequeCustomerEmployer() = True Then

            Else
                saveEmployers("Cheque Customers", 0.0, 0, "No", "", "", "", "")
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
                If decision("Are you sure the ID Number is Correct ?") = DialogResult.Yes Then
                    tbcustomerid.Enabled = False
                    tbcustomername.Enabled = True
                    addcustomer = True
                    Button3.Enabled = True
                Else
                    tbcustomerid.Enabled = False
                    'tbcustomername.Enabled = False
                    addcustomer = False
                    Button3.Enabled = False
                End If
            Else
                tbcustomerid.Enabled = False
                'tbcustomername.Enabled = False
                AddCustomer = True
                Button3.Enabled = False
            End If
        Else
            tbcustomerid.Enabled = False
            'tbcustomername.Enabled = False
            tbemployername.Text = getEmployerNameByCustomerID(tbcustomerid.Text.Trim)
            tbemployerid.Text = getEmployerIDByCustomerID(tbcustomerid.Text.Trim)
            PreEmployerName = tbemployername.Text.Trim

            tbgender.Text = getGender(tbcustomerid.Text.Trim)
            addcustomer = False
            Button3.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If AddCustomer = True Then
            If tbcustomername.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbgender.Text.Trim = "" Then
                MessageBox.Show("Enter Customer Name", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure save customer details ?") = DialogResult.Yes Then
                    saveCustomerDetails(tbcustomerid.Text.Trim, tbcustomername.Text.Trim, Integer.Parse(tbemployerid.Text.Trim), tbgender.Text.Trim)
                    wrapper.saveAction("Customer Added", Date.Now, logedinname, tbcustomerid.Text.Trim)
                    tbcustomerid.Enabled = False
                    tbcustomername.Enabled = False
                    addcustomer = False
                    Button3.Enabled = False
                    MessageBox.Show("Done Saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Close()
                End If
            End If
        End If
    End Sub

    Private Sub AddCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If getValueBooleanOtherSystems("select * from CompanyDetails where Integrate = 'Yes'", con) = True Then
            Me.Enabled = False
        Else
            loadEmployers()
            tbemployername.Text = ""
            tbemployerid.Text = ""
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

    Private Sub cbgender_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbgender.SelectedIndexChanged
        tbgender.Text = cbgender.Text
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcustomername.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbgender.Text.Trim = "" Then
            MessageBox.Show("Enter Customer Name", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure update changes ?") = DialogResult.Yes Then
                updateCustomerDetails(tbcustomerid.Text.Trim, tbcustomername.Text.Trim, Integer.Parse(tbemployerid.Text.Trim), tbgender.Text.Trim)
                wrapper.saveAction("Customer Updated", Date.Now, logedinname, tbcustomerid.Text.Trim)
                If PreEmployerName = tbemployername.Text.Trim Then

                Else
                    wrapper.saveAction("Employer Changed from " & PreEmployerName & " to " & tbemployername.Text.Trim, Date.Now, logedinname, tbcustomerid.Text.Trim)
                End If
                tbcustomerid.Enabled = False
                tbcustomername.Enabled = False
                addcustomer = False
                Button3.Enabled = False
                Button2.Enabled = False
                MessageBox.Show("Changes updated", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
            End If
    End Sub
End Class