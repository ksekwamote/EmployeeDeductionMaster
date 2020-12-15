Public Class UpdateLodgingReport

    Dim ds As New DataSet
    Dim total As Decimal = 0
    Dim commissionproductid As Integer = 0
    Dim totalamounttocalculatecommission, commissionamount As Decimal
    Dim dscommproducts As DataSet

    Public fromrefundlist As Boolean = False
    Public prts As New PotentialRefundsToStop
    Public index As Integer = 0
    Public refundamount As Decimal
    Public refundid As String
    Public depositdate As Date

    Dim goahead As Boolean = True

    Private Sub loadCustomerIDs(ByVal customerid As String, ByVal employerid As Integer)
        Dim qry As String = "select DISTINCT(CustomerID) from CustomerDetails where CustomerID = '" & customerid & "' and EmployerID = '" & employerid & "'"

        'Dim qry As String = "select DISTINCT(CustomerID) from ExpectedDeductions where CustomerID = '738521706'"
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                dt.Rows.Add()
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs

                DataGridView1.Columns(0).ReadOnly = True
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing

            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCustomerIDs()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomerIDs()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If cbchangelodgingmonth.Checked = True Then
            If checkExpectedDeductions(dtpformonth.Value, employerid) = True Then
                MessageBox.Show("There are expected deductions for this month. The File is CLosed", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                updateValues()
            End If
        Else
            If checkExpectedDeductions(dtpformonth.Value, employerid) = True Then
                MessageBox.Show("There are expected deductions for this month. The File is CLosed", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                updateValues()
            End If
        End If
    End Sub

    Public Sub updateValues()
        If TextBox1.Text.Trim = "" Then
            MessageBox.Show("Enter CustomerID", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Button1.Enabled = False
            loadCustomerIDs(TextBox1.Text.Trim, employerid)
            ds = getProducts()
            If checkCustomerDetailsInDatabase(TextBox1.Text.Trim) = True Then
                If checkCustomerDetails(TextBox1.Text.Trim, employerid) = True Then
                    Dim col1 As New DataGridViewTextBoxColumn()
                    col1.HeaderText = "Total"
                    col1.Name = "Total"
                    DataGridView1.Columns.Add(col1)
                    DataGridView1.Columns(1).ReadOnly = True
                    'DataGridView1.Rows.Add()
                    DataGridView1.Rows(1).ReadOnly = True

                    For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
                        Dim col As New DataGridViewTextBoxColumn()
                        col.HeaderText = ds.Tables(0).Rows(index).Item(0).ToString()
                        col.Name = ds.Tables(0).Rows(index).Item(1).ToString()
                        DataGridView1.Columns.Add(col)
                    Next

                    For Each col As DataGridViewColumn In DataGridView1.Columns
                        DataGridView1.Rows(1).Cells(col.Index).Style.BackColor = Color.Gray
                    Next

                    ProgressBar1.Maximum = DataGridView1.Rows.Count
                    commissionproductid = getProductID(employerid)
                    For Each row As DataGridViewRow In DataGridView1.Rows

                        '' Update Commission
                        If getCommissionPercent(employerid) > 0 Then
                            'MessageBox.Show("Here 1", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            totalamounttocalculatecommission = 0
                            dscommproducts = getCommissionProdcuts(employerid)
                            For count As Integer = 0 To dscommproducts.Tables(0).Rows.Count - 1
                                totalamounttocalculatecommission = totalamounttocalculatecommission + getExpectedPerProduct(dtpformonth.Value, Integer.Parse(dscommproducts.Tables(0).Rows(count).Item(3).ToString()), row.Cells(0).Value.ToString, dtppredate.Value, employerid)
                            Next
                            'MessageBox.Show(totalamounttocalculatecommission, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            commissionamount = ((getCommissionPercent(employerid) / 100) * totalamounttocalculatecommission)
                            'MessageBox.Show(commissionamount, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            If checkAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, employerid) = True Then
                                updateAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, commissionamount, employerid)
                            ElseIf checkTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, employerid) = True Then
                                updateTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, commissionamount, employerid)
                            Else
                                saveAmendments(row.Cells(0).Value.ToString, commissionamount, getLastDateOfMonth(dtpformonth.Value), Date.Now, commissionproductid, employerid)
                            End If
                        End If
                        ''end of Update Commission

                        total = 0
                        For Each col As DataGridViewColumn In DataGridView1.Columns
                            If col.Index = 0 Or col.Index = 1 Then

                            Else
                                If row.Index = 0 Then
                                    row.Cells(col.Index).Value = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtppredate.Value, employerid)
                                    total = total + Decimal.Parse(row.Cells(col.Index).Value)
                                ElseIf row.Index = 1 Then
                                    row.Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value
                                    row.Cells(col.Index).Value = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtppredate.Value, employerid)
                                    total = total + Decimal.Parse(row.Cells(col.Index).Value)
                                End If
                            End If
                        Next
                        row.Cells(1).Value = total
                        ProgressBar1.Value = row.Index
                    Next
                    ProgressBar1.Value = 0
                    Button2.Enabled = True

                    If fromrefundlist = True Then
                        Button4.Enabled = True
                    Else
                        Button4.Enabled = False
                    End If

                Else
                    MessageBox.Show("Customer exists in " & getEmployerNameByCustomerID(TextBox1.Text.Trim), "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("There customer does not exist in the database", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpformonth.ValueChanged
        dtpformonth.Value = getLastDateOfMonth(dtpformonth.Value)
        dtppredate.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))
        Button1.Enabled = True
    End Sub

    Function getProducts() As DataSet
        Dim qry As String = "select Products,AutoID from Products Where Status = 'Active' order by AllocationOrder ASC"

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button3.Enabled = True
        total = 0
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If col.Index = 0 Or col.Index = 1 Then

            Else
                total = total + Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value)
            End If
        Next
        DataGridView1.Rows(0).Cells(1).Value = total
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = False
        Button1.Enabled = False

        goahead = True

        If fromrefundlist = True Then
            goahead = True
        Else
            For Each col As DataGridViewColumn In DataGridView1.Columns
                If col.Index = 0 Or col.Index = 1 Then

                Else
                    If Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value) = Decimal.Parse(DataGridView1.Rows(1).Cells(col.Index).Value) Then

                    Else
                        goahead = True
                    End If
                End If
            Next
        End If

        If goahead = True Then
            Button3.Enabled = False
            For Each col As DataGridViewColumn In DataGridView1.Columns
                If col.Index = 0 Or col.Index = 1 Then

                Else
                    If Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value) = Decimal.Parse(DataGridView1.Rows(1).Cells(col.Index).Value) Then

                    Else
                        If checkAmendment(DataGridView1.Rows(0).Cells(0).Value, dtpformonth.Value, Integer.Parse(col.Name), employerid) = True Then
                            updateAmendment(DataGridView1.Rows(0).Cells(0).Value, dtpformonth.Value, Integer.Parse(col.Name), Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value), employerid)
                        ElseIf checkTermination(DataGridView1.Rows(0).Cells(0).Value, dtpformonth.Value, Integer.Parse(col.Name), employerid) = True Then
                            updateTermination(DataGridView1.Rows(0).Cells(0).Value, dtpformonth.Value, Integer.Parse(col.Name), Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value), employerid)
                        Else
                            saveAmendments(DataGridView1.Rows(0).Cells(0).Value, Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value), dtpformonth.Value, Date.Now, Integer.Parse(col.Name), employerid)
                        End If
                    End If
                End If
            Next
            MessageBox.Show("Done updating changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            wrapper.saveAction("Updated Lodging File For Customer", Date.Now, logedinname, DataGridView1.Rows(0).Cells(0).Value)

            If fromrefundlist = True Then
                If decision("Remove refund from the list ?") = DialogResult.Yes Then
                    prts.DataGridView1.Rows.RemoveAt(index)
                    saveStoppedRefunds(refundid, refundamount, depositdate, "Stopped", Label6.Text, TextBox1.Text.Trim)
                'prts.Label2.Text = Integer.Parse(Label2.Text) - 1
                    Me.Close()
                End If
            Else
                Dim vrl As New UpdateLodgingReport
                vrl.MdiParent = mform
                vrl.Text = vrl.Text & " : Employer Name = " & employername
                vrl.TextBox1.Text = TextBox1.Text.Trim
                vrl.Show()
                vrl.updateValues()
                Me.Close()
            End If
        Else
            MessageBox.Show("There are no changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub UpdateLodgingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtpformonth.Value = getLastDateOfMonth(addMonth(getLastLodgingDate(employerid), 1))
        dtppredate.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))
        Button1.Enabled = True
    End Sub

    Private Sub cbchangelodgingmonth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbchangelodgingmonth.CheckedChanged
        If cbchangelodgingmonth.Checked = True Then
            dtpformonth.Enabled = True
        Else
            dtpformonth.Enabled = False
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim vrl As New UpdateLodgingReport
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
        vrl.TextBox1.Text = TextBox1.Text
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If fromrefundlist = True Then
            If decision("Refund already stopped ?") = DialogResult.Yes Then
                prts.DataGridView1.Rows.RemoveAt(index)
                If Decimal.Parse(DataGridView1.Rows(0).Cells(1).Value) = Decimal.Parse(DataGridView1.Rows(1).Cells(1).Value) Then
                    saveStoppedRefunds(refundid, refundamount, depositdate, "Stopped", Label6.Text, TextBox1.Text.Trim)
                End If
                Me.Close()
            End If
        End If
    End Sub
End Class