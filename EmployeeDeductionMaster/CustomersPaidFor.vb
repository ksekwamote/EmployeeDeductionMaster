Imports DgvFilterPopup

Public Class CustomersPaidFor

    Dim go As Boolean

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vrl As New CustomerList
        vrl.MdiParent = mform
        vrl.whichform = "CustomersPaidFor"
        vrl.cpf = Me
        vrl.Show()
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
    End Sub

    Private Sub loadProducts(ByVal customerid As String, ByVal payerid As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,Products,'' as Deduction from Products where Status = 'Active'"

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
            tbtotal.Text = "0"
            If ds.Tables(0).Rows.Count > 0 Then
                For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    ds.Tables(0).Rows(count).Item(2) = getPaidForAmountPerProductID(customerid, Integer.Parse(ds.Tables(0).Rows(count).Item(0).ToString), payerid, employerid)
                    tbtotal.Text = Decimal.Parse(tbtotal.Text.Trim) + Decimal.Parse(ds.Tables(0).Rows(count).Item(2).ToString)
                Next
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                filterManager = New DgvFilterManager(DataGridView1)
                For Each col As DataGridViewColumn In DataGridView1.Columns
                    If col.Index = 2 Then
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
            MessageBox.Show(ex.Message, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcustomername.Text.Trim = "" Or tbpayername.Text.Trim = "" Then
            MessageBox.Show("Select the customer and payer first", "Meesage()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            loadProducts(tbcustomerid.Text.Trim, tbpayerid.Text.Trim)
            Button4.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                go = True
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Decimal.Parse(row.Cells(2).Value) < 0 Then
                        go = False
                    End If
                Next
                If go = True Then
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        If Decimal.Parse(row.Cells(2).Value) > 0 Then
                            If checkProductForCustomerPaidFor(tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value), tbpayerid.Text.Trim, employerid) = True Then
                                'MessageBox.Show("MM22222", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                updatePaidForCustomer(Decimal.Parse(row.Cells(2).Value), "Active", tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value), tbpayerid.Text.Trim, employerid)
                            Else
                                'MessageBox.Show("MM11111", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                savePaidForCustomer(tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value), Decimal.Parse(row.Cells(2).Value), "Active", tbpayerid.Text.Trim, employerid, con)
                            End If
                        ElseIf Decimal.Parse(row.Cells(2).Value) = 0 Then
                            'MessageBox.Show("MM33333", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            updatePaidForCustomerStatus("Active", tbcustomerid.Text.Trim, Integer.Parse(row.Cells(0).Value), tbpayerid.Text.Trim, employerid)
                        End If
                    Next
                End If
                MessageBox.Show("Done Saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button4.Enabled = True
        tbtotal.Text = 0.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            tbtotal.Text = Decimal.Parse(tbtotal.Text.Trim) + Decimal.Parse(row.Cells(2).Value)
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim vrl As New CustomerList
        vrl.MdiParent = mform
        vrl.whichform = "CustomersPaidForPayer"
        vrl.cpf = Me
        vrl.Show()
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
    End Sub

    Private Sub CustomersPaidFor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class