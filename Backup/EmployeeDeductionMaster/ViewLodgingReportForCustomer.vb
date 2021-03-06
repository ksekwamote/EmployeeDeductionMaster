Public Class ViewLodgingReportForCustomer

    Dim ds As New DataSet
    Dim total As Decimal = 0
    Dim commissionproductid As Integer = 0
    Dim totalamounttocalculatecommission, commissionamount As Decimal
    Dim dscommproducts As DataSet

    Private Sub loadCustomerIDs(ByVal customerid As String)
        Dim qry As String = "select DISTINCT(CustomerID) from CustomerDetails where CustomerID = '" & customerid & "'"

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
        updateValues()
    End Sub

    Public Sub updateValues()
        If TextBox1.Text.Trim = "" Then
            MessageBox.Show("Enter CustomerID", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Button1.Enabled = False
            loadCustomerIDs(TextBox1.Text.Trim)
            ds = getProducts()
            If checkCustomerDetailsInDatabase(TextBox1.Text.Trim) = True Then
                dtpformonth.Value = getLastDateOfMonth(getLastLodgingDate(getEmployerIDByCustomerID(TextBox1.Text.Trim)))
                dtppredate.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))

                Label5.Text = getEmployerNameByCustomerID(TextBox1.Text.Trim)
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

                Dim col3 As New DataGridViewTextBoxColumn()
                col3.HeaderText = "CustomerName"
                col3.Name = 0
                DataGridView1.Columns.Add(col3)

                For Each col As DataGridViewColumn In DataGridView1.Columns
                    DataGridView1.Rows(1).Cells(col.Index).Style.BackColor = Color.Gray
                Next

                ProgressBar1.Maximum = DataGridView1.Rows.Count
                commissionproductid = getProductID(getEmployerIDByCustomerID(TextBox1.Text.Trim))
                For Each row As DataGridViewRow In DataGridView1.Rows

                    '' Update Commission
                    If getCommissionPercent(getEmployerIDByCustomerID(TextBox1.Text.Trim)) > 0 Then
                        'MessageBox.Show("Here 1", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        totalamounttocalculatecommission = 0
                        dscommproducts = getCommissionProdcuts(getEmployerIDByCustomerID(TextBox1.Text.Trim))
                        For count As Integer = 0 To dscommproducts.Tables(0).Rows.Count - 1
                            totalamounttocalculatecommission = totalamounttocalculatecommission + getExpectedPerProduct(dtpformonth.Value, Integer.Parse(dscommproducts.Tables(0).Rows(count).Item(3).ToString()), row.Cells(0).Value.ToString, dtppredate.Value, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                        Next
                        'MessageBox.Show(totalamounttocalculatecommission, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        commissionamount = ((getCommissionPercent(getEmployerIDByCustomerID(TextBox1.Text.Trim)) / 100) * totalamounttocalculatecommission)
                        'MessageBox.Show(commissionamount, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If checkAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, getEmployerIDByCustomerID(TextBox1.Text.Trim)) = True Then
                            updateAmendment(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, commissionamount, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                        ElseIf checkTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, getEmployerIDByCustomerID(TextBox1.Text.Trim)) = True Then
                            updateTermination(row.Cells(0).Value.ToString, getLastDateOfMonth(dtpformonth.Value), commissionproductid, commissionamount, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                        Else
                            saveAmendments(row.Cells(0).Value.ToString, commissionamount, getLastDateOfMonth(dtpformonth.Value), Date.Now, commissionproductid, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                        End If
                    End If
                    ''end of Update Commission

                    total = 0
                    For Each col As DataGridViewColumn In DataGridView1.Columns
                        If col.Index = 0 Or col.Index = 1 Then

                        ElseIf col.Index = DataGridView1.Columns.Count - 1 Then
                            row.Cells(col.Index).Value = getCustomerName(row.Cells(0).Value)
                        Else
                            If row.Index = 0 Then
                                row.Cells(col.Index).Value = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtppredate.Value, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                                total = total + Decimal.Parse(row.Cells(col.Index).Value)
                            ElseIf row.Index = 1 Then
                                row.Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value
                                row.Cells(col.Index).Value = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtppredate.Value, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                                total = total + Decimal.Parse(row.Cells(col.Index).Value)
                            End If
                        End If
                    Next
                    row.Cells(1).Value = total
                    ProgressBar1.Value = row.Index
                Next
                ProgressBar1.Value = 0
                Button2.Enabled = True
            Else
                MessageBox.Show("There customer does not exist in the database", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
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
        total = 0
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If col.Index = 0 Or col.Index = 1 Then

            ElseIf col.Index = DataGridView1.Columns.Count - 1 Then

            Else
                total = total + Decimal.Parse(DataGridView1.Rows(0).Cells(col.Index).Value)
            End If
        Next
        DataGridView1.Rows(0).Cells(1).Value = total
    End Sub

    Private Sub ViewLodgingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Enabled = True
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim vrl As New ViewLodgingReportForCustomer
        vrl.MdiParent = mform
        vrl.Show()
        Me.Close()
    End Sub

End Class