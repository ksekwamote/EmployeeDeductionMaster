Public Class ViewAmendments

    Dim ds As New DataSet
    Dim total As Decimal = 0
    Dim commissionproductid As Integer = 0
    Dim totalamounttocalculatecommission, commissionamount As Decimal
    Dim dscommproducts As DataSet

    Private Sub loadCustomerIDs(ByVal sdate As Date, ByVal employerid As Integer)
        Dim qry As String = ""
        qry = "select DISTINCT(CustomerID )from TerminationsAndAmendments where ForMonth = '" & sdate & "' and CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "') and ProductID not in (select AutoID from Products where AutoAllocate = 1)"

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
        
        loadCustomerIDs(dtpformonth.Value, employerid)
        ds = getProducts(employerid)
        Dim col1 As New DataGridViewTextBoxColumn()
        col1.HeaderText = "Total"
        col1.Name = "Total"
        DataGridView1.Columns.Add(col1)

        For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = ds.Tables(0).Rows(index).Item(0).ToString()
            col.Name = ds.Tables(0).Rows(index).Item(1).ToString()
            DataGridView1.Columns.Add(col)
        Next

        Dim col2 As New DataGridViewTextBoxColumn()
        col2.HeaderText = "Status"
        col2.Name = 0
        DataGridView1.Columns.Add(col2)

        Dim col3 As New DataGridViewTextBoxColumn()
        col3.HeaderText = "CustomerName"
        col3.Name = 0
        DataGridView1.Columns.Add(col3)

        ProgressBar1.Maximum = DataGridView1.Rows.Count
        commissionproductid = getProductID(employerid)
        For Each row As DataGridViewRow In DataGridView1.Rows

            '' Update Commission
            If getCommissionPercent(employerid) > 0 Then
                'MessageBox.Show("Here 1", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                totalamounttocalculatecommission = 0
                dscommproducts = getCommissionProdcuts(employerid)
                For count As Integer = 0 To dscommproducts.Tables(0).Rows.Count - 1
                    totalamounttocalculatecommission = totalamounttocalculatecommission + getExpectedPerProduct(dtpformonth.Value, Integer.Parse(dscommproducts.Tables(0).Rows(count).Item(3).ToString()), row.Cells(0).Value.ToString, dtpprevious.Value, employerid)
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

                ElseIf col.Index = DataGridView1.Columns.Count - 1 Then
                    row.Cells(col.Index).Value = getCustomerName(row.Cells(0).Value.ToString)
                ElseIf (col.Index - (DataGridView1.Columns.Count - 1)) = -1 Then
                    If checkAmendmentsAndTermiantionsForCustomer(row.Cells(0).Value.ToString, dtpformonth.Value, employerid) = True Then
                        row.Cells(col.Index).Value = "New"
                    Else
                        row.Cells(col.Index).Value = "Old"
                    End If
                Else
                    row.Cells(col.Index).Value = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtpprevious.Value, employerid)
                    total = total + Decimal.Parse(row.Cells(col.Index).Value)
                End If
            Next
            row.Cells(1).Value = total
            ProgressBar1.Value = row.Index
        Next
        ProgressBar1.Value = 0

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpformonth.ValueChanged
        dtpformonth.Value = getLastDateOfMonth(dtpformonth.Value)
        dtpprevious.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))
        Button1.Enabled = True
    End Sub

    Function getProducts(ByVal EmployerID As Integer) As DataSet
        ''Dim qry As String = "select Products,AutoID from Products Where Status = 'Active' order by AllocationOrder ASC"
        Dim qry As String = "select Products,EmployerProducts.ProductID from EmployerProducts inner join Products on EmployerProducts.ProductID = Products.AutoID where EmployerProducts.Status = 'Active' and Products.Status = 'Active' and EmployerID = '" & EmployerID & "' order by OrderNumber ASC"

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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, employername & " " & dtpformonth.Value)
        MessageBox.Show("Done Exporting", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ViewAmendments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtpformonth.Value = getLastDateOfMonth(addMonth(getLastLodgingDate(employerid), 1))
        dtpprevious.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))
        Button1.Enabled = True
    End Sub
End Class