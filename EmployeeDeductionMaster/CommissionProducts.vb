Public Class CommissionProducts

    Private Sub CommissionProducts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCommissionProducts()
        loadEmployers()
        loadProducts()
        tbemployerid.Text = ""
        tbproductid.Text = ""
        tbstatus.Text = ""
        tbemployername.Text = ""
        tbproductname.Text = ""
    End Sub

    Private Sub loadCommissionProducts()
        DataGridView1.DataSource = Nothing

        Dim qry As String = "select * from CommissionProducts"
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
            MessageBox.Show(ex.Message, "loadCommissionProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Public Sub loadProducts()
        Dim qry As String = "select * from Products where AutoAllocate not in (1)"

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
                Me.cbproduct.DataSource = ds
                Me.cbproduct.DisplayMember = "Emp.Products"
                Me.cbproduct.ValueMember = "Emp.AutoID"
                Me.cbproduct.SelectedIndex = 0
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


    Public Sub saveCommissionProduct(ByVal EmployerID As Integer, ByVal EmployerName As String, ByVal ProductID As Integer, ByVal ProductName As String, ByVal Status As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into CommissionProducts values ('" & EmployerID & "','" & EmployerName & "','" & ProductID & "','" & ProductName & "','" & Status & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details saved", "saveEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveCommissionProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveCommissionProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateCommissionProduct(ByVal EmployerID As Integer, ByVal EmployerName As String, ByVal ProductID As Integer, ByVal ProductName As String, ByVal Status As Integer, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CommissionProducts set EmployerID = '" & EmployerID & "',EmployerName = '" & EmployerName & "',ProductID = '" & ProductID & "',ProductName = '" & ProductName & "',Status = '" & Status & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details updated", "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateCommissionProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateCommissionProduct()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try

    End Sub

    Private Sub cbemployer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbemployer.SelectedIndexChanged
        tbemployerid.Text = cbemployer.SelectedValue.ToString
        tbemployername.Text = cbemployer.Text
    End Sub

    Private Sub cbproduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbproduct.SelectedIndexChanged
        tbproductid.Text = cbproduct.SelectedValue.ToString
        tbproductname.Text = cbproduct.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbproductid.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbstatus.Text.Trim = "" Then
            MessageBox.Show("Make Sure to all the details are entered", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            saveCommissionProduct(Integer.Parse(tbemployerid.Text.Trim), tbemployername.Text, Integer.Parse(tbproductid.Text.Trim), tbproductname.Text, Integer.Parse(tbstatus.Text.Trim))
            loadCommissionProducts()
            Button2.Enabled = False
            Button1.Enabled = True
            tbemployerid.Text = ""
            tbproductid.Text = ""
            tbemployername.Text = ""
            tbproductname.Text = ""
            tbstatus.Text = ""
            tbcommissionid.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbemployerid.Text = "" Or tbproductid.Text.Trim = "" Or tbstatus.Text.Trim = "" Or tbcommissionid.Text.Trim = "" Then
            MessageBox.Show("Make Sure to select the commission to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateCommissionProduct(Integer.Parse(tbemployerid.Text.Trim), tbemployername.Text, Integer.Parse(tbproductid.Text.Trim), tbproductname.Text, Integer.Parse(tbstatus.Text.Trim), Integer.Parse(tbcommissionid.Text.Trim))
            loadCommissionProducts()
            Button2.Enabled = False
            Button1.Enabled = True
            tbemployerid.Text = ""
            tbproductid.Text = ""
            tbemployername.Text = ""
            tbproductname.Text = ""
            tbstatus.Text = ""
            tbcommissionid.Text = ""
        End If
    End Sub

    Private Sub cbstatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbstatus.SelectedIndexChanged
        tbstatus.Text = cbstatus.Text
    End Sub

    Private Sub tbstatus_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbstatus.TextChanged
        If tbstatus.Text.Trim = "1" Then
            Label3.Text = "Active"
        ElseIf tbstatus.Text.Trim = "0" Then
            Label3.Text = "Not Active"
        Else
            Label3.Text = "text here"
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        Try
            If DataGridView1.Rows.Count > 0 Then
                tbcommissionid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString
                tbemployerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString
                tbemployername.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString
                tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString
                tbproductname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString
                If Boolean.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value) = True Then
                    tbstatus.Text = 1
                Else
                    tbstatus.Text = 0
                End If
                Button1.Enabled = False
                Button2.Enabled = True
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        tbproductid.Text = ""
        tbproductname.Text = ""
         Button1.Enabled = True
        Button2.Enabled = False
        tbemployername.Text = ""
        tbproductname.Text = ""
        tbstatus.Text = ""
        Label3.Text = "text here"
        tbcommissionid.Text = ""
    End Sub
End Class