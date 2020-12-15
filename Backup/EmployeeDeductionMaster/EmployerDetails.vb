Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class EmployerDetails

    Private Sub loadEmployers()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Employers"

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

            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbEmployerName.Text.Trim = "" Or tbcommissionpercent.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            saveEmployers(tbEmployerName.Text.Trim, Decimal.Parse(tbcommissionpercent.Text.Trim), Integer.Parse(tbproductid.Text.Trim), "Yes")
            loadEmployers()
            Button2.Enabled = False
            Button1.Enabled = True
            tbemployerid.Text = ""
            tbEmployerName.Text = ""
            tbcommissionpercent.Text = ""
            tbproductid.Text = ""
            tballowtoedit.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbemployerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbEmployerName.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbcommissionpercent.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tballowtoedit.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbemployerid.Text = "" Or tbcommissionpercent.Text.Trim = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tballowtoedit.Text.Trim = "Yes" Then
                updateEmployer(tbEmployerName.Text.Trim, Decimal.Parse(tbcommissionpercent.Text.Trim), Integer.Parse(tbproductid.Text.Trim), "Yes", Integer.Parse(tbemployerid.Text.Trim))
                loadEmployers()
                Button2.Enabled = False
                Button1.Enabled = True
                tbemployerid.Text = ""
                tbEmployerName.Text = ""
                tbcommissionpercent.Text = ""
                tbproductid.Text = ""
                tballowtoedit.Text = ""
            Else
                MessageBox.Show("Not Allowed To Edit This Employer", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = False
        Button1.Enabled = True
        tbemployerid.Text = ""
        tbEmployerName.Text = ""
        tbcommissionpercent.Text = ""
        tbproductid.Text = ""
        tballowtoedit.Text = ""
    End Sub

    Public Sub updateEmployer(ByVal EmployerName As String, ByVal CommissionPercent As Decimal, ByVal ProductID As Integer, ByVal AllowToEdit As String, ByVal EmployerID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Employers set EmployerName = '" & EmployerName & "',CommissionPercent = '" & CommissionPercent & "',ProductID = '" & ProductID & "',AllowToEdit = '" & AllowToEdit & "' where AutoID = '" & EmployerID & "'"
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
            MessageBox.Show(ex.Message, "updateEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateEmployer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try

    End Sub

    Private Sub EmplyerDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadEmployers()
        loadProducts()
        tbproductid.Text = ""
    End Sub

    Public Sub loadProducts()
        Dim qry As String = "select * from Products where AutoAllocate = 1"

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

    Private Sub cbproduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbproduct.SelectedIndexChanged
        tbproductid.Text = cbproduct.SelectedValue.ToString
    End Sub
End Class