Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Branches

    Private Sub loadBranches()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from CompanyBranches"

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
            MessageBox.Show(ex.Message, "loadBranches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadBranches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Branches_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadBranches()
    End Sub

    Public Sub saveBranch(ByVal BranchName As String, ByVal PhysicalAddress As String, ByVal BranchPhone As String, ByVal BranchFax As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into CompanyBranches values ('" & BranchName & "','" & PhysicalAddress & "','" & BranchPhone & "','" & BranchFax & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            cmdinsert.Transaction = transactioninsert
            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details saved", "saveBranch()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim = "" Or TextBox2.Text.Trim = "" Or TextBox3.Text.Trim = "" Or TextBox4.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            saveBranch(TextBox1.Text.Trim, TextBox2.Text.Trim, TextBox3.Text.Trim, TextBox4.Text.Trim)
            loadBranches()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            TextBox5.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            TextBox1.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            TextBox2.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            TextBox3.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            TextBox4.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox5.Text = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateBranch(TextBox1.Text.Trim, TextBox2.Text.Trim, TextBox3.Text.Trim, TextBox4.Text.Trim, Integer.Parse(TextBox5.Text.Trim))
            loadBranches()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = False
        Button1.Enabled = True
        TextBox5.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Public Sub updateBranch(ByVal BranchName As String, ByVal PhysicalAddress As String, ByVal BranchPhone As String, ByVal BranchFax As String, ByVal BranchID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CompanyBranches set BranchName = '" & BranchName & "',PhysicalAddress = '" & PhysicalAddress & "',BranchPhone = '" & BranchPhone & "',BranchFax = '" & BranchFax & "' where AutoID = '" & BranchID & "'"
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
            MessageBox.Show(ex.Message, "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateBranch()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub
End Class