Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class SalesCustomers

    Dim selected As Boolean = False

    Private Sub loadSalesCustomers()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from SalesCustomers Order By CustomerID ASC"

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
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbcustomerid.Text.Trim = "" Or tbfullname.Text.Trim = "" Or tbcellnumber.Text.Trim = "" Or tbagentid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If getValueBooleanOtherSystems("select * from SalesCustomers where CustomerID = '" & tbcustomerid.Text.Trim & "'", con) = True Then
                    MessageBox.Show("Customer already exists in the system", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                        saveSalesCustomers(tbcustomerid.Text.Trim, tbfullname.Text.Trim, tbcellnumber.Text.Trim, dtpcontacteddate.Value, tbagentid.Text.Trim, logedinname, Date.Now, "Active", dtpvaliditydate.Value)
                        loadSalesCustomers()
                        clearText()
                    End If
                End If

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            selected = True
            tbautoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbcustomerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbfullname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbcellnumber.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            dtpcontacteddate.Value = Date.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString())
            tbagentid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            dtpvaliditydate.Value = Date.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(9).Value)
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbcustomerid.Text.Trim = "" Or tbfullname.Text.Trim = "" Or tbcellnumber.Text.Trim = "" Or tbagentid.Text.Trim = "" Or tbautoid.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the group to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateSalesCustomers(tbcustomerid.Text.Trim, tbfullname.Text.Trim, tbcellnumber.Text.Trim, dtpcontacteddate.Value, tbagentid.Text.Trim, dtpvaliditydate.Value, Integer.Parse(tbautoid.Text.Trim))
                    loadSalesCustomers()
                    clearText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearText()
    End Sub

    Public Sub clearText()
        Button2.Enabled = False
        Button1.Enabled = True
        tbautoid.Text = ""
        tbcustomerid.Text = ""
        tbfullname.Text = ""
        selected = False

        tbcellnumber.Text = ""
        tbagentid.Text = ""

    End Sub

    Public Sub saveSalesCustomers(ByVal CustomerID As String, ByVal FullName As String, ByVal CellNumber As String, ByVal ContactedDate As Date, ByVal AgentIDNumber As String, ByVal CapturedBy As String, ByVal CapturedDate As Date, ByVal Status As String, ByVal ValidityDate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into SalesCustomers values ('" & CustomerID & "','" & FullName & "','" & CellNumber & "','" & ContactedDate & "','" & AgentIDNumber & "','" & CapturedBy & "','" & CapturedDate & "','" & Status & "','" & ValidityDate & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateSalesCustomers(ByVal CustomerID As String, ByVal FullName As String, ByVal CellNumber As String, ByVal ContactedDate As Date, ByVal AgentIDNumber As String, ByVal ValidityDate As Date, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update SalesCustomers set CustomerID = '" & CustomerID & "',FullName = '" & FullName & "',CellNumber = '" & CellNumber & "',ContactedDate = '" & ContactedDate & "',AgentIDNumber = '" & AgentIDNumber & "',ValidityDate = '" & ValidityDate & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSalesCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub SalesCustomers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadSalesCustomers()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New AgentsList
        vrl.MdiParent = mform
        vrl.whichform = "Customers"
        vrl.sc = Me
        vrl.Show()
    End Sub

    Private Sub dtpcontacteddate_ValueChanged(sender As Object, e As EventArgs) Handles dtpcontacteddate.ValueChanged
        dtpvaliditydate.Value = dtpcontacteddate.Value.AddMonths(getSalesCustomerValidity())
    End Sub
End Class