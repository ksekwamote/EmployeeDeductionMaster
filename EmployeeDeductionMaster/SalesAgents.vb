Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class SalesAgents

    Dim selected As Boolean = False

    Private Sub loadSalesAgents()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from SalesAgents Order By Fullname ASC"

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
            MessageBox.Show(ex.Message, "loadSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbidnumber.Text.Trim = "" Or tbfullname.Text.Trim = "" Or tbcontact.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveSalesAgents(tbidnumber.Text.Trim, tbfullname.Text.Trim, tbcontact.Text.Trim, "Active", logedinname, Date.Now)
                    MessageBox.Show("Done saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    loadSalesAgents()
                    clearText()
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
            tbidnumber.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbfullname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbcontact.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbautoid.Text.Trim = "" Or tbidnumber.Text.Trim = "" Or tbfullname.Text.Trim = "" Or tbcontact.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the Sales Agent to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateSalesAgents(tbidnumber.Text.Trim, tbfullname.Text.Trim, tbcontact.Text.Trim, Integer.Parse(tbautoid.Text.Trim))
                    MessageBox.Show("Done updating", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    loadSalesAgents()
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
        tbidnumber.Text = ""
        selected = False

        tbidnumber.Text = ""
        tbfullname.Text = ""
        tbcontact.Text = ""

    End Sub

    Public Sub saveSalesAgents(ByVal IDNumber As String, ByVal Fullname As String, ByVal ContactNumber As String, ByVal Status As String, ByVal CapturedBy As String, ByVal CapturedDate As Date)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into SalesAgents values ('" & IDNumber & "','" & Fullname & "','" & ContactNumber & "','" & Status & "','" & CapturedBy & "','" & CapturedDate & "')"
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
            MessageBox.Show(ex.Message, "saveSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateSalesAgents(ByVal IDNumber As String, ByVal Fullname As String, ByVal ContactNumber As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update SalesAgents set IDNumber = '" & IDNumber & "',Fullname = '" & Fullname & "',ContactNumber = '" & ContactNumber & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updateSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSalesAgents()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub SalesAgents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadSalesAgents()
    End Sub
End Class