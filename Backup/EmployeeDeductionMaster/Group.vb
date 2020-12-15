Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class Group

    Dim selected As Boolean = False

    Private Sub loadGroup()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Groups Order By GroupName ASC"

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
            MessageBox.Show(ex.Message, "loadGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbgroupid.Text.Trim = "" Or tbgroup.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveGroup(tbgroupid.Text.Trim, tbgroup.Text.Trim, "Active")
                    loadGroup()
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
            tbgroupid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbgroup.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbgroupid.Text = "" Or tbgroup.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the group to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateGroup(tbgroupid.Text.Trim, tbgroup.Text.Trim)
                    loadGroup()
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
        tbgroupid.Text = ""
        tbgroup.Text = ""
        selected = False
    End Sub


    Private Sub tbbankname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbgroup.TextChanged
        If Button1.Enabled = True Then
            If tbgroup.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tbgroupid.Text = getNextGroupID(tbgroup.Text.Substring(0, 3))
                End If
            End If
        End If
    End Sub

    Public Function getNextGroupID(ByVal str As String) As String
        Dim qry As String = "select * from Groups where GroupID like '" & str & "%' order by GroupID DESC"

        Dim perc As String = ""

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            SQLAdapter.Fill(ds, "MS")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            If ds.Tables(0).Rows.Count > 0 Then
                Dim int As Integer = 0
                int = Integer.Parse(ds.Tables(0).Rows(0).Item(0).ToString.Substring(3)) + 1
                perc = str & int
            Else
                perc = str & "1"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "getNextGroupID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getNextGroupID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveGroup(ByVal GroupID As String, ByVal GroupName As String, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into Groups values ('" & GroupID & "','" & GroupName & "','" & Status & "')"
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
            MessageBox.Show(ex.Message, "saveGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateGroup(ByVal GroupID As String, ByVal GroupName As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Groups set GroupName = '" & GroupName & "' where GroupID = '" & GroupID & "'"
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
            MessageBox.Show(ex.Message, "updateGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateGroup()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Group_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadGroup()
    End Sub
End Class