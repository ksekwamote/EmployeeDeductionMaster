Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class Supervisors

    Private Sub loadSuperviors()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Supervisors"

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
            MessageBox.Show(ex.Message, "loadSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Branches_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadSuperviors()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbname.Text.Trim = "" Or tbtitle.Text.Trim = "" Or tbemail.Text.Trim = "" Or tbdepartment.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            saveSuperviors(tbname.Text.Trim, tbtitle.Text.Trim, tbemail.Text.Trim, tbdepartment.Text.Trim)
            loadSuperviors()
            tbautoid.Text = ""
            tbname.Text = ""
            tbtitle.Text = ""
            tbemail.Text = ""
            tbdepartment.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbautoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbtitle.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbemail.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbdepartment.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbautoid.Text = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbname.Text.Trim = "" Or tbtitle.Text.Trim = "" Or tbemail.Text.Trim = "" Or tbautoid.Text.Trim = "" Or tbdepartment.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                updateSuperviors(tbname.Text.Trim, tbtitle.Text.Trim, tbemail.Text.Trim, tbdepartment.Text.Trim, Integer.Parse(tbautoid.Text.Trim))
                loadSuperviors()
                tbautoid.Text = ""
                tbname.Text = ""
                tbtitle.Text = ""
                tbemail.Text = ""
                tbdepartment.Text = ""
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = False
        Button1.Enabled = True
        tbautoid.Text = ""
        tbname.Text = ""
        tbtitle.Text = ""
        tbemail.Text = ""
        tbdepartment.Text = ""
    End Sub

    Public Sub updateSuperviors(ByVal SupervisorName As String, ByVal Title As String, ByVal EmailAddress As String, ByVal Department As String, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Supervisors set SupervisorName = '" & SupervisorName & "',Title = '" & Title & "',EmailAddress = '" & EmailAddress & "',Department = '" & Department & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If
            MessageBox.Show("Details updated", "updateSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateSuperviors()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try

    End Sub
End Class