Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class EmploymentType

    Dim selected As Boolean = False

    Private Sub loadEmploymentType()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from EmploymentType Order By EmploymentTypeName ASC"

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
            MessageBox.Show(ex.Message, "loadEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbemploymenttypeid.Text.Trim = "" Or tbemploymenttype.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveEmploymentType(tbemploymenttypeid.Text.Trim, tbemploymenttype.Text.Trim, "Active", tbmlma.Text.Trim, tbmlme.Text.Trim)
                    loadEmploymentType()
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
            tbemploymenttypeid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbemploymenttype.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbmlma.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbmlmaid.Text = getValueStringOtherSystems("select AutoID from EmploymentType where EmploymentTypeName = '" & tbmlma.Text.Trim & "'", conMLMAdvances)
            tbmlme.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            tbmlmeid.Text = getValueStringOtherSystems("select AutoID from EmploymentType where EmploymentTypeName = '" & tbmlme.Text.Trim & "'", conMLMEducational)

            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbemploymenttypeid.Text = "" Or tbemploymenttype.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the country to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateEmploymentType(tbemploymenttypeid.Text.Trim, tbemploymenttype.Text.Trim, tbmlma.Text.Trim, tbmlme.Text.Trim)
                    loadEmploymentType()
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
        tbemploymenttypeid.Text = ""
        tbemploymenttype.Text = ""
        selected = False

        tbmlmaid.Text = ""
        tbmlma.Text = ""
        tbmlmeid.Text = ""
        tbmlme.Text = ""
    End Sub


    Private Sub tbbankname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbemploymenttype.TextChanged
        If Button1.Enabled = True Then
            If tbemploymenttype.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tbemploymenttypeid.Text = getNextEmploymentTypeID(tbemploymenttype.Text.Substring(0, 3))
                End If
            End If
        End If
    End Sub

    Public Function getNextEmploymentTypeID(ByVal str As String) As String
        Dim qry As String = "select * from EmploymentType where EmploymentTypeID like '" & str & "%' order by EmploymentTypeID DESC"

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
                Dim go As Boolean = False
                While go = False
                    perc = str & int
                    If CheckID("select * from EmploymentType where EmploymentTypeID = '" & perc & "'") = True Then
                        int = int + 1
                    Else
                        go = True
                    End If
                End While
            Else
                perc = str & "1"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "getNextEmploymentTypeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getNextEmploymentTypeID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return perc
    End Function

    Public Sub saveEmploymentType(ByVal EmploymentTypeID As String, ByVal EmploymentTypeName As String, ByVal Status As String, ByVal MLMAEmpType As String, ByVal MLMEEmpType As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into EmploymentType values ('" & EmploymentTypeID & "','" & EmploymentTypeName & "','" & Status & "','" & MLMAEmpType & "','" & MLMEEmpType & "')"
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
            MessageBox.Show(ex.Message, "saveEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateEmploymentType(ByVal EmploymentTypeID As String, ByVal EmploymentTypeName As String, ByVal MLMAEmpType As String, ByVal MLMEEmpType As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update EmploymentType set EmploymentTypeName = '" & EmploymentTypeName & "',MLMAEmpType = '" & MLMAEmpType & "',MLMEEmpType = '" & MLMEEmpType & "' where EmploymentTypeID = '" & EmploymentTypeID & "'"
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
            MessageBox.Show(ex.Message, "updateEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub EmploymentType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadEmploymentType()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New EmploymentTypeList
        vrl.MdiParent = mform
        vrl.whichform = "MLMA"
        vrl.etype = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New EmploymentTypeList
        vrl.MdiParent = mform
        vrl.whichform = "MLME"
        vrl.etype = Me
        vrl.Show()
    End Sub
End Class