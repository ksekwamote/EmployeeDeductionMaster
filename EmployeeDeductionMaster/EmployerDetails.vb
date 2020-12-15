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
        If tbEmployerName.Text.Trim = "" Or tbcommissionpercent.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            saveEmployers(tbEmployerName.Text.Trim, Decimal.Parse(tbcommissionpercent.Text.Trim), Integer.Parse(tbproductid.Text.Trim), "Yes", tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim)
            loadEmployers()
            clearTexts()
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbemployerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbEmployerName.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbcommissionpercent.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tballowtoedit.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            tbrmoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            tbrmo.Text = getValueStringOtherSystems("select Employer_Name from Employer where Code = '" & tbrmoid.Text.Trim & "'", conRMOrange)
            tbrmbid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
            tbrmb.Text = getValueStringOtherSystems("select Employer_Name from Employer where Code = '" & tbrmbid.Text.Trim & "'", conRMBeMobile)
            tbmlmaid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
            tbmlma.Text = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & tbmlmaid.Text.Trim & "'", conMLMAdvances)
            tbmlmeid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(8).Value.ToString()
            tbmlme.Text = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & tbmlmeid.Text.Trim & "'", conMLMEducational)
            Button2.Enabled = True
            Button1.Enabled = False
            If tballowtoedit.Text.Trim = "Yes" Then
                tbEmployerName.Enabled = True
                tbcommissionpercent.Enabled = True
            Else
                tbEmployerName.Enabled = False
                tbcommissionpercent.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbemployerid.Text = "" Or tbcommissionpercent.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            updateEmployer(tbEmployerName.Text.Trim, Decimal.Parse(tbcommissionpercent.Text.Trim), Integer.Parse(tbproductid.Text.Trim), "Yes", Integer.Parse(tbemployerid.Text.Trim), tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim)
            loadEmployers()
            clearTexts()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        clearTexts()
    End Sub

    Public Sub clearTexts()
        Button2.Enabled = False
        Button1.Enabled = True
        tbemployerid.Text = ""
        tbEmployerName.Text = ""
        tbcommissionpercent.Text = ""
        tbproductid.Text = ""
        tballowtoedit.Text = ""
        tbrmoid.Text = ""
        tbrmbid.Text = ""
        tbmlmaid.Text = ""
        tbmlmeid.Text = ""
    End Sub

    Public Sub updateEmployer(ByVal EmployerName As String, ByVal CommissionPercent As Decimal, ByVal ProductID As Integer, ByVal AllowToEdit As String, ByVal EmployerID As Integer, ByVal RMOEmployer As String, ByVal RMBEmployer As String, ByVal MLMAEmployer As String, ByVal MLMEEmployer As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update Employers set EmployerName = '" & EmployerName & "',CommissionPercent = '" & CommissionPercent & "',ProductID = '" & ProductID & "',AllowToEdit = '" & AllowToEdit & "',RMOEmployer = '" & RMOEmployer & "',RMBEmployer = '" & RMBEmployer & "',MLMAEmployer = '" & MLMAEmployer & "',MLMEEmployer = '" & MLMEEmployer & "' where AutoID = '" & EmployerID & "'"
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "RMO"
        vrl.emp = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "RMB"
        vrl.emp = Me
        vrl.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "MLMA"
        vrl.emp = Me
        vrl.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "MLME"
        vrl.emp = Me
        vrl.Show()
    End Sub
End Class