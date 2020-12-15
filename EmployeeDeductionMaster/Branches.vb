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

    Public Sub saveBranch(ByVal BranchName As String, ByVal PhysicalAddress As String, ByVal BranchPhone As String, ByVal BranchFax As String, ByVal RMOBranch As String, ByVal RMBBranch As String, ByVal MLMABranchID As String, ByVal MLMEBranchID As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            transactioninsert = con.BeginTransaction()
            Dim qryinsert As String = "Insert into CompanyBranches values ('" & BranchName & "','" & PhysicalAddress & "','" & BranchPhone & "','" & BranchFax & "','" & RMOBranch & "','" & RMBBranch & "','" & MLMABranchID & "','" & MLMEBranchID & "')"
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
        If tbbranchname.Text.Trim = "" Or tbphysicaladdress.Text.Trim = "" Or tbbranchphone.Text.Trim = "" Or tbbranchfax.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
            MessageBox.Show("Enter all the details before saving")
        Else
            saveBranch(tbbranchname.Text.Trim, tbphysicaladdress.Text.Trim, tbbranchphone.Text.Trim, tbbranchfax.Text.Trim, tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim)
            loadBranches()
            Button2.Enabled = False
            Button1.Enabled = True
            tbautoid.Text = ""
            tbbranchname.Text = ""
            tbphysicaladdress.Text = ""
            tbbranchphone.Text = ""
            tbbranchfax.Text = ""

            tbrmo.Text = ""
            tbrmoid.Text = ""
            tbrmb.Text = ""
            tbrmbid.Text = ""
            tbmlma.Text = ""
            tbmlmaid.Text = ""
            tbmlme.Text = ""
            tbmlmeid.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbautoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbbranchname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbphysicaladdress.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbbranchphone.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbbranchfax.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()

            tbrmoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            tbrmo.Text = getValueStringOtherSystems("select BranchName from CompanyBranches where AutoID = '" & tbrmoid.Text.Trim & "'", conRMOrange)
            tbrmbid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
            tbrmb.Text = getValueStringOtherSystems("select BranchName from CompanyBranches where AutoID = '" & tbrmbid.Text.Trim & "'", conRMBeMobile)
            tbmlmaid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
            tbmlma.Text = getValueStringOtherSystems("select BranchName from CompanyBranches where AutoID = '" & tbmlmaid.Text.Trim & "'", conMLMAdvances)
            tbmlmeid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(8).Value.ToString()
            tbmlme.Text = getValueStringOtherSystems("select BranchName from CompanyBranches where AutoID = '" & tbmlmeid.Text.Trim & "'", conMLMEducational)

            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbautoid.Text = "" Then
            MessageBox.Show("Make Sure to select the branch to update changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbbranchname.Text.Trim = "" Or tbphysicaladdress.Text.Trim = "" Or tbbranchphone.Text.Trim = "" Or tbbranchfax.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving")
            Else
                updateBranch(tbbranchname.Text.Trim, tbphysicaladdress.Text.Trim, tbbranchphone.Text.Trim, tbbranchfax.Text.Trim, tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim, Integer.Parse(tbautoid.Text.Trim))
                loadBranches()

                Button2.Enabled = False
                Button1.Enabled = True
                tbautoid.Text = ""
                tbbranchname.Text = ""
                tbphysicaladdress.Text = ""
                tbbranchphone.Text = ""
                tbbranchfax.Text = ""

                tbrmo.Text = ""
                tbrmoid.Text = ""
                tbrmb.Text = ""
                tbrmbid.Text = ""
                tbmlma.Text = ""
                tbmlmaid.Text = ""
                tbmlme.Text = ""
                tbmlmeid.Text = ""

            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = False
        Button1.Enabled = True
        tbautoid.Text = ""
        tbbranchname.Text = ""
        tbphysicaladdress.Text = ""
        tbbranchphone.Text = ""
        tbbranchfax.Text = ""

        tbrmo.Text = ""
        tbrmoid.Text = ""
        tbrmb.Text = ""
        tbrmbid.Text = ""
        tbmlma.Text = ""
        tbmlmaid.Text = ""
        tbmlme.Text = ""
        tbmlmeid.Text = ""

    End Sub

    Public Sub updateBranch(ByVal BranchName As String, ByVal PhysicalAddress As String, ByVal BranchPhone As String, ByVal BranchFax As String, ByVal RMOBranch As String, ByVal RMBBranch As String, ByVal MLMABranchID As String, ByVal MLMEBranchID As String, ByVal BranchID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update CompanyBranches set BranchName = '" & BranchName & "',PhysicalAddress = '" & PhysicalAddress & "',BranchPhone = '" & BranchPhone & "',BranchFax = '" & BranchFax & "',RMOBranch = '" & RMOBranch & "',RMBBranch = '" & RMBBranch & "',MLMABranchID = '" & MLMABranchID & "',MLMEBranchID = '" & MLMEBranchID & "' where AutoID = '" & BranchID & "'"
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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim vrl As New CompanyBranchesList
        vrl.MdiParent = mform
        vrl.whichform = "RMO"
        vrl.cb = Me
        vrl.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New CompanyBranchesList
        vrl.MdiParent = mform
        vrl.whichform = "RMB"
        vrl.cb = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New CompanyBranchesList
        vrl.MdiParent = mform
        vrl.whichform = "MLMA"
        vrl.cb = Me
        vrl.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim vrl As New CompanyBranchesList
        vrl.MdiParent = mform
        vrl.whichform = "MLME"
        vrl.cb = Me
        vrl.Show()
    End Sub
End Class