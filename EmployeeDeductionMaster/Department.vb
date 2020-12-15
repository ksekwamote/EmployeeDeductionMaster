Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class Department

    Dim selected As Boolean = False

    Private Sub loadDepartment()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Department Order By DepartmentName ASC"

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
            MessageBox.Show(ex.Message, "loadDepartment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadDepartment()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If tbdepartmentid.Text.Trim = "" Or tbdepartmentname.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Or tbfpmid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveDepartment(tbdepartmentid.Text.Trim, tbdepartmentname.Text.Trim, "Active", tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim, tbfpmid.Text.Trim)
                    loadDepartment()
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
        Try
            If DataGridView1.Rows.Count > 0 Then
                selected = True
                tbdepartmentid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
                tbdepartmentname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
                tbfpmid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
                tbfpm.Text = getValueStringOtherSystems("select DepartmentName from Departments where DepID = '" & tbfpmid.Text.Trim & "'", conFPM)
                tbrmoid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
                tbrmo.Text = getValueStringOtherSystems("select Department_Name from Department where DepartmentCode = '" & tbrmoid.Text.Trim & "'", conRMOrange)
                tbrmbid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
                tbrmb.Text = getValueStringOtherSystems("select Department_Name from Department where DepartmentCode = '" & tbrmbid.Text.Trim & "'", conRMBeMobile)
                tbmlmaid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
                tbmlma.Text = getValueStringOtherSystems("select DepartmentName from EmployerDepartment where DepartmentID = '" & tbmlmaid.Text.Trim & "'", conMLMAdvances)
                tbmlmeid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
                tbmlme.Text = getValueStringOtherSystems("select DepartmentName from EmployerDepartment where DepartmentID = '" & tbmlmeid.Text.Trim & "'", conMLMEducational)
                Button2.Enabled = True
                Button1.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If tbdepartmentid.Text = "" Or tbdepartmentname.Text.Trim = "" Or tbrmoid.Text.Trim = "" Or tbrmbid.Text.Trim = "" Or tbmlmaid.Text.Trim = "" Or tbmlmeid.Text.Trim = "" Or tbfpmid.Text.Trim = "" Then
                MessageBox.Show("Make Sure to select the country to update changes and enter the changes", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateDepartment(tbdepartmentid.Text.Trim, tbdepartmentname.Text.Trim, tbrmoid.Text.Trim, tbrmbid.Text.Trim, tbmlmaid.Text.Trim, tbmlmeid.Text.Trim, tbfpmid.Text.Trim)
                    loadDepartment()
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
        tbdepartmentid.Text = ""
        tbdepartmentname.Text = ""
        selected = False

        tbrmoid.Text = ""
        tbrmbid.Text = ""
        tbmlmaid.Text = ""
        tbmlmeid.Text = ""
        tbfpmid.Text = ""

        tbrmo.Text = ""
        tbrmb.Text = ""
        tbmlma.Text = ""
        tbmlme.Text = ""
        tbfpm.Text = ""

    End Sub


    Private Sub tbbankname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbdepartmentname.TextChanged
        If Button1.Enabled = True Then
            If tbdepartmentname.Text.Trim.Length > 3 Then
                If selected = True Then

                Else
                    tbdepartmentid.Text = getNextDepartmentID(tbdepartmentname.Text.Substring(0, 3))
                End If
            End If
        End If
    End Sub

    Private Sub Department_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadDepartment()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "FPM"
        vrl.dep = Me
        vrl.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "RMO"
        vrl.dep = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "RMB"
        vrl.dep = Me
        vrl.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "MLMA"
        vrl.dep = Me
        vrl.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim vrl As New DepartmentList
        vrl.MdiParent = mform
        vrl.whichform = "MLME"
        vrl.dep = Me
        vrl.Show()
    End Sub
End Class