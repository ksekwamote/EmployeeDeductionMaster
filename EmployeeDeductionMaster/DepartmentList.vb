Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class DepartmentList

    Public whichform As String
    Public cust As Customer
    Public autoid As String
    Public dep As Department

    Private Sub loadDepartment(ByVal tform As String, ByVal con As SqlClient.SqlConnection)
        DataGridView1.DataSource = Nothing
        Dim qry As String

        If tform = "Customer" Then
            qry = "select * from Department order by DepartmentName ASC"
        ElseIf tform = "RMO" Then
            qry = "select * from Department order by Department_Name ASC"
        ElseIf tform = "RMB" Then
            qry = "select * from Department order by Department_Name ASC"
        ElseIf tform = "MLMA" Then
            qry = "select * from EmployerDepartment order by DepartmentName ASC"
        ElseIf tform = "MLME" Then
            qry = "select * from EmployerDepartment order by DepartmentName ASC"
        ElseIf tform = "FPM" Then
            qry = "select * from Departments order by DepartmentName ASC"
        End If

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

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "Customer" Then
                cust.tbdepartmentid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cust.tbdepartment.Text = DataGridView1.CurrentRow.Cells(1).Value
                cust.RMODepartment = DataGridView1.CurrentRow.Cells(3).Value
                cust.RMBDepartment = DataGridView1.CurrentRow.Cells(4).Value
                cust.MLMADepartment = DataGridView1.CurrentRow.Cells(5).Value
                cust.MLMEDepartment = DataGridView1.CurrentRow.Cells(6).Value
                cust.FPMDepartment = DataGridView1.CurrentRow.Cells(7).Value
                Close()
            ElseIf whichform = "RMO" Then
                dep.tbrmoid.Text = DataGridView1.CurrentRow.Cells(0).Value
                dep.tbrmo.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "RMB" Then
                dep.tbrmbid.Text = DataGridView1.CurrentRow.Cells(0).Value
                dep.tbrmb.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLMA" Then
                dep.tbmlmaid.Text = DataGridView1.CurrentRow.Cells(0).Value
                dep.tbmlma.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLME" Then
                dep.tbmlmeid.Text = DataGridView1.CurrentRow.Cells(0).Value
                dep.tbmlme.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "FPM" Then
                dep.tbfpmid.Text = DataGridView1.CurrentRow.Cells(0).Value
                dep.tbfpm.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub DepartmentList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If whichform = "Customer" Then
            loadDepartment(whichform, con)
        ElseIf whichform = "RMO" Then
            loadDepartment(whichform, conRMOrange)
        ElseIf whichform = "RMB" Then
            loadDepartment(whichform, conRMBeMobile)
        ElseIf whichform = "MLMA" Then
            loadDepartment(whichform, conMLMAdvances)
        ElseIf whichform = "MLME" Then
            loadDepartment(whichform, conMLMEducational)
        ElseIf whichform = "FPM" Then
            loadDepartment(whichform, conFPM)
        End If
    End Sub
End Class