Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class EmployerList

    Public whichform As String
    Public cust As Customer
    Public gpg As EmployerProductGrouping
    Public autoid As String
    Public emp As EmployerDetails
    Public ep As EmployerProducts

    Private Sub loadEmployers(ByVal tform As String, ByVal con As SqlClient.SqlConnection)
        DataGridView1.DataSource = Nothing
        Dim qry As String

        If tform = "Customer" Then
            qry = "select * from Employers order by EmployerName ASC"
        ElseIf tform = "RMO" Then
            qry = "select * from Employer order by Employer_Name ASC"
        ElseIf tform = "RMB" Then
            qry = "select * from Employer order by Employer_Name ASC"
        ElseIf tform = "MLMA" Then
            qry = "select * from Employer order by Employer ASC"
        ElseIf tform = "MLME" Then
            qry = "select * from Employer order by Employer ASC"
        ElseIf tform = "Employer Products" Then
            qry = "select * from Employers order by EmployerName ASC"
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
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "Customer" Then
                cust.tbemployerid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cust.tbemployer.Text = DataGridView1.CurrentRow.Cells(1).Value
                cust.RMOEmployer = DataGridView1.CurrentRow.Cells(5).Value
                cust.RMBEmployer = DataGridView1.CurrentRow.Cells(6).Value
                cust.MLMAEmployer = DataGridView1.CurrentRow.Cells(7).Value
                cust.MLMEEmployer = DataGridView1.CurrentRow.Cells(8).Value
                cust.MLMAEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & DataGridView1.CurrentRow.Cells(7).Value & "'", conMLMAdvances)
                cust.MLMEEmployerName = getValueStringOtherSystems("select Employer from Employer where AutoID = '" & DataGridView1.CurrentRow.Cells(8).Value & "'", conMLMEducational)
                Close()
            ElseIf whichform = "Employer Product Grouping" Then
                gpg.tbemployerid.Text = DataGridView1.CurrentRow.Cells(0).Value
                gpg.tbemployername.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "RMO" Then
                emp.tbrmoid.Text = DataGridView1.CurrentRow.Cells(0).Value
                emp.tbrmo.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "RMB" Then
                emp.tbrmbid.Text = DataGridView1.CurrentRow.Cells(0).Value
                emp.tbrmb.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLMA" Then
                emp.tbmlmaid.Text = DataGridView1.CurrentRow.Cells(0).Value
                emp.tbmlma.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLME" Then
                emp.tbmlmeid.Text = DataGridView1.CurrentRow.Cells(0).Value
                emp.tbmlme.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "Employer Products" Then
                ep.tbemployerid.Text = DataGridView1.CurrentRow.Cells(0).Value
                ep.tbemployername.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub EmployerList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If whichform = "Customer" Then
            loadEmployers(whichform, con)
        ElseIf whichform = "RMO" Then
            loadEmployers(whichform, conRMOrange)
        ElseIf whichform = "RMB" Then
            loadEmployers(whichform, conRMBeMobile)
        ElseIf whichform = "MLMA" Then
            loadEmployers(whichform, conMLMAdvances)
        ElseIf whichform = "MLME" Then
            loadEmployers(whichform, conMLMEducational)
        ElseIf whichform = "Employer Products" Then
            loadEmployers(whichform, con)
        End If
    End Sub
End Class