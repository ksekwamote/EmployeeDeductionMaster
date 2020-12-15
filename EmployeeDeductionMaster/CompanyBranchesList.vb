Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class CompanyBranchesList

    Public whichform As String
    Public cb As Branches
    Public autoid As String

    Private Sub loadCompanyBranches(ByVal tform As String, ByVal con As SqlClient.SqlConnection)
        DataGridView1.DataSource = Nothing
        Dim qry As String

        If tform = "RMO" Then
            qry = "select * from CompanyBranches order by BranchName ASC"
        ElseIf tform = "RMB" Then
            qry = "select * from CompanyBranches order by BranchName ASC"
        ElseIf tform = "MLMA" Then
            qry = "select * from CompanyBranches order by BranchName ASC"
        ElseIf tform = "MLME" Then
            qry = "select * from CompanyBranches order by BranchName ASC"
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
            MessageBox.Show(ex.Message, "loadLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadLocation()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "RMO" Then
                cb.tbrmoid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cb.tbrmo.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "RMB" Then
                cb.tbrmbid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cb.tbrmb.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLMA" Then
                cb.tbmlmaid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cb.tbmlma.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLME" Then
                cb.tbmlmeid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cb.tbmlme.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub CompanyBranchesList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If whichform = "RMO" Then
            loadCompanyBranches(whichform, conRMOrange)
        ElseIf whichform = "RMB" Then
            loadCompanyBranches(whichform, conRMBeMobile)
        ElseIf whichform = "MLMA" Then
            loadCompanyBranches(whichform, conMLMAdvances)
        ElseIf whichform = "MLME" Then
            loadCompanyBranches(whichform, conMLMEducational)
        End If
    End Sub
End Class