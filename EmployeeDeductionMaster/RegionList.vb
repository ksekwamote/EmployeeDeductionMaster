Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class RegionList

    Public whichform As String
    Public cust As Customer
    Public reg As Region
    Public autoid As String

    Private Sub loadRegion(ByVal tform As String, ByVal con As SqlClient.SqlConnection)
        DataGridView1.DataSource = Nothing
        Dim qry As String
        If tform = "Customer" Then
            qry = "select * from Region order by RegionName ASC"
        ElseIf tform = "RMO" Then
            qry = "select * from Region order by RegionName ASC"
        ElseIf tform = "RMB" Then
            qry = "select * from Region order by RegionName ASC"
        ElseIf tform = "FPM" Then
            qry = "select * from Regions order by RegionName ASC"
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
            MessageBox.Show(ex.Message, "loadRegion()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadRegion()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "Customer" Then
                cust.tbregionid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cust.tbregion.Text = DataGridView1.CurrentRow.Cells(1).Value
                cust.RMORegionID = DataGridView1.CurrentRow.Cells(3).Value
                cust.RMBRegionID = DataGridView1.CurrentRow.Cells(4).Value
                cust.FPMRegionID = DataGridView1.CurrentRow.Cells(5).Value
                Close()
            ElseIf whichform = "RMO" Then
                reg.tbrmoid.Text = DataGridView1.CurrentRow.Cells(0).Value
                reg.tbrmo.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "RMB" Then
                reg.tbrmbid.Text = DataGridView1.CurrentRow.Cells(0).Value
                reg.tbrmb.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "FPM" Then
                reg.tbfpmid.Text = DataGridView1.CurrentRow.Cells(0).Value
                reg.tbfpm.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub RegionList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If whichform = "Customer" Then
            loadRegion(whichform, con)
        ElseIf whichform = "RMO" Then
            loadRegion(whichform, conRMOrange)
        ElseIf whichform = "RMB" Then
            loadRegion(whichform, conRMBeMobile)
        ElseIf whichform = "FPM" Then
            loadRegion(whichform, conFPM)
        End If
    End Sub
End Class