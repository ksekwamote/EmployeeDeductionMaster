Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class LocationList

    Public whichform As String
    Public cust As Customer
    Public autoid As String

    Private Sub loadLocation()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from Location order by LocationName ASC"

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
            If whichform = "Customer" Then
                cust.tbtownid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cust.tbtown.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub LocationList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadLocation()
    End Sub
End Class