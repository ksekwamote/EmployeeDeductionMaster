Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class PrintingStatus

    Private Sub PrintingStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadRequested()
        loadPrinting()
        loadPrinted()
    End Sub

    Private Sub loadRequested()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,Name as PrintFor,'Sheet'+SheetNumber as SheetNumber,RequestTime,RequestedBy,Status from PrintingRequests inner join Users on PrintingRequests.Username = Users.Username where Status = 'Requested' order by RequestTime ASC"

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
            MessageBox.Show(ex.Message, "loadRequested()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadRequested()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadPrinting()
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select AutoID,Name as PrintFor,'Sheet'+SheetNumber as SheetNumber,PrintedDate,RequestedBy,RequestTime,Status from PrintingRequests inner join Users on PrintingRequests.Username = Users.Username where Status = 'Printing' order by PrintedDate ASC"

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
                DataGridView2.DataSource = bs
                BindingNavigator2.BindingSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView2)
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadPrinting()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadPrinting()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadPrinted()
        DataGridView3.DataSource = Nothing
        Dim qry As String = "select AutoID,Name as PrintFor,'Sheet'+SheetNumber as SheetNumber,PrintedDate,RequestedBy,RequestTime,Status from PrintingRequests inner join Users on PrintingRequests.Username = Users.Username where Status = 'Printed' order by PrintedDate DESC"

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
                DataGridView3.DataSource = bs
                BindingNavigator3.BindingSource = bs
                DataGridView3.Columns(DataGridView3.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView3)
            Else
                DataGridView3.DataSource = Nothing
                BindingNavigator3.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadPrinted()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadPrinted()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If decision("You are printing " & DataGridView1.CurrentRow.Cells(2).Value & " for " & DataGridView1.CurrentRow.Cells(1).Value & " ?") = DialogResult.Yes Then
            updatePrintingStatusAndDate(Date.Now, "Printing", Integer.Parse(DataGridView1.CurrentRow.Cells(0).Value))
            loadRequested()
            loadPrinting()
            loadPrinted()
        End If
    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        If decision("Done printing " & DataGridView2.CurrentRow.Cells(2).Value & " for " & DataGridView2.CurrentRow.Cells(1).Value & " ?") = DialogResult.Yes Then
            updatePrintingStatusAlone("Printed", Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value))
            loadRequested()
            loadPrinting()
            loadPrinted()
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        loadRequested()
        loadPrinting()
        loadPrinted()
    End Sub
End Class