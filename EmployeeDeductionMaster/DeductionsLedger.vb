Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class DeductionsLedger

    Private Sub DeductionsLedger_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub loadActualDeductions(ByVal CustomerID As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,CustomerID,ActualDeduction,ForMonth,TransactionDate from ActualDeductions where CustomerID = '" & CustomerID & "' order by TransactionDate ASC"

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
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadActualDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadActualDeductions()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadActualDeductions(TextBox1.Text.Trim)
    End Sub
End Class