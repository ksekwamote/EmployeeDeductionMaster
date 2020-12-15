Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Defaulters_List

    Dim expected, actual, balance As Decimal
    Dim ds As New DataSet

    Public Function expectedDeductions(ByVal adate As Date, ByVal employerid As Integer) As DataSet
        Dim qry As String = "select * from ExpectedDeductions where ForMonth = '" & adate & "' and EmployerID = '" & employerid & "' and CustomerID in (select CustomerID from CustomerDetails where EmployerID = '" & employerid & "')"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Batches")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.StackTrace, "batches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.Message, "batches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function

    Public Sub refreshData()
        Dim dt As New DataTable
        ds.Clear()
        DataGridView1.DataSource = Nothing

        dt.Columns.Add("CustomerID").DataType = System.Type.GetType("System.String")
        dt.Columns.Add("Expected").DataType = System.Type.GetType("System.String")
        dt.Columns.Add("Actual").DataType = System.Type.GetType("System.String")
        dt.Columns.Add("Balance").DataType = System.Type.GetType("System.String")
        dt.Columns.Add("CustomerName").DataType = System.Type.GetType("System.String")
        
        ds = expectedDeductions(DateTimePicker1.Value, employerid)

        If ds.Tables(0).Rows.Count > 0 Then
            ProgressBar1.Maximum = ds.Tables(0).Rows.Count - 1
            For rec As Integer = 0 To ds.Tables(0).Rows.Count - 1

                expected = Decimal.Parse(ds.Tables(0).Rows(rec).Item(2).ToString)
                actual = getActualDeduction(DateTimePicker1.Value, ds.Tables(0).Rows(rec).Item(1).ToString, employerid)
                balance = expected - actual

                addData(ds.Tables(0).Rows(rec).Item(1).ToString, expected, actual, balance, getCustomerName(ds.Tables(0).Rows(rec).Item(1).ToString), dt)

                ProgressBar1.Value = rec
            Next
            ProgressBar1.Value = 0

            Dim dataView As New DataView(dt)
            dt = dataView.ToTable()
            DataGridView1.DataSource = dt

            Dim bs As New BindingSource
            bs.DataSource = dt.DefaultView
            BindingNavigator1.BindingSource = bs
            DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If
    End Sub


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, employername)
        Close()
    End Sub

    Public Sub addData(ByVal CustomerID As String, ByVal Expected As String, ByVal Actual As String, ByVal Balance As String, ByVal CustomerName As String, ByVal dt As DataTable)

        Dim newRow As DataRow
        newRow = dt.NewRow

        newRow("CustomerID") = CustomerID
        newRow("Expected") = Expected
        newRow("Actual") = Actual
        newRow("Balance") = Balance
        newRow("CustomerName") = CustomerName

        dt.Rows.Add(newRow)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If checkExpectedDeductionDate(DateTimePicker1.Value, employerid) = True And checkActualDeductionDate(DateTimePicker1.Value, employerid) = True Then
            refreshData()
        Else
            MessageBox.Show("Actual Deductions not yet uploaded", "batches()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
    End Sub

    Private Sub Defaulters_List_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If getLastActualDeductionDate(employerid) = getLastLodgingDate(employerid) Then
            DateTimePicker1.Value = getLastDateOfMonth(getLastActualDeductionDate(employerid))
        Else
            DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
        End If
    End Sub
End Class