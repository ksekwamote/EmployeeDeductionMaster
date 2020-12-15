Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class Progress

    Dim total, lastexpected As Decimal
    Dim index As Integer = 0
    Dim productid As Integer = 0


    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
        Button3.Enabled = True
    End Sub

    Private Sub loadProducts()
        Dim qry As String = "select AutoID,Products, Products as TotalAmendments, Products as Count, Products as TerminationsCount from Products Where Status = 'Active' order by AllocationOrder ASC"
        DataGridView2.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView2.DataSource = bs
                BindingNavigator1.BindingSource = bs

                Dim chk1 As New DataGridViewCheckBoxColumn()
                chk1.HeaderText = "Choose?"
                chk1.Name = "confirm"
                If DataGridView2.Columns("confirm") Is Nothing Then
                    DataGridView2.Columns.Add(chk1)
                End If

                For Each row As DataGridViewRow In DataGridView2.Rows
                    row.Cells(2).Value = getProductAmendment(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), employerid)
                    row.Cells(3).Value = getProductsCount(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), employerid)
                    row.Cells(4).Value = getTerminationsCount(DateTimePicker1.Value, Integer.Parse(row.Cells(0).Value.ToString), employerid)
                    total = total + Decimal.Parse(row.Cells(2).Value)
                Next
                'DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(2).Value = getLastExpectedDeduction(getLastDateOfMonth(addMonth(DateTimePicker1.Value, -1)))
                'DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(3).Value = getLastExpectedDeductionCount(getLastDateOfMonth(addMonth(DateTimePicker1.Value, -1)))
                
                DataGridView2.Columns(0).ReadOnly = True
                DataGridView2.Columns(1).ReadOnly = True
                DataGridView2.Columns(2).ReadOnly = True
                TextBox1.Text = total
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        loadProducts()
        Button3.Enabled = False
        DateTimePicker1.Enabled = False
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView2, "")
    End Sub

    Private Sub Progress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = getLastDateOfMonth(addMonth(getLastLodgingDate(employerid), 1))
    End Sub
End Class