Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class ReceiptsReportForModeOfPayment

    Dim totalamount As Decimal

    Private Sub loadReceiptsPerMode(ByVal sdate As Date, ByVal edate As Date)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select Receipts.ModeID,ModeName,SUM(Amount) from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status not in ('New','Batch','Banking','Declined') and TransactionDate between '" & sdate & "' and '" & edate & "' group by Receipts.ModeID,ModeName order by ModeName"

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
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsPerMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsPerMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadReceiptsPerCustomerAndPaymentMode(ByVal sdate As Date, ByVal edate As Date, ByVal ModeID As Integer)
        DataGridView3.DataSource = Nothing
        Dim qry As String = "select PayerID,ModeName,Amount,TransactionDate,Reference,CapturedTime,CapturedBy from Receipts inner join ModeOfPayment on Receipts.ModeID = ModeOfPayment.ModeID where Receipts.Status not in ('New','Batch','Banking','Declined') and Receipts.ModeID = '" & ModeID & "' and TransactionDate between '" & sdate & "' and '" & edate & "' order by PayerID"

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
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadReceiptsPerCustomerAndPaymentMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadReceiptsPerCustomerAndPaymentMode()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadReceiptsPerMode(DateTimePicker1.Value, DateTimePicker2.Value)
        totalamount = 0

        totalamount = 0
        For Each row As DataGridViewRow In DataGridView2.Rows
            totalamount = totalamount + Decimal.Parse(row.Cells(2).Value)
        Next
        Label4.Text = getCurrency(totalamount)

        DataGridView3.DataSource = Nothing
        BindingNavigator3.BindingSource = Nothing
        Label5.Text = "0.00"
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        ExportToExcel(DataGridView2, "")
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        ExportToExcel(DataGridView3, "")
    End Sub

    Private Sub DataGridView2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.Click
        If DataGridView2.Rows.Count > 0 Then
            loadReceiptsPerCustomerAndPaymentMode(DateTimePicker1.Value, DateTimePicker2.Value, Integer.Parse(DataGridView2.CurrentRow.Cells(0).Value))
            totalamount = 0
            For Each row As DataGridViewRow In DataGridView3.Rows
                totalamount = totalamount + Decimal.Parse(row.Cells(2).Value)
            Next
            Label5.Text = getCurrency(totalamount)
        End If
    End Sub

    Private Sub ReceiptsReportForModeOfPayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class