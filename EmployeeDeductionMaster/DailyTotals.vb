Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class DailyTotals
    Dim ds As New DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            loadTotalForTheSpecifiedDate(tbpayby.Text.Trim, DataGridView1, BindingNavigator1, DateTimePicker1.Value, "Written")
            ds.Clear()
            ds = getProductsForTotals(DateTimePicker1.Value, tbpayby.Text.Trim, "Written")
            For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
                Dim col As New DataGridViewTextBoxColumn()
                col.HeaderText = ds.Tables(0).Rows(index).Item(1).ToString()
                col.Name = ds.Tables(0).Rows(index).Item(0).ToString()
                DataGridView1.Columns.Add(col)
                col.ReadOnly = True
            Next
            Button1.Enabled = False
            ComboBox1.Enabled = False
            DateTimePicker1.Enabled = False

            For Each row As DataGridViewRow In DataGridView1.Rows
                For Each col As DataGridViewColumn In DataGridView1.Columns
                    If col.Index > 7 Then
                        row.Cells(col.Index).Value = getTotalPerProductPerCustomerIDPerPaymentDate(DateTimePicker1.Value, tbpayby.Text.Trim, Integer.Parse(col.Name), row.Cells(0).Value, "Written")
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub loadTotalForTheSpecifiedDate(ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator, ByVal transactiondate As Date, ByVal Status As String)
        dgv.DataSource = Nothing

        Dim qry As String = "select ProductsCheques.CustomerID,CustomerName,SUM(Amount) as Total,PayBy,Bank,BranchCode,AccountNumber,PayeeName from ProductsCheques inner join CustomerDetails on ProductsCheques.CustomerID = CustomerDetails.CustomerID where PaymentDate = '" & transactiondate & "' and PayBy = '" & PayBy & "' and ChequeNumber not in ('') and Status = '" & Status & "' group by ProductsCheques.CustomerID,CustomerName,PayBy,Bank,BranchCode,AccountNumber,PayeeName"
        Dim ds As New DataSet
        Dim cmd As New SqlClient.SqlCommand(qry, con)
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
                bn.BindingSource = bs
                dgv.DataSource = bs
                dgv.Columns(dgv.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Dim filterManager As New DgvFilterManager
                filterManager = New DgvFilterManager(dgv)
            Else
                dgv.DataSource = Nothing
                bn.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadTotalForTheSpecifiedDate()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbpayby.Text = ComboBox1.Text
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If decision("Know that this list will be removed from the next export") = DialogResult.Yes Then
            If decision("Are you sure to proceed to export ?") = DialogResult.Yes Then
                ExportToExcel(DataGridView1, tbpayby.Text.Trim)
                updateDailyTotalsExport(DateTimePicker1.Value, tbpayby.Text.Trim, "Written")
                Close()
            End If
        End If
    End Sub

    Private Sub DailyTotals_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class