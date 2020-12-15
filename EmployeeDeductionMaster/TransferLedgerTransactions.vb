Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class TransferLedgerTransactions



    Private Sub loadRechargesFrom(ByVal custid As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select EntryNo,CustomerDeals.Account,CustomerDeals.DealNo,RechargeDate,RechargeAmount,FreebieAmount as DealAmount,OldRechargeNumber as RechargeNumber from Recharges inner join CustomerDeals on Recharges.DealNo = CustomerDeals.DealNo where CustomerDeals.DealNo in (Select DealNo from CustomerDeals where Account = '" & custid & "')"

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
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadRechargesFrom()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadRechargesTo(ByVal custid As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select EntryNo,CustomerDeals.Account,CustomerDeals.DealNo,RechargeDate,RechargeAmount,FreebieAmount as DealAmount,OldRechargeNumber as RechargeNumber from Recharges inner join CustomerDeals on Recharges.DealNo = CustomerDeals.DealNo where CustomerDeals.DealNo in (Select DealNo from CustomerDeals where Account = '" & custid & "')"

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
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadRechargesTo()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim = "" Or TextBox2.Text.Trim = "" Then
            MessageBox.Show("Enter all the customer ids before ", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            loadRechargesFrom(TextBox1.Text.Trim)
            loadRechargesTo(TextBox2.Text.Trim)
            Button2.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            TextBox3.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString
            Button2.Enabled = True
        Else
            Button2.Enabled = False
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dealno As Integer
        dealno = getDefaultDealNo(TextBox2.Text.Trim)
        If decision("Are you sure to transfer " & DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString & " from Customer ID " & DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString & " to " & TextBox2.Text.Trim & ", dealno " & dealno) = DialogResult.Yes Then
            updateRechargesCustomer(Integer.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString), dealno)
            MessageBox.Show("Recharge transfered", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Button2.Enabled = False
            loadRechargesFrom(TextBox1.Text.Trim)
            loadRechargesTo(TextBox2.Text.Trim)
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class