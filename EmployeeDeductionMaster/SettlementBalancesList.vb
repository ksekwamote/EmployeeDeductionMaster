Imports DgvFilterPopup

Public Class SettlementBalancesList

    Dim switchcount As Integer = 0
    Dim dsbatches As New DataSet

    Private Sub SettlementBalancesList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadSettlements()
    End Sub

    Private Sub loadSettlements()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select CustomerID,SUM(Amount) as Settlement,BatchName from SettlementBalances group by CustomerID,BatchName"

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
            MessageBox.Show(ex.Message, "loadSettlements()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSettlements()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadSettlementsByID(ByVal customerid As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select CustomerID,SUM(Amount) as Settlement,BatchName from SettlementBalances where CustomerID = '" & customerid & "' group by CustomerID,BatchName"

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
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadSettlementsByID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSettlementsByID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadSettlementsBreakdown(ByVal customerid As String, ByVal BatchName As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select CustomerID,Products,Amount,RequestedDate,CapturedBy,CapturedTime,BatchName from SettlementBalances inner join Products on SettlementBalances.ProductID = Products.AutoID where CustomerID = '" & customerid & "' and BatchName = '" & BatchName & "' order by SettlementBalances.AutoID Desc"

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
                DataGridView2.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
                BindingNavigator2.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadSettlementsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadSettlementsBreakdown()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadSettlementsByID(tbcustomerid.Text.Trim)
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
        'colorDisplay()
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            loadSettlementsBreakdown(DataGridView1.CurrentRow.Cells(0).Value, DataGridView1.CurrentRow.Cells(2).Value)
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            MessageBox.Show(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value & "  : " & getCurrency(getTotalBalance(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value)), "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub colorDisplay()
        switchcount = 0
        dsbatches = getBatchesForCustomer(tbcustomerid.Text.Trim)
        For count As Integer = 0 To dsbatches.Tables(0).Rows.Count - 1
            switchcount = switchcount + 1
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells(6).Value.ToString = dsbatches.Tables(0).Rows(count).Item(0).ToString Then
                    If switchcount = 1 Then
                        row.DefaultCellStyle.BackColor = Color.LightGray
                    ElseIf switchcount = 2 Then
                        row.DefaultCellStyle.BackColor = Color.WhiteSmoke
                    End If
                End If
            Next
            If switchcount = 2 Then
                switchcount = 0
            End If
        Next
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        loadSettlements()
        DataGridView2.DataSource = Nothing
        BindingNavigator2.BindingSource = Nothing
    End Sub

End Class