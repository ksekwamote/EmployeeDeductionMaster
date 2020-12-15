Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class CustomerList

    Public whichform As String
    Public cpf As CustomersPaidFor

    Private Sub loadCustomerList()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select CustomerID,CustomerName from CustomerDetails"

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
                filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadCustomerList()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomerList()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "CustomersPaidFor" Then
                cpf.tbcustomerid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cpf.tbcustomername.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "CustomersPaidForPayer" Then
                cpf.tbpayerid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cpf.tbpayername.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub BankList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCustomerList()
    End Sub
End Class