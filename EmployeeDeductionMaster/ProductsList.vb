Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class ProductsList

    Public whichform As String
    Public cd As CompanyDetails
    Public gpg As EmployerProductGrouping
    Public ep As EmployerProducts
    Public theemployerid As Integer

    Private Sub loadProducts(ByVal which As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = ""

        If which = "CompanyDetails" Then
            qry = "select AutoID,Products from Products where Status = 'Active'"
        ElseIf which = "Product List" Then
            qry = "select AutoID,Products from Products where Status = 'Active'"
        ElseIf which = "Employer Products" Then
            qry = "select AutoID,Products from Products where Status = 'Active' and AutoID not in (select ProductID from EmployerProducts where EmployerID = '" & theemployerid & "')"
        End If
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
            MessageBox.Show(ex.Message, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "CompanyDetails" Then
                cd.tbproductid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cd.tbproductname.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "Product List" Then
                gpg.tbproductid.Text = DataGridView1.CurrentRow.Cells(0).Value
                gpg.tbproductname.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "Employer Products" Then
                ep.tbproductid.Text = DataGridView1.CurrentRow.Cells(0).Value
                ep.tbproductname.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub ProductsList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadProducts(whichform)
    End Sub
End Class