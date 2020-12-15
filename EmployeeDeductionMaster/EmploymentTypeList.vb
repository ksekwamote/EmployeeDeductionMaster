Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class EmploymentTypeList

    Public whichform As String
    Public cust As Customer
    Public autoid As String
    Public etype As EmploymentType

    Private Sub loadEmploymentType(ByVal tform As String, ByVal con As SqlClient.SqlConnection)
        DataGridView1.DataSource = Nothing
        Dim qry As String

        If tform = "Customer" Then
            qry = "select * from EmploymentType order by EmploymentTypeName ASC"
        ElseIf tform = "MLMA" Then
            qry = "select * from EmploymentType order by EmploymentTypeName ASC"
        ElseIf tform = "MLME" Then
            qry = "select * from EmploymentType order by EmploymentTypeName ASC"
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
            MessageBox.Show(ex.Message, "loadEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmploymentType()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If whichform = "Customer" Then
                cust.tbemploymenttypeid.Text = DataGridView1.CurrentRow.Cells(0).Value
                cust.tbemploymenttype.Text = DataGridView1.CurrentRow.Cells(1).Value
                cust.MLMAEmpType = DataGridView1.CurrentRow.Cells(3).Value
                cust.MLMEEmpType = DataGridView1.CurrentRow.Cells(4).Value
                Close()
            ElseIf whichform = "MLMA" Then
                etype.tbmlmaid.Text = DataGridView1.CurrentRow.Cells(0).Value
                etype.tbmlma.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            ElseIf whichform = "MLME" Then
                etype.tbmlmeid.Text = DataGridView1.CurrentRow.Cells(0).Value
                etype.tbmlme.Text = DataGridView1.CurrentRow.Cells(1).Value
                Close()
            End If
        End If
    End Sub

    Private Sub EmploymentTypeList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If whichform = "Customer" Then
            loadEmploymentType(whichform, con)
        ElseIf whichform = "MLMA" Then
            loadEmploymentType(whichform, conMLMAdvances)
        ElseIf whichform = "MLME" Then
            loadEmploymentType(whichform, conMLMEducational)
        End If
    End Sub

End Class