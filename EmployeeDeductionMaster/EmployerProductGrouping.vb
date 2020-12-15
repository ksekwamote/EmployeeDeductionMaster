Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class EmployerProductGrouping
    Private Sub GroupProductGrouping_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadEmployerProductGrouping()
    End Sub

    Private Sub loadEmployerProductGrouping()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select EmployerProductGrouping.AutoID,EmployerProductGrouping.ProductID,Products,ProductGroupingID,ProductGrouping.ProductGroupingName,EmployerProductGrouping.EmployerID,EmployerName,EmployerProductGrouping.Status from EmployerProductGrouping inner Join Products on EmployerProductGrouping.ProductID = Products.AutoID inner join ProductGrouping on EmployerProductGrouping.ProductGroupingID = ProductGrouping.AutoID inner  join Employers on EmployerProductGrouping.EmployerID = Employers.AutoID"

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
            MessageBox.Show(ex.Message, "loadEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim vrl As New ProductGroupingList
        vrl.MdiParent = mform
        vrl.whichform = "Employer Product Grouping"
        vrl.gpg = Me
        vrl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "Employer Product Grouping"
        vrl.gpg = Me
        vrl.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim vrl As New ProductsList
        vrl.MdiParent = mform
        vrl.whichform = "Product List"
        vrl.gpg = Me
        vrl.Show()
    End Sub

    Public Sub clearText()
        Button2.Enabled = False
        Button1.Enabled = True
        tbgroupingid.Text = ""
        tbproductname.Text = ""
        tbproductid.Text = ""
        tbemployername.Text = ""
        tbemployerid.Text = ""
        tbproductgroupingname.Text = ""
        tbproductgroupingid.Text = ""
        tbStatus.Text = ""
    End Sub

    Public Sub saveEmployerProductGrouping(ByVal ProductID As Integer, ByVal ProductGroupingID As Integer, ByVal EmployerID As String, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into EmployerProductGrouping values ('" & ProductID & "','" & ProductGroupingID & "','" & EmployerID & "','" & Status & "')"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "saveEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateEmployerProductGrouping(ByVal AutoID As String, ByVal ProductID As Integer, ByVal ProductGroupingID As Integer, ByVal EmployerID As String, ByVal Status As String)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update EmployerProductGrouping set ProductID = '" & ProductID & "',ProductGroupingID = '" & ProductGroupingID & "',EmployerID = '" & EmployerID & "',Status = '" & Status & "' where AutoID = '" & AutoID & "'"
            Dim cmdinsert As New SqlClient.SqlCommand(qryinsert, con)
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            transactioninsert = con.BeginTransaction()
            cmdinsert.Transaction = transactioninsert

            If cmdinsert.ExecuteNonQuery() Then
                transactioninsert.Commit()
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "updateEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateEmployerProductGrouping()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbgroupingid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
            tbproductid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
            tbproductname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
            tbproductgroupingid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
            tbproductgroupingname.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
            tbemployerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
            tbemployername.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
            tbStatus.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim vrl As New StatusList
        vrl.MdiParent = mform
        vrl.whichform = "Status List"
        vrl.gpg = Me
        vrl.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If tbproductid.Text.Trim = "" Or tbproductgroupingid.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbStatus.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveEmployerProductGrouping(Integer.Parse(tbproductid.Text.Trim), Integer.Parse(tbproductgroupingid.Text.Trim), tbemployerid.Text.Trim, tbStatus.Text.Trim)
                    loadEmployerProductGrouping()
                    clearText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If tbproductid.Text.Trim = "" Or tbproductgroupingid.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbStatus.Text.Trim = "" Or tbgroupingid.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateEmployerProductGrouping(Integer.Parse(tbgroupingid.Text.Trim), Integer.Parse(tbproductid.Text.Trim), Integer.Parse(tbproductgroupingid.Text.Trim), tbemployerid.Text.Trim, tbStatus.Text.Trim)
                    loadEmployerProductGrouping()
                    clearText()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clearText()
    End Sub
End Class