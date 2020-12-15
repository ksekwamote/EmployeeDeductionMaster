Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class EmployerProducts

    Private Sub loadEmployerProducts()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,EmployerID,'' as EmployerName,ProductID,'' as ProductName,OrderNumber,Status from EmployerProducts where Status = 'Active' order by EmployerID,OrderNumber ASC"

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
                For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    ds.Tables(0).Rows(count).Item(2) = getEmployerNameByEmployerID(Integer.Parse(ds.Tables(0).Rows(count).Item(1).ToString))
                    ds.Tables(0).Rows(count).Item(4) = getProductName(Integer.Parse(ds.Tables(0).Rows(count).Item(3).ToString))
                Next
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
            MessageBox.Show(ex.Message, "loadEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadEmployerProductsByEmployerID(ByVal EmployerID As Integer)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,EmployerID,'' as EmployerName,ProductID,'' as ProductName,OrderNumber,Status from EmployerProducts where Status = 'Active' and EmployerID = '" & EmployerID & "' order by OrderNumber ASC"

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
                For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    ds.Tables(0).Rows(count).Item(2) = getEmployerNameByEmployerID(Integer.Parse(ds.Tables(0).Rows(count).Item(1).ToString))
                    ds.Tables(0).Rows(count).Item(4) = getProductName(Integer.Parse(ds.Tables(0).Rows(count).Item(3).ToString))
                Next
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
            MessageBox.Show(ex.Message, "loadEmployerProductsByEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadEmployerProductsByEmployerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "Employer Products"
        vrl.ep = Me
        vrl.Show()
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If tbemployerid.Text.Trim = "" Then
            MessageBox.Show("Select Employer First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim vrl As New ProductsList
            vrl.MdiParent = mform
            vrl.whichform = "Employer Products"
            vrl.ep = Me
            vrl.theemployerid = Integer.Parse(tbemployerid.Text.Trim)
            vrl.Show()
        End If
    End Sub

    Public Sub clearText(ByVal addsameemployer As Boolean)
        Button2.Enabled = False
        Button1.Enabled = True
        tbautoid.Text = ""
        tbproductname.Text = ""
        tbproductid.Text = ""
        If addsameemployer = True Then

        Else
            tbemployername.Text = ""
            tbemployerid.Text = ""
        End If
        tbStatus.Text = ""
        tbordernumber.Text = ""
    End Sub

    Public Sub saveEmployerProducts(ByVal EmployerID As Integer, ByVal ProductID As Integer, ByVal CapturedBy As String, ByVal CapturedDate As Date, ByVal LastModifiedBy As String, ByVal LastModifiedDate As Date, ByVal Status As String, ByVal OrderNumber As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try
            Dim qryinsert As String = "Insert into EmployerProducts values ('" & EmployerID & "','" & ProductID & "','" & CapturedBy & "','" & CapturedDate & "','" & LastModifiedBy & "','" & LastModifiedDate & "','" & Status & "','" & OrderNumber & "')"
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
            MessageBox.Show(ex.Message, "saveEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "saveEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Public Sub updateEmployerProducts(ByVal EmployerID As Integer, ByVal ProductID As Integer, ByVal LastModifiedBy As String, ByVal LastModifiedDate As Date, ByVal Status As String, ByVal OrderNumber As Integer, ByVal AutoID As Integer)
        Dim transactioninsert As SqlClient.SqlTransaction
        Try

            Dim qryinsert As String = "update EmployerProducts set EmployerID = '" & EmployerID & "',ProductID = '" & ProductID & "',LastModifiedBy = '" & LastModifiedBy & "',LastModifiedDate = '" & LastModifiedDate & "',Status = '" & Status & "',OrderNumber = '" & OrderNumber & "' where AutoID = '" & AutoID & "'"
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
            MessageBox.Show(ex.Message, "updateEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "updateEmployerProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim vrl As New StatusList
        vrl.MdiParent = mform
        vrl.whichform = "Employer Products"
        vrl.ep = Me
        vrl.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If tbproductid.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbStatus.Text.Trim = "" Or tbordernumber.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    saveEmployerProducts(Integer.Parse(tbemployerid.Text.Trim), Integer.Parse(tbproductid.Text.Trim), logedinname, Date.Now, logedinname, Date.Now, tbStatus.Text.Trim, Integer.Parse(tbordernumber.Text.Trim))
                    loadEmployerProductsByEmployerID(Integer.Parse(tbemployerid.Text.Trim))
                    If decision("Do you want to add another product under the Employer ?") = DialogResult.Yes Then
                        clearText(True)
                    Else
                        clearText(False)
                    End If
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
            If tbproductid.Text.Trim = "" Or tbemployerid.Text.Trim = "" Or tbStatus.Text.Trim = "" Or tbautoid.Text.Trim = "" Or tbordernumber.Text.Trim = "" Then
                MessageBox.Show("Enter all the details before saving", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                    updateEmployerProducts(Integer.Parse(tbemployerid.Text.Trim), Integer.Parse(tbproductid.Text.Trim), logedinname, Date.Now, tbStatus.Text.Trim, Integer.Parse(tbordernumber.Text.Trim), Integer.Parse(tbautoid.Text.Trim))
                    loadEmployerProductsByEmployerID(Integer.Parse(tbemployerid.Text.Trim))
                    clearText(False)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clearText(False)
        loadEmployerProducts()
    End Sub

    Private Sub EmployerProducts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadEmployerProducts()
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            tbautoid.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString()
            tbemployerid.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString()
            tbemployername.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString()
            tbproductid.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString()
            tbproductname.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString()
            tbordernumber.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString()
            tbStatus.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString()
            Button2.Enabled = True
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If tbemployerid.Text.Trim = "" Then

        Else
            loadEmployerProductsByEmployerID(Integer.Parse(tbemployerid.Text.Trim))
        End If
    End Sub
End Class