Public Class RemoveDuplicateCustomerRecords

    Dim index As Integer
    Dim customerid As String
    Dim searchedcustomer As Boolean = False
    Dim employername As String
    Dim status As String

    Private Sub RemoveDuplicateCustomerRecords_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCustomers()
    End Sub

    Private Sub loadCustomers()
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select CustomerID,Count(CustomerID) from CustomerDetails group by CustomerID"

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
                DataGridView2.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadDuplicateCustomer(ByVal customerid As String)
        DataGridView4.DataSource = Nothing
        Dim qry As String = "select * from CustomerDetails where CustomerID = '" & customerid & "'"

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
                DataGridView4.DataSource = bs
                DataGridView4.Columns(DataGridView4.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView4.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadDuplicateCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadDuplicateCustomer()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadDuplicateCustomerNoZeros(ByVal customerid As String)
        DataGridView4.DataSource = Nothing
        Dim qry As String = "select * from CustomerDetails where CustomerID like '%" & customerid & "%'"

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
                DataGridView4.DataSource = bs
                DataGridView4.Columns(DataGridView4.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView4.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadDuplicateCustomerNoZeros()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadDuplicateCustomerNoZeros()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        For Each row As DataGridViewRow In DataGridView2.Rows
            If Integer.Parse(row.Cells(1).Value) > 1 Then
                DataGridView1.Rows.Add()
                index = DataGridView1.Rows.Count - 1
                DataGridView1.Rows(index).Cells(0).Value = row.Cells(0).Value
                DataGridView1.Rows(index).Cells(1).Value = row.Cells(1).Value
            End If
        Next
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        If DataGridView1.Rows.Count > 0 Then
            DataGridView3.Rows.Clear()
            loadDuplicateCustomer(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value)
            For Each row As DataGridViewRow In DataGridView4.Rows
                DataGridView3.Rows.Add()
                index = DataGridView3.Rows.Count - 1
                DataGridView3.Rows(index).Cells(0).Value = row.Cells(0).Value
                DataGridView3.Rows(index).Cells(1).Value = row.Cells(1).Value
                DataGridView3.Rows(index).Cells(2).Value = row.Cells(2).Value
                DataGridView3.Rows(index).Cells(3).Value = getEmployerNameByCustomerID(row.Cells(1).Value)
                DataGridView3.Rows(index).Cells(4).Value = row.Cells(4).Value
                searchedcustomer = False
            Next
        End If
    End Sub

    Private Sub DataGridView3_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.DoubleClick
        If DataGridView3.Rows.Count > 1 Then
            If searchedcustomer = True Then
                If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
                    customerid = DataGridView3.Rows(DataGridView3.CurrentRow.Index).Cells(1).Value
                    deleteCustomer(Integer.Parse(DataGridView3.Rows(DataGridView3.CurrentRow.Index).Cells(0).Value))

                    DataGridView3.Rows.Clear()
                    loadDuplicateCustomerNoZeros(customerid)
                    For Each row As DataGridViewRow In DataGridView4.Rows
                        DataGridView3.Rows.Add()
                        index = DataGridView3.Rows.Count - 1
                        DataGridView3.Rows(index).Cells(0).Value = row.Cells(0).Value
                        DataGridView3.Rows(index).Cells(1).Value = row.Cells(1).Value
                        DataGridView3.Rows(index).Cells(2).Value = row.Cells(2).Value
                        DataGridView3.Rows(index).Cells(3).Value = getEmployerNameByCustomerID(row.Cells(1).Value)
                        DataGridView3.Rows(index).Cells(4).Value = row.Cells(4).Value
                        searchedcustomer = True
                    Next
                End If
            ElseIf searchedcustomer = False Then
                If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
                    customerid = DataGridView3.Rows(DataGridView3.CurrentRow.Index).Cells(1).Value
                    deleteCustomer(Integer.Parse(DataGridView3.Rows(DataGridView3.CurrentRow.Index).Cells(0).Value))
                    loadCustomers()

                    DataGridView1.Rows.Clear()
                    For Each row As DataGridViewRow In DataGridView2.Rows
                        If Integer.Parse(row.Cells(1).Value) > 1 Then
                            DataGridView1.Rows.Add()
                            index = DataGridView1.Rows.Count - 1
                            DataGridView1.Rows(index).Cells(0).Value = row.Cells(0).Value
                            DataGridView1.Rows(index).Cells(1).Value = row.Cells(1).Value
                        End If
                    Next

                    DataGridView3.Rows.Clear()
                    loadDuplicateCustomer(customerid)
                    For Each row As DataGridViewRow In DataGridView4.Rows
                        DataGridView3.Rows.Add()
                        index = DataGridView3.Rows.Count - 1
                        DataGridView3.Rows(index).Cells(0).Value = row.Cells(0).Value
                        DataGridView3.Rows(index).Cells(1).Value = row.Cells(1).Value
                        DataGridView3.Rows(index).Cells(2).Value = row.Cells(2).Value
                        DataGridView3.Rows(index).Cells(3).Value = getEmployerNameByCustomerID(row.Cells(1).Value)
                        DataGridView3.Rows(index).Cells(4).Value = row.Cells(4).Value
                        searchedcustomer = False
                    Next
                End If
            End If
        Else
            MessageBox.Show("Cannot delete customer because the record is not duplicated", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView3.Rows.Clear()
        loadDuplicateCustomerNoZeros(TextBox1.Text.Trim)
        For Each row As DataGridViewRow In DataGridView4.Rows
            DataGridView3.Rows.Add()
            index = DataGridView3.Rows.Count - 1
            DataGridView3.Rows(index).Cells(0).Value = row.Cells(0).Value
            DataGridView3.Rows(index).Cells(1).Value = row.Cells(1).Value
            DataGridView3.Rows(index).Cells(2).Value = row.Cells(2).Value
            DataGridView3.Rows(index).Cells(3).Value = getEmployerNameByCustomerID(row.Cells(1).Value)
            DataGridView3.Rows(index).Cells(4).Value = row.Cells(4).Value
            searchedcustomer = True
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DataGridView1.Rows.Count > 0 Then
            For Each row1 As DataGridViewRow In DataGridView1.Rows
                DataGridView3.Rows.Clear()
                loadDuplicateCustomer(row1.Cells(0).Value)
                status = "Remove"
                For Each row4 As DataGridViewRow In DataGridView4.Rows
                    DataGridView3.Rows.Add()
                    index = DataGridView3.Rows.Count - 1
                    DataGridView3.Rows(index).Cells(0).Value = row4.Cells(0).Value
                    DataGridView3.Rows(index).Cells(1).Value = row4.Cells(1).Value
                    DataGridView3.Rows(index).Cells(2).Value = row4.Cells(2).Value
                    DataGridView3.Rows(index).Cells(3).Value = getEmployerNameByCustomerID(row4.Cells(1).Value)
                    DataGridView3.Rows(index).Cells(4).Value = row4.Cells(4).Value

                    If row4.Index = 0 Then
                        employername = getEmployerNameByCustomerID(row4.Cells(1).Value)
                        DataGridView3.Rows(index).Cells(5).Value = "Keep"
                    Else
                        If getEmployerNameByCustomerID(row4.Cells(1).Value) = employername Then
                            DataGridView3.Rows(index).Cells(5).Value = "Remove"
                        Else
                            status = "Keep"
                            DataGridView3.Rows(index).Cells(5).Value = "Keep"
                        End If
                    End If
                Next
                If status = "Remove" Then
                    For Each row3 As DataGridViewRow In DataGridView3.Rows
                        If row3.Cells(5).Value = "Remove" Then
                            geleteCustomerDetails(Integer.Parse(row3.Cells(0).Value), row3.Cells(1).Value)
                        End If
                    Next
                End If
            Next
            MessageBox.Show("Done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class