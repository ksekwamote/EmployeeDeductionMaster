Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports DgvFilterPopup

Public Class ListOfCustomers

    Private Sub loadCustomers()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoCustID,CustomerID,CustomerName,EmployerID,EmployerName from CustomerDetails inner join Employers on CustomerDetails.EmployerID = Employers.AutoID"

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
            MessageBox.Show(ex.Message, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadCustomersByID(ByVal customerid As String)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoCustID,CustomerID,CustomerName,EmployerID,EmployerName from CustomerDetails inner join Employers on CustomerDetails.EmployerID = Employers.AutoID where CustomerID = '" & customerid & "'"

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
            MessageBox.Show(ex.Message, "loadCustomersByID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomersByID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ListOfCustomers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCustomers()
        loadEmployers()
        tbemployername.Text = ""
        tbemployerid.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If checkCustomerDetailsInDatabase(TextBox1.Text.Trim) = True Then
            loadCustomersByID(TextBox1.Text.Trim)
        Else
            MessageBox.Show("Customer does not exist in the database", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Public Sub loadEmployers()
        Dim qry As String = "select * from Employers"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Emp")
            If ds.Tables(0).Rows.Count > 0 Then
                Me.ComboBox1.DataSource = ds
                Me.ComboBox1.DisplayMember = "Emp.EmployerName"
                Me.ComboBox1.ValueMember = "Emp.AutoID"
                Me.ComboBox1.SelectedIndex = 0
            Else

            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbemployername.Text = ComboBox1.Text
        tbemployerid.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        tbcustomerid.Text = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbcustomerid.Text.Trim = "" Then
            MessageBox.Show("Select the customer first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf tbemployername.Text.Trim = "" Then
            MessageBox.Show("Select the employer first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure you want to proceed ?") = DialogResult.Yes Then
                updateCustomerEmployer(Integer.Parse(tbemployerid.Text.Trim), tbcustomerid.Text.Trim)
                MessageBox.Show("Done updating", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                tbcustomerid.Text = ""

                loadCustomers()
                tbemployername.Text = ""
                tbemployerid.Text = ""

            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If decision("Are you sure you want to proceed ?") = DialogResult.Yes Then
            Dim gender As String
            For Each row As DataGridViewRow In DataGridView1.Rows
                gender = ""
                gender = getFPMCustomerGender(row.Cells(1).Value, conFPM)
                If gender = "" Then
                    gender = getMLMCustomerGender(row.Cells(1).Value, conMLMAdvances)
                    If gender = "" Then
                        gender = getRMCustomerGender(row.Cells(1).Value, conRMOrange)
                        If gender = "" Then
                            gender = getRMCustomerGender(row.Cells(1).Value, conRMBeMobile)
                        End If
                    End If
                End If

                If gender = "" Then

                Else
                    updateGender(gender, Integer.Parse(row.Cells(0).Value))
                End If
            Next
            MessageBox.Show("Done updating Gender", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class