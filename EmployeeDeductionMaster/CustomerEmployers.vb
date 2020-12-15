Imports DgvFilterPopup
Public Class CustomerEmployers
    Private Sub CustomerEmployers_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub loadCustomers()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoCustID,CustomerID,EmployerID,EmployerName,'' as RMOrange,'' as RMBemobile,'' as MLMAdvances,'' as MLMEducational from CustomerDetails inner join Employers on CustomerDetails.EmployerID = Employers.AutoID"

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
                ProgressBar1.Maximum = ds.Tables(0).Rows.Count - 1
                For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    ds.Tables(0).Rows(count).Item(4) = getStringValue("select Employer_Name from Employer where Code in (select EmployerCode from CustomerMaster where Account = '" & ds.Tables(0).Rows(count).Item(1).ToString & "')", conRMOrange)
                    ds.Tables(0).Rows(count).Item(5) = getStringValue("select Employer_Name from Employer where Code in (select EmployerCode from CustomerMaster where Account = '" & ds.Tables(0).Rows(count).Item(1).ToString & "')", conRMBeMobile)
                    ds.Tables(0).Rows(count).Item(6) = getStringValue("select Employer from Employer where AutoID in (select EmployerID from CustomerDetails where ID_Number = '" & ds.Tables(0).Rows(count).Item(1).ToString & "')", conMLMAdvances)
                    ds.Tables(0).Rows(count).Item(7) = getStringValue("select Employer from Employer where AutoID in (select EmployerID from CustomerDetails where ID_Number = '" & ds.Tables(0).Rows(count).Item(1).ToString & "')", conMLMEducational)
                    ProgressBar1.Value = count
                Next
                ProgressBar1.Value = 0
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
            MessageBox.Show(ex.Message, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        loadCustomers()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        loadCustomerForEmployers()
    End Sub

    Private Sub loadCustomerForEmployers()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoCustID,CustomerID,EmployerID,EmployerName,RMOEmployer as RMOEmployerID,RMBEmployer as RMBEMployerID,MLMAEmployer as MLMAEmployerID,'' as MLMAEmployerName,MLMEEmployer as MLMEEMployerID,'' as MLMEEmployerName from CustomerDetails inner join Employers on CustomerDetails.EmployerID = Employers.AutoID"

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
                ProgressBar1.Maximum = ds.Tables(0).Rows.Count - 1
                For count As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    ds.Tables(0).Rows(count).Item(7) = getStringValue("select Employer from Employer where AutoID = '" & ds.Tables(0).Rows(count).Item(6).ToString & "'", conMLMAdvances)
                    ds.Tables(0).Rows(count).Item(9) = getStringValue("select Employer from Employer where AutoID = '" & ds.Tables(0).Rows(count).Item(8).ToString & "'", conMLMEducational)
                    ProgressBar1.Value = count
                Next
                ProgressBar1.Value = 0
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
            MessageBox.Show(ex.Message, "loadCustomerForEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomerForEmployers()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.Rows.Count > 0 Then
            ProgressBar1.Maximum = DataGridView1.Rows.Count - 1
            For Each row As DataGridViewRow In DataGridView1.Rows
                If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & row.Cells(1).Value & "'", conRMOrange) = True Then
                    doAction("update CustomerMaster set EmployerCode = '" & row.Cells(4).Value & "' where Account = '" & row.Cells(1).Value & "'", conRMOrange)
                End If
                If getValueBooleanOtherSystems("select * from CustomerMaster where Account = '" & row.Cells(1).Value & "'", conRMBeMobile) = True Then
                    doAction("update CustomerMaster set EmployerCode = '" & row.Cells(5).Value & "' where Account = '" & row.Cells(1).Value & "'", conRMBeMobile)
                End If
                If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & row.Cells(1).Value & "'", conMLMAdvances) = True Then
                    doAction("update CustomerDetails set EmployerID = '" & Integer.Parse(row.Cells(6).Value) & "',Name_Of_Employer = '" & row.Cells(7).Value & "' where ID_Number = '" & row.Cells(1).Value & "'", conMLMAdvances)
                End If
                If getValueBooleanOtherSystems("select * from CustomerDetails where ID_Number = '" & row.Cells(1).Value & "'", conMLMEducational) = True Then
                    doAction("update CustomerDetails set EmployerID = '" & Integer.Parse(row.Cells(8).Value) & "',Name_Of_Employer = '" & row.Cells(9).Value & "' where ID_Number = '" & row.Cells(1).Value & "'", conMLMEducational)
                End If
                ProgressBar1.Value = row.Index
            Next
            ProgressBar1.Value = 0
            MessageBox.Show("Done", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class