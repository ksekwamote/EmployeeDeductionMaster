Imports DgvFilterPopup
Public Class TerminationsAndAmendmentsCorrection

    Private Sub TerminationsAndAmendmentsCorrection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim vrl As New EmployerList
        vrl.MdiParent = mform
        vrl.whichform = "Terminations"
        vrl.ter = Me
        vrl.Show()
    End Sub


    Private Sub loadAmendments(ByVal Formonth As Date, ByVal EmployerID As Integer, ByVal SavedTime As Date)
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select AutoID,CustomerID,Amount,ForMonth,SavedTime,ProductID as SavedProductID,'' as ProductName,'' as ExpectedProductID from TerminationsAndAmendments where ForMonth = '" & Formonth & "' and EmployerID = '" & EmployerID & "' and SavedTime >= '" & SavedTime & "' order by SavedTime"

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
                    ds.Tables(0).Rows(count).Item(6) = getStringValue("select Products from EmployerProducts inner join Products on EmployerProducts.ProductID = Products.AutoID where EmployerProducts.Status = 'Active' and Products.Status = 'Active' and EmployerID = '" & EmployerID & "' and EmployerProducts.AutoID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(5).ToString) & "'", con)
                    ds.Tables(0).Rows(count).Item(7) = getStringValue("select EmployerProducts.AutoID from EmployerProducts inner join Products on EmployerProducts.ProductID = Products.AutoID where EmployerProducts.Status = 'Active' and Products.Status = 'Active' and EmployerID = '" & EmployerID & "' and Products = '" & ds.Tables(0).Rows(count).Item(6).ToString & "'", con)

                    If ds.Tables(0).Rows(count).Item(6) = "" Then
                        ds.Tables(0).Rows(count).Item(6) = getStringValue("select Products from EmployerProducts inner join Products on EmployerProducts.ProductID = Products.AutoID where EmployerProducts.Status = 'Active' and Products.Status = 'Active' and EmployerID = '" & EmployerID & "' and EmployerProducts.ProductID = '" & Integer.Parse(ds.Tables(0).Rows(count).Item(5).ToString) & "'", con)
                        ds.Tables(0).Rows(count).Item(7) = Integer.Parse(ds.Tables(0).Rows(count).Item(5).ToString)
                    End If

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
            MessageBox.Show(ex.Message, "loadAmendments()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadAmendments()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Value = getLastDateOfMonth(DateTimePicker1.Value)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If tbemployerid.Text.Trim = "" Then
            MessageBox.Show("Select Employer First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            loadAmendments(DateTimePicker1.Value, Integer.Parse(tbemployerid.Text.Trim), DateTimePicker2.Value)
            MessageBox.Show("Done Loading", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class