Public Class ViewActualDeductionForCustomer

    Dim ds As New DataSet
    Dim expected, actual As Decimal
    Dim amount As Decimal = 0

    Private Sub loadCustomerIDs(ByVal customerid As String)
        Dim qry As String = "select DISTINCT(CustomerID) from CustomerDetails where CustomerID = '" & customerid & "'"


        'Dim qry As String = "select DISTINCT(CustomerID) from ExpectedDeductions where CustomerID = '738521706'"
        DataGridView1.DataSource = Nothing
        BindingNavigator1.BindingSource = Nothing

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim dt As DataTable

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

            If ds.Tables(0).Rows.Count > 0 Then
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView1.DataSource = bs
                BindingNavigator1.BindingSource = bs

                DataGridView1.Columns(0).ReadOnly = True
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCustomerIDs()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadCustomerIDs()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If checkCustomerDetailsInDatabase(TextBox1.Text.Trim) = True Then
            Label5.Text = getEmployerNameByCustomerID(TextBox1.Text.Trim)
            dtpformonth.Value = getLastDateOfMonth(getLastActualDeductionDateByCustomerID(TextBox1.Text.Trim))
            dtppredate.Value = getLastDateOfMonth(addMonth(dtpformonth.Value, -1))

            Button1.Enabled = False
            loadCustomerIDs(TextBox1.Text.Trim)
            ds = getProducts()
            Dim col1 As New DataGridViewTextBoxColumn()
            col1.HeaderText = "Expected"
            col1.Name = "Expected"
            DataGridView1.Columns.Add(col1)

            Dim col2 As New DataGridViewTextBoxColumn()
            col2.HeaderText = "Actual"
            col2.Name = "Actual"
            DataGridView1.Columns.Add(col2)

            Dim col3 As New DataGridViewTextBoxColumn()
            col3.HeaderText = "Comments"
            col3.Name = "Comments"
            DataGridView1.Columns.Add(col3)

            For index As Integer = 0 To ds.Tables(0).Rows.Count - 1
                Dim col As New DataGridViewTextBoxColumn()
                col.HeaderText = ds.Tables(0).Rows(index).Item(0).ToString()
                col.Name = ds.Tables(0).Rows(index).Item(1).ToString()
                DataGridView1.Columns.Add(col)
            Next

            Dim col4 As New DataGridViewTextBoxColumn()
            col4.HeaderText = "CustomerName"
            col4.Name = "CustomerName"
            DataGridView1.Columns.Add(col4)

            ProgressBar1.Maximum = DataGridView1.Rows.Count
            For Each row As DataGridViewRow In DataGridView1.Rows
                row.Cells(1).Value = getExpectedDeductionPerCustomer(dtpformonth.Value, row.Cells(0).Value.ToString, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                row.Cells(2).Value = getActualDeduction(dtpformonth.Value, row.Cells(0).Value.ToString, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                expected = Decimal.Parse(row.Cells(1).Value)
                actual = Decimal.Parse(row.Cells(2).Value)
                If (expected = 0) And (expected < actual) Then
                    If row.Cells(3).Value = "" Then
                        row.Cells(3).Value = "Expected : 0"
                    Else
                        row.Cells(3).Value = row.Cells(3).Value & " ; " & "Expected : 0"
                    End If
                ElseIf (actual = 0) And (expected > actual) Then
                    If row.Cells(3).Value = "" Then
                        row.Cells(3).Value = "Defaulted"
                    Else
                        row.Cells(3).Value = row.Cells(3).Value & " ; " & "Defaulted"
                    End If
                Else
                    For Each col As DataGridViewColumn In DataGridView1.Columns
                        If col.Index = 0 Or col.Index = 1 Or col.Index = 2 Or col.Index = 3 Then

                        ElseIf col.Index = DataGridView1.Columns.Count - 1 Then
                            row.Cells(col.Index).Value = getCustomerName(row.Cells(0).Value)
                        Else
                            amount = getExpectedPerProduct(dtpformonth.Value, Integer.Parse(col.Name), row.Cells(0).Value.ToString, dtppredate.Value, getEmployerIDByCustomerID(TextBox1.Text.Trim))
                            If actual > 0 Then
                                If actual >= amount Then
                                    row.Cells(col.Index).Value = amount
                                    actual = actual - amount
                                ElseIf (actual < amount) And (actual > 0) Then
                                    row.Cells(col.Index).Value = actual
                                    If row.Cells(3).Value = "" Then
                                        row.Cells(3).Value = "Less Allocated : " & col.HeaderText & " : " & amount - actual
                                    Else
                                        row.Cells(3).Value = row.Cells(3).Value & " ; " & "Less Allocated : " & col.HeaderText & " : " & amount - actual
                                    End If
                                    actual = 0
                                End If
                            Else
                                row.Cells(col.Index).Value = 0.0
                            End If
                        End If
                    Next
                    If actual > 0 Then
                        If row.Cells(3).Value = "" Then
                            row.Cells(3).Value = "Excess : " & actual
                        Else
                            row.Cells(3).Value = row.Cells(3).Value & " ; " & "Excess : " & actual
                        End If
                    ElseIf actual = 0 Then
                        If (expected > Decimal.Parse(row.Cells(2).Value)) And (expected > 0) And (Decimal.Parse(row.Cells(2).Value) > 0) Then
                            If row.Cells(3).Value = "" Then
                                row.Cells(3).Value = "Deducted Less : " & expected - Decimal.Parse(row.Cells(2).Value)
                            Else
                                row.Cells(3).Value = row.Cells(3).Value & " ; " & "Deducted Less : " & expected - Decimal.Parse(row.Cells(2).Value)
                            End If
                        End If
                    End If
                    ProgressBar1.Value = row.Index
                End If
            Next
            ProgressBar1.Value = 0
        Else
            MessageBox.Show("Customer does not exists", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Function getProducts() As DataSet
        Dim qry As String = "select Products,AutoID from Products where Status = 'Active' order by AllocationOrder ASC"

        Dim cmd As New SqlClient.SqlCommand(qry, con)
        Dim ds As New DataSet
        Dim SQLAdapter As New System.Data.SqlClient.SqlDataAdapter(cmd)

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim ctr As Integer = 0
            SQLAdapter.Fill(ds, "Defaulters")

        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "getProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "getProducts()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
        Return ds
    End Function


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, employername & " " & dtpformonth.Value)
        wrapper.saveAction("Exported Duduction File ", Date.Now, logedinname, "")
    End Sub

    Private Sub Deduction_File_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Enabled = True
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim vrl As New ViewActualDeductionForCustomer
        vrl.MdiParent = mform
        vrl.Show()
        Me.Close()
    End Sub
End Class