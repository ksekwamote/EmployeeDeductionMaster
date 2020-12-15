'Imports DgvFilterPopup
Public Class WrittenChequesProcessing

    Private Sub loadCheques(ByVal paymentdate As Date, ByVal Status As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select CustomerID,SUM(Amount) as TotalAmount,PaymentDate,ChequeNumber from ProductsCheques where PaymentDate = '" & paymentdate & "' and Status = '" & Status & "' and PayBy = '" & PayBy & "' group by CustomerID,PaymentDate,ChequeNumber"
        Dim ds As New DataSet
        Dim cmd As New SqlClient.SqlCommand(qry, con)
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
                bn.BindingSource = bs
                dgv.DataSource = bs
                dgv.Columns(dgv.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'filterManager = New DgvFilterManager(dgv)
            Else
                dgv.DataSource = Nothing
                bn.BindingSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub WrittenChequesProcessing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadCheques(DateTimePicker1.Value, "Requested", "Cheque", DataGridView2, BindingNavigator1)
        loadCheques(DateTimePicker1.Value, "Written", "Cheque", DataGridView3, BindingNavigator2)
        DateTimePicker1.Enabled = False
    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        If DataGridView2.Rows.Count > 0 Then
            If decision("Write Cheque For : " & getCustomerName(DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value) & " of ID " & DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value & " ?") = DialogResult.Yes Then
                Label1.Text = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value
                Label3.Text = getCustomerName(DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value)
                Label5.Text = Decimal.Parse(DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(1).Value)
                GroupBox2.Enabled = False
                GroupBox1.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbchequenumber.Text.Trim = "" Then
            MessageBox.Show("Enter the cheque number", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Have you written cheque for : " & Label3.Text & " of ID " & Label1.Text & " ?") = DialogResult.Yes Then
                updateSystemRefundStatusToWritten(Label1.Text, DateTimePicker1.Value, tbchequenumber.Text.Trim)
                wrapper.saveAction("Cheque Written", Date.Now, logedinname, Label1.Text)
                GroupBox2.Enabled = True
                GroupBox1.Enabled = False
                Label1.Text = "Customer ID"
                Label3.Text = "Customer Name"
                Label5.Text = "Amount"
                tbchequenumber.Text = "CHQ"
                Label6.Text = "CHQ"
                loadCheques(DateTimePicker1.Value, "Requested", "Cheque", DataGridView2, BindingNavigator1)
                loadCheques(DateTimePicker1.Value, "Written", "Cheque", DataGridView3, BindingNavigator2)
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        GroupBox2.Enabled = True
        GroupBox1.Enabled = False
        Label1.Text = "Customer ID"
        Label3.Text = "Customer Name"
        Label5.Text = "Amount"
        tbchequenumber.Text = "CHQ"
        Label6.Text = "CHQ"
    End Sub

    Private Sub tbchequenumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbchequenumber.TextChanged
        Label6.Text = tbchequenumber.Text.Trim
    End Sub
End Class