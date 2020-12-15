Public Class CancelCheques

    Private Sub CancelCheques_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadCheques(DateTimePicker1.Value, "Written", "Cheque", DataGridView1, BindingNavigator1)
        loadCancelledCheques(DateTimePicker1.Value, "Cheque", DataGridView2, BindingNavigator2)
        DateTimePicker1.Enabled = False
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If tbchequenumber.Text.Trim = "" Then
            MessageBox.Show("Enter cheque number", "loadCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Cancell Cheque Number " & tbchequenumber.Text.Trim & " ?") = DialogResult.Yes Then
                saveCancelledCheque("", tbchequenumber.Text.Trim, Date.Now, logedinname, DateTimePicker1.Value, "New", "Cheque")
                wrapper.saveAction("Cheque Cancelled (Without CustomerID)", Date.Now, logedinname, tbchequenumber.Text.Trim)
                MessageBox.Show("Cancelled succesfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                tbchequenumber.Text = "CHQ"
                loadCancelledCheques(DateTimePicker1.Value, "Cheque", DataGridView2, BindingNavigator2)
            End If
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If decision("Cancel Cheque Number " & DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value & " For : " & getCustomerName(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value) & " of ID " & DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value & " ?") = DialogResult.Yes Then
                If decision("Do you want to re-write the cheque with the same amount ?") = DialogResult.Yes Then
                    updateSystemRefundStatusToRequestedAfterCancellation(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value, Date.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value), DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value, "Requested")
                Else
                    updateSystemRefundStatusToRequestedAfterCancellation(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value, Date.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value), DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value, "Cancelled")
                End If
                saveCancelledCheque(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value, Date.Now, logedinname, DateTimePicker1.Value, "New", "Cheque")
                wrapper.saveAction("Cheque Cancelled", Date.Now, logedinname, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value)
                MessageBox.Show("Cancelled succesfully", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                loadCheques(DateTimePicker1.Value, "Written", "Cheque", DataGridView1, BindingNavigator1)
                loadCancelledCheques(DateTimePicker1.Value, "Cheque", DataGridView2, BindingNavigator2)
                DateTimePicker1.Enabled = False
                GroupBox1.Enabled = True
                GroupBox2.Enabled = True
            End If
        End If
    End Sub

    Private Sub loadCancelledCheques(ByVal paymentdate As Date, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select * from CancelledCheques where PaymentDate = '" & paymentdate & "' and PayBy = '" & PayBy & "'"
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
            MessageBox.Show(ex.Message, "loadCancelledCheques()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

End Class