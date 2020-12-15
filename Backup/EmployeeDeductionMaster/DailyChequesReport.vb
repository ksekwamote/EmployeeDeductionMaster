Public Class DailyChequesReport

    Private Sub DailyChequesReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub loadCheques(ByVal sdate As Date, ByVal edate As Date, ByVal Status As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select ProductID,ChequeNumber,ProductsCheques.CustomerID,CustomerName,Amount as TotalCheque,'' as Others,'' as Loan,'' as Quick,PayeeName,AccountNumber,Bank,BranchCode,PaymentDate,CapturedDate from ProductsCheques inner join CustomerDetails on ProductsCheques.CustomerID = CustomerDetails.CustomerID where PaymentDate between '" & sdate & "' and '" & edate & "' and PayBy = '" & PayBy & "' and ChequeNumber not in ('')"
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

                For count As Integer = 0 To dt.Rows.Count - 1
                    If getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Loan" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Normal")
                        dt.Rows(count).Item(7) = 0.0
                    ElseIf getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Quick" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Quick")
                    Else
                        dt.Rows(count).Item(5) = Decimal.Parse(dt.Rows(count).Item(4))
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = 0
                    End If
                Next

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

    Private Sub loadChequesAll(ByVal Status As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select ProductID,ChequeNumber,ProductsCheques.CustomerID,CustomerName,Amount as TotalCheque,'' as Others,'' as Loan,'' as Quick,PayeeName,AccountNumber,Bank,BranchCode,PaymentDate,CapturedDate from ProductsCheques inner join CustomerDetails on ProductsCheques.CustomerID = CustomerDetails.CustomerID where PayBy = '" & PayBy & "' and ChequeNumber not in ('')"
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

                For count As Integer = 0 To dt.Rows.Count - 1
                    If getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Loan" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Normal")
                        dt.Rows(count).Item(7) = 0.0
                    ElseIf getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Quick" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Quick")
                    Else
                        dt.Rows(count).Item(5) = Decimal.Parse(dt.Rows(count).Item(4))
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = 0
                    End If
                Next

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
            MessageBox.Show(ex.Message, "loadChequesAll()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadChequesAllPerCustomerID(ByVal customerid As String, ByVal Status As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select ProductID,ChequeNumber,ProductsCheques.CustomerID,CustomerName,Amount as TotalCheque,'' as Others,'' as Loan,'' as Quick,PayeeName,AccountNumber,Bank,BranchCode,PaymentDate,CapturedDate from ProductsCheques inner join CustomerDetails on ProductsCheques.CustomerID = CustomerDetails.CustomerID where PayBy = '" & PayBy & "' and ChequeNumber not in ('') and ProductsCheques.CustomerID = '" & customerid & "'"
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

                For count As Integer = 0 To dt.Rows.Count - 1
                    If getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Loan" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Normal")
                        dt.Rows(count).Item(7) = 0.0
                    ElseIf getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Quick" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Quick")
                    Else
                        dt.Rows(count).Item(5) = Decimal.Parse(dt.Rows(count).Item(4))
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = 0
                    End If
                Next

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
            MessageBox.Show(ex.Message, "loadChequesAllPerCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadChequesAllPerChequeNumber(ByVal ChequeNumber As String, ByVal Status As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select ProductID,ChequeNumber,ProductsCheques.CustomerID,CustomerName,Amount as TotalCheque,'' as Others,'' as Loan,'' as Quick,PayeeName,AccountNumber,Bank,BranchCode,PaymentDate,CapturedDate from ProductsCheques inner join CustomerDetails on ProductsCheques.CustomerID = CustomerDetails.CustomerID where PayBy = '" & PayBy & "' and ChequeNumber not in ('') and ChequeNumber = '" & ChequeNumber & "'"
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

                Label6.Text = 0

                For count As Integer = 0 To dt.Rows.Count - 1
                    If getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Loan" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Normal")
                        dt.Rows(count).Item(7) = 0.0
                    ElseIf getProductSystem(Integer.Parse(dt.Rows(count).Item(0).ToString)) = "MLM Quick" Then
                        dt.Rows(count).Item(5) = 0.0
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = getPrincipal(dt.Rows(count).Item(2).ToString, Date.Parse(dt.Rows(count).Item(12).ToString), "Quick")
                    Else
                        dt.Rows(count).Item(5) = Decimal.Parse(dt.Rows(count).Item(4))
                        dt.Rows(count).Item(6) = 0.0
                        dt.Rows(count).Item(7) = 0
                    End If
                    Label6.Text = Decimal.Parse(Label6.Text.Trim) + Decimal.Parse(dt.Rows(count).Item(4))
                Next

                Label6.Text = "Value = P " + Label6.Text

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
            MessageBox.Show(ex.Message, "loadChequesAllPerChequeNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "Daily Cheques Report")
        ExportToExcel(DataGridView2, "Cancelled Cheques Report")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Label6.Text = "Value"
            loadCheques(DateTimePicker1.Value, DateTimePicker2.Value, "Written", tbpayby.Text.Trim, DataGridView1, BindingNavigator1)
            loadCancelledCheques(DateTimePicker1.Value, DateTimePicker2.Value, tbpayby.Text.Trim, DataGridView2, BindingNavigator2)
        End If
    End Sub

    Private Sub loadCancelledCheques(ByVal sdate As Date, ByVal edate As Date, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select * from CancelledCheques where PaymentDate between '" & sdate & "' and '" & edate & "' and PayBy = '" & PayBy & "'"
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

    Private Sub loadCancelledChequesAll(ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select * from CancelledCheques where PayBy = '" & PayBy & "'"
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
            MessageBox.Show(ex.Message, "loadCancelledChequesAll()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadCancelledChequesAllPerCustomerID(ByVal customerid As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select * from CancelledCheques where PayBy = '" & PayBy & "' and CustomerID = '" & customerid & "'"
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
            MessageBox.Show(ex.Message, "loadCancelledChequesAllPerCustomerID()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub loadCancelledChequesAllPerChequeNumber(ByVal ChequeNumber As String, ByVal PayBy As String, ByVal dgv As DataGridView, ByVal bn As BindingNavigator)
        dgv.DataSource = Nothing

        Dim qry As String = "select * from CancelledCheques where PayBy = '" & PayBy & "' and ChequeNumber = '" & ChequeNumber & "'"
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
            MessageBox.Show(ex.Message, "loadCancelledChequesAllPerChequeNumber()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbpayby.Text = ComboBox1.Text
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Label6.Text = "Value"
            loadChequesAll("Written", tbpayby.Text.Trim, DataGridView1, BindingNavigator1)
            loadCancelledChequesAll(tbpayby.Text.Trim, DataGridView2, BindingNavigator2)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbcustomerid.Text.Trim = "" Then
                MessageBox.Show("Enter Customer ID First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Label6.Text = "Value"
                loadChequesAllPerCustomerID(tbcustomerid.Text.Trim, "Written", tbpayby.Text.Trim, DataGridView1, BindingNavigator1)
                loadCancelledChequesAllPerCustomerID(tbcustomerid.Text.Trim, tbpayby.Text.Trim, DataGridView2, BindingNavigator2)
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If tbchequenumber.Text.Trim = "" Then
                MessageBox.Show("Enter Checque Number First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                loadChequesAllPerChequeNumber(tbchequenumber.Text.Trim, "Written", tbpayby.Text.Trim, DataGridView1, BindingNavigator1)
                loadCancelledChequesAllPerChequeNumber(tbchequenumber.Text.Trim, tbpayby.Text.Trim, DataGridView2, BindingNavigator2)
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class