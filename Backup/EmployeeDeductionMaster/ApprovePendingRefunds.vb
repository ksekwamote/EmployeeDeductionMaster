'Imports DgvFilterPopup

Public Class ApprovePendingRefunds

    Private Sub loadNewChequeRefunds(ByVal PayBy As String)
        DataGridView1.DataSource = Nothing

        Dim qry As String = "select ProductsCheques.AutoID,CustomerID,ProductID,Products,Amount,CapturedBy,CapturedDate,ChequeNumber,ProductsCheques.Status,PayBy from ProductsCheques inner join Products on ProductsCheques.ProductID = Products.AutoID where ProductsCheques.Status = 'New' and PayBy = '" & PayBy & "'"
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
                DataGridView1.DataSource = bs
                DataGridView1.Columns(DataGridView1.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'filterManager = New DgvFilterManager(DataGridView1)
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            MessageBox.Show(ex.Message, "loadNewChequeRefunds()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure to continue ?") = DialogResult.Yes Then
                updateRefundStatus(Integer.Parse(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value), "Approved")
                wrapper.saveAction("Refund Approved", Date.Now, logedinname, DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value)
                MessageBox.Show("Refund Approved", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                loadNewChequeRefunds(tbpayby.Text.Trim)
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbpayby.Text.Trim = "" Then
            MessageBox.Show("Select PayBy First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            loadNewChequeRefunds(tbpayby.Text.Trim)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbpayby.Text = ComboBox1.Text
    End Sub

    Private Sub ApprovePendingRefunds_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class