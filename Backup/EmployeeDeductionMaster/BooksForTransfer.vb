Public Class BooksForTransfer

    Public trbb As TransferReceiptBetweenBooks

    Private Sub BooksForTransfer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadBooks()
    End Sub

    Private Sub loadBooks()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select * from CompanyBooks"

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
            MessageBox.Show(ex.Message, "loadBooks()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadBooks()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            If decision("Are you sure you want to transfer receit to " & DataGridView1.CurrentRow.Cells(1).Value & " receipt book ?") = DialogResult.Yes Then
                trbb.tbcompanynameto.Text = DataGridView1.CurrentRow.Cells(1).Value
                trbb.tbcompanybookidto.Text = Integer.Parse(DataGridView1.CurrentRow.Cells(0).Value)
                Close()
            End If
        End If
    End Sub

End Class