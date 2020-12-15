Imports DgvFilterPopup

Public Class CustomersPaidForList

    Private Sub CustomersPaidForList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadPaidList()
    End Sub

    Private Sub loadPaidList()
        DataGridView1.DataSource = Nothing
        Dim qry As String = "select PaidForCustomers.AutoID,CustomerID,ProductID,Products,Amount,PaidByCustomerID from PaidForCustomers inner join Products on PaidForCustomers.ProductID = Products.AutoID where PaidForCustomers.STatus = 'Active' order by CustomerID ASC,ProductID ASC"

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
                filterManager = New DgvFilterManager(DataGridView1)
                For Each col As DataGridViewColumn In DataGridView1.Columns
                    col.ReadOnly = True
                Next
            Else
                DataGridView1.DataSource = Nothing
                BindingNavigator1.BindingSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadPaidList()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadPaidList()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ExportToExcel(DataGridView1, "")
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim vrl As New Lodging
        vrl.MdiParent = mform
        vrl.Text = vrl.Text & " : Employer Name = " & employername
        vrl.Show()
        Me.Close()
    End Sub
End Class