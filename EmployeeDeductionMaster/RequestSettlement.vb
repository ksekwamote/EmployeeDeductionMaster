Public Class RequestSettlement

    Public sb As New SettlementBalance
    Dim index As Integer
    Dim batchname As String

    Private Sub RequestSettlement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each row As DataGridViewRow In sb.DataGridView1.Rows
            If Boolean.Parse(row.Cells(5).Value) = True Then
                DataGridView1.Rows.Add()
                index = DataGridView1.Rows.Count - 1
                DataGridView1.Rows(index).Cells(0).Value = row.Cells(0).Value
                DataGridView1.Rows(index).Cells(1).Value = row.Cells(1).Value
                DataGridView1.Rows(index).Cells(2).Value = row.Cells(3).Value
                DataGridView1.Rows(index).Cells(3).Value = row.Cells(4).Value
                TextBox1.Text = Decimal.Parse(TextBox1.Text.Trim) + Decimal.Parse(row.Cells(3).Value)
                tbtotalinstalment.Text = Decimal.Parse(tbtotalinstalment.Text.Trim) + Decimal.Parse(row.Cells(4).Value)
            End If
            TextBox2.Text = Decimal.Parse(TextBox2.Text.Trim) + Decimal.Parse(row.Cells(3).Value)
        Next
        TextBox3.Text = Decimal.Parse(TextBox2.Text.Trim) - Decimal.Parse(TextBox1.Text.Trim)

        batchname = "Grouping:" & Date.Now
        For Each row As DataGridViewRow In DataGridView1.Rows

            If checkIfProductHasAGroup(Integer.Parse(row.Cells(0).Value), getEmployerIDByCustomerID(tbcustomerid.Text.Trim)) = True Then
                If checkIfCustomerGroupIsForMerging(tbcustomerid.Text.Trim, conFPM) = True Then
                    saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(2).Value), Decimal.Parse(row.Cells(3).Value), getProductGroupingName(Integer.Parse(row.Cells(0).Value), getEmployerIDByCustomerID(tbcustomerid.Text.Trim)), batchname)
                ElseIf checkIfCustomerGroupIsForMerging(tbcustomerid.Text.Trim, conFPM) = False Then
                    saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(2).Value), Decimal.Parse(row.Cells(3).Value), row.Cells(1).Value, batchname)
                End If
            Else
                saveGroupingBreakdown(Integer.Parse(row.Cells(0).Value), row.Cells(1).Value, Decimal.Parse(row.Cells(2).Value), Decimal.Parse(row.Cells(3).Value), row.Cells(1).Value, batchname)
            End If
        Next
        loadProductSummary(batchname)
        deleteGroupingBreakdown(batchname)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            sb.rs = Me
            sb.totalbalance = ": " + getCurrency(Decimal.Parse(TextBox2.Text.Trim))
            sb.totalinstalments = ": " + getCurrency(Decimal.Parse(tbtotalinstalment.Text.Trim)) + " monthly instalment(s)"
            sb.generateSettlementReport()
        End If
    End Sub

    Private Sub loadProductSummary(ByVal BatchName As String)
        DataGridView2.DataSource = Nothing
        Dim qry As String = "select GroupingName,SUM(Balance) as Balance,SUM(Instalment) as Instalment from GroupingBreakdown where BatchName = '" & BatchName & "'  group by GroupingName"

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
                dt = ds.Tables(0)
                Dim bs As New BindingSource
                bs.DataSource = dt.DefaultView
                DataGridView2.DataSource = bs
                DataGridView2.Columns(DataGridView2.ColumnCount - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                DataGridView2.DataSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "loadProductSummary()", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show(ex.StackTrace, "loadProductSummary()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            con.Close()
        End Try
    End Sub

End Class