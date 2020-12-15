Public Class RequestSettlement

    Public sb As New SettlementBalance
    Dim index As Integer

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
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
            sb.totalbalance = ": " + getCurrency(Decimal.Parse(TextBox2.Text.Trim))
            sb.totalinstalments = ": " + getCurrency(Decimal.Parse(tbtotalinstalment.Text.Trim)) + " monthly instalment(s)"
            sb.generateSettlementReport()
            Me.Close()
        End If
    End Sub

End Class