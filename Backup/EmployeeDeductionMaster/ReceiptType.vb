Public Class ReceiptType

    Public rc As ReceiptingCustomer
    Public rowindex As Integer

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tbreceipttype.Text = ComboBox1.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbreceipttype.Text.Trim = "" Then
            MessageBox.Show("Select the Receipt Type First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure you want to proceed?") = DialogResult.Yes Then
                rc.DataGridView1.Rows(rowindex).Cells(5).Value = tbreceipttype.Text
                Close()
            End If
        End If
    End Sub

End Class