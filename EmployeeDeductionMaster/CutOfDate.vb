Public Class CutOfDate

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        tbdate.Text = DateTimePicker1.Value.Day
    End Sub

    Private Sub CutOfDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If checkIfMonthHasCutOfDate(getFirstDateOfMonth(Date.Now), getLastDateOfMonth(Date.Now)) = True Then
            DateTimePicker1.Value = getCutOffDate()
            If decision("Current Cut Off Date = " & getCutOffDate() & ", Do you want to change it ?") = DialogResult.Yes Then
                Button1.Enabled = True
            Else
                Button1.Enabled = False
            End If
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbdate.Text.Trim = "" Then
            MessageBox.Show("Select the cut off date first", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If decision("Are you sure to proceed ?") = DialogResult.Yes Then
                updateCutOffDateToClosed()
                updateCustOffDateInCompanyDetails(Integer.Parse(tbdate.Text.Trim))
                saveCutOffDate(Integer.Parse(tbdate.Text.Trim), DateTimePicker1.Value, logedinname, Date.Now, "Active")
                MessageBox.Show("Cut Off Date Confirmed", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Close()
            End If
        End If
    End Sub
End Class