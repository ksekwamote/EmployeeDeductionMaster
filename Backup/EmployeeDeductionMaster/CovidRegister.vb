Public Class CovidRegister

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tbcustomerid.Text.Trim = "" Or tbfirstname.Text.Trim = "" Or tbsurname.Text.Trim = "" Or tbhomeaddress.Text.Trim = "" Or tbcontacts.Text.Trim = "" Or cbgender.Text.Trim = "" Or tbtemperature.Text.Trim = "" Or tboffice.Text.Trim = "" Then
            MessageBox.Show("Enter All The Details First", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            saveCovidDetails(tbcustomerid.Text.Trim, tbfirstname.Text.Trim, tbsurname.Text.Trim, tbhomeaddress.Text.Trim, tbcontacts.Text.Trim, cbgender.Text.Trim, Decimal.Parse(tbtemperature.Text.Trim), dtpdate.Value, DateTime.Parse(dtptime.Value), Date.Now, Date.Now, logedinname, tboffice.Text.Trim)
            MessageBox.Show("Save Details", "Message()", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cleraText()
        End If
    End Sub

    Public Sub cleraText()
        tbcustomerid.Text = ""
        tbfirstname.Text = ""
        tbsurname.Text = ""
        tbhomeaddress.Text = ""
        tbcontacts.Text = ""
        cbgender.Text = ""
        tbtemperature.Text = ""
        tboffice.Text = ""
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim vrl As New UploadCovid19Register
        vrl.MdiParent = mform
        vrl.Show()
    End Sub
End Class