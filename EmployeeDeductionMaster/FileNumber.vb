Public Class FileNumber

    Public fd As FIllingDocuments
    Private Sub FileNumber_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbfilenumber.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If tbfilenumber.Text.Trim = "" Then

        Else
            fd.saveFileNumber(tbfilenumber.Text.Trim)
            Close()
        End If
    End Sub
End Class