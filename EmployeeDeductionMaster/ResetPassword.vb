Public Class ResetPassword
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Fill all textboxes")

        Else
            If TextBox1.Text = TextBox2.Text Then
                MsgBox("Your password has been succesfuly reset, directing you to Login")

                Me.Close()

            Else
                MsgBox("Your passwords do not match , Try again")
            End If


        End If
    End Sub

    Private Sub ResetPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class