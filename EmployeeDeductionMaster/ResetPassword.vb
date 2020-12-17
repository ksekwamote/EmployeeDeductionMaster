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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox1.Text = "Hide"
            TextBox1.PasswordChar = ""
        Else

            CheckBox1.Text = "Show"
            TextBox1.PasswordChar = "*"

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            CheckBox2.Text = "Hide"
            TextBox2.PasswordChar = ""
        Else

            CheckBox2.Text = "Show"
            TextBox2.PasswordChar = "*"

        End If
    End Sub
End Class