Imports System.Net.Mail

Public Class ForgotPassword

    Public resetPass As New ResetPassword
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = Methods.verificationCode Then
            Me.Close()
            resetPass.Show()

        Else
            MsgBox("This is the wrong Verification Code")
        End If




    End Sub
    Dim rand As New Random()

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        Methods.verificationCode = CStr(rand.Next(1000, 9999))

        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("thitoithelpdesk@gmail.com", "2hi2oHoldings.")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("thitoithelpdesk@gmail.com")
            e_mail.To.Add(Methods.verificationEmail)
            e_mail.Subject = "Password Recovery Verification Code"
            e_mail.IsBodyHtml = True
            e_mail.Body = "<div align='center'>
                    <img  src='https://www.pngitem.com/pimgs/m/235-2350797_rbac-in-kubernetes-png-download-lock-icon-png.png' width='150' height='200' /> </div>
                    <div align='center'><div>Hello</div><br><br>    
                    <div>Your verification code is: <b>" + Methods.verificationCode + "<b></b></b></div><br><br>
                    <div>Please enter this code to verify your identity and reset your password </div><br>
                    <div>Should you encounter any problems, contact IT Support.</div><br><div>Kind Regards</div>"
            Smtp_Server.Send(e_mail)
            MsgBox("A verification code has been succesfully sent to your email address. ")



        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try

    End Sub

End Class