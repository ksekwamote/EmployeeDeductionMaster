Imports System.Net.Mail

Public Class RecoverPassword

    Public verifyCode As New ForgotPassword
    Dim rand As New Random()


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Methods.verificationCode = CStr(rand.Next(1000, 9999))
        Methods.verificationEmail = TextBox1.Text

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
            e_mail.To.Add(TextBox1.Text)
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

            Me.Close()
            verifyCode.Show()

        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try

    End Sub
End Class