Imports System.Net
Imports System.Net.Mail
Imports System.Configuration
Imports SendGrid

Module SendGridVbExample

    Sub Main()

        Dim sendGridUserName = ConfigurationManager.AppSettings("SENDGRID_USERNAME")
        Dim sendGridPassword = ConfigurationManager.AppSettings("SENDGRID_PASSWORD")
        Dim tos = ConfigurationManager.AppSettings("TOS").Split(",")
        Dim from = ConfigurationManager.AppSettings("FROM")

        Dim smtpapi = New Smtpapi.Header()
        smtpapi.SetTo(tos)
        smtpapi.AddSubstitution("fullname", {"田中 太郎", "佐藤 次郎", "鈴木 三郎"})
        smtpapi.AddSubstitution("familyname", {"田中", "佐藤", "鈴木"})
        smtpapi.AddSubstitution("place", {"office", "home", "office"})
        smtpapi.AddSection("office", "中野")
        smtpapi.AddSection("home", "目黒")
        smtpapi.SetCategory("カテゴリ1")

        Dim email = New SendGrid.SendGridMessage()
        email.AddTo(from)  ' SmtpapiのSetTo()を使用しているため、実際にはこのアドレスにはメールは送信されない
        email.From = New MailAddress(from, "送信者名")
        email.Subject = "[sendgrid-vb-example] フクロウのお名前はfullnameさん"
        email.Text = "familyname さんは何をしていますか？\r\n 彼はplaceにいます。"
        email.Html = "<strong> familyname さんは何をしていますか？</strong><br />彼はplaceにいます。"
        email.Headers.Add("X-Smtpapi", smtpapi.JsonString())
        email.AddAttachment("..\..\gif.gif")

        Dim credentials = New NetworkCredential(sendGridUserName, sendGridPassword)
        Dim web As Web = New Web(credentials)
        web.Deliver(email)

    End Sub

End Module
