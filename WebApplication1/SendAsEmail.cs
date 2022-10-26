using System;
using System.Net.Mail;

namespace ExcelAppOpenXML
{
    public class SendAsEmail
    {
        public static void SendEmail(string emailAddress, Attachment attachment1, Attachment attachment2)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient();

                mail.To.Add(emailAddress);
                mail.From = new MailAddress("productmaster@rocketsoftware.com");
                mail.Subject = "Product Hierarchy Information";
                mail.IsBodyHtml = true;

                string textBody = "Hi, <br /><br /> Please find attached files related to Product Hierarchy Information.<br />";
                                
                textBody += "<br />Thanks,<br />Rocket Team";

                mail.Body = textBody;

                if (attachment1 != null && attachment2 != null)
                {
                    mail.Attachments.Add(attachment1);
                    mail.Attachments.Add(attachment2);
                }

                SmtpServer.Host = "smtp.rocketsoftware.com";
                SmtpServer.Port = 25;
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpServer.Send(mail);

                mail.Dispose();
                SmtpServer.Dispose();
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }
    }
}