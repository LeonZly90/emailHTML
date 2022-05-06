using System;
using System.Net.Mail;


public class LocalEmailService
{
    
    public static void SendEmail(string[] to, string subject, string body)
    {
        string fromAddress = "hm-web@pepperconstruction.com";
        string host = "10.1.5.175";
        try
        {
            MailMessage msg = new();
            msg.From = new MailAddress(fromAddress);
            foreach (string email in to)
            {
                msg.To.Add(new MailAddress(email));
            }
            msg.Subject = subject;
            msg.Body = body;
            msg.IsBodyHtml = true;

            Attachment attachment;
            attachment = new Attachment(@"C:\PepperPepper\HM_Expired.xlsx"); 
            msg.Attachments.Add(attachment);

            SmtpClient client = new();
            client.Host = host;
            client.Send(msg);
        }
        catch (Exception)
        {
            throw;
        }
    }
}
