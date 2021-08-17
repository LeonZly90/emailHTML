using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks;
using System.Net.Mail;
using Microsoft.Extensions.Configuration;


public class LocalEmailService
{
    
    public void SendEmail(string to, string subject, string body)
    {
        string fromAddress = "hm-web@pepperconstruction.com";
        string host = "10.1.5.175";

        try
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(fromAddress);
            msg.To.Add(new MailAddress(to));
            msg.Subject = subject;
            msg.Body = body;
            msg.IsBodyHtml = true;
            SmtpClient client = new SmtpClient();
            client.Host = host;
            client.Send(msg);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }
}
