using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Threading.Tasks;

namespace MySchedule
{
    class Mail
    {
        public Mail(string message)
        {
            MailMessage newMail = new MailMessage();

            newMail.From = new MailAddress("*********@hotmail.com");
            newMail.To.Add(new MailAddress("*********@hotmail.com"));
            newMail.IsBodyHtml = true;
            if (DateTime.Today.DayOfWeek == DayOfWeek.Sunday)
            {
                newMail.Subject = "Schedule for the week of " + DateTime.Now.Date.AddDays(1).ToString("yyyy/MM/dd") + " to " + DateTime.Now.Date.AddDays(5).ToString("yyyy/MM/dd");
                message = "<h3>" + "Schedule for the week of " + DateTime.Now.Date.AddDays(1).ToString("yyyy/MM/dd") + " to " + DateTime.Now.Date.AddDays(5).ToString("yyyy/MM/dd") +"</h3>" + message;
            }
            else
            {
                newMail.Subject = "Schedule for " + DateTime.Today.DayOfWeek + " " + DateTime.Now.Date.ToString("yyyy/MM/dd");
                message = "<h3>" + "Schedule for " + DateTime.Today.DayOfWeek + " " + DateTime.Now.Date.ToString("yyyy/MM/dd") + "</h3>" + message;
            }
            newMail.Body = message;

            SmtpClient smtp = new SmtpClient();
            smtp.Port = 25;
            smtp.Host = "smtp.live.com";

            smtp.UseDefaultCredentials = false;
            NetworkCredential cred = new NetworkCredential("********@hotmail.com", "**********");
            smtp.Credentials = cred;
            smtp.EnableSsl = true;

            try
            {
                smtp.Send(newMail);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
