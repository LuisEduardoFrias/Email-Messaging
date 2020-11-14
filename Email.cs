using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using Telerik.Reporting.Processing;
using System.IO;
using System.Threading.Tasks;

public async Task<string> SendMail(string From, string Password, string To, string Subject, string MessageBsody, ICollection<AttachDocuments> AttachDocuments, ClientEmail clientEmail)
        {
            return await Task.Run(() => 
            { 
                try
                {
                    var Mail = new MailMessage();
                    Mail.To.Add(new MailAddress(To));
                    Mail.From = new MailAddress(From);
                    Mail.Subject = Subject;
                    Mail.Body = MessageBsody;
                    Mail.IsBodyHtml = false;


                    Parallel.ForEach(AttachDocuments, Document =>
                    {
                         Mail.Attachments.Add(new Attachment(Document.Document, Document.Name + Document.Extension));
                    });


                    using (var client = new SmtpClient( GetClientEmail(clientEmail), clientEmail))
                    {
                        client.Credentials = new System.Net.NetworkCredential(From, Password);
                        client.EnableSsl = true;
                        client.Send(Mail);
                    }

                    return "Mensaje enviado Exitosamente";

                }
                catch (Exception ex)
                {
                    return ex.Message.ToString();
                }

            });
        }

        public async Task<string> SendMail(string From, string Password, string To, string Subject, string MessageBsody, ClientEmail clientEmail)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var Mail = new MailMessage();
                    Mail.To.Add(new MailAddress(To));
                    Mail.From = new MailAddress(From);
                    Mail.Subject = Subject;
                    Mail.Body = MessageBsody;
                    Mail.IsBodyHtml = false;

                    using (var client = new SmtpClient(GetClientEmail(clientEmail), clientEmail))
                    {
                        client.Credentials = new System.Net.NetworkCredential(From, Password);
                        client.EnableSsl = true;
                        client.Send(Mail);
                    }

                    return "Mensaje enviado Exitosamente";

                }
                catch (Exception ex)
                {
                    return ex.Message.ToString();
                }

            });
        }


        private static string GetClientEmail(ClientEmail clientEmail) =>
        clientEmail switch
        {
            ClientEmail.Live => "smtp.live.com",
            ClientEmail.Hotmail => "smtp.hotmail.com",
            ClientEmail.Gmail => "smtp.gmail.com",
            ClientEmail.Yahoo => "smtp.mail.yahoo.com",
            ClientEmail.Outlook => "Smtp.live.com",
            _ => throw new NotImplementedException()
        };