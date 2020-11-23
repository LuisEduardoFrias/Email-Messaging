
namespace EmailMessaging
{
    using System;
    using System.Collections.Generic;
    using System.Net.Mail;
    using System.Threading.Tasks;
    using System.Text.RegularExpressions;

    public static class Email
    {
        /// <summary>
        /// Metodo para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBsody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static int SendMail
        (EmailCredential emailCredential, string To, string Subject, string MessageBsody, ICollection<AttachDocuments> AttachDocuments = null)
        {
            if (!VerifyEmail(To))
                return 2; //Email de destino incorrecto.

            if (!VerifyEmail(emailCredential.Email))
                return 3; //Email incorrecto.

            MailMessage Mail = new MailMessage();
            Mail.To.Add(new MailAddress(To));
            Mail.From = new MailAddress(emailCredential.Email);
            Mail.Subject = Subject;
            Mail.Body = MessageBsody;
            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachDocuments.Equals(null))
                Parallel.ForEach(AttachDocuments, Document =>
                {
                    Mail.Attachments.Add(new Attachment(Document.Document, Document.Name + Document.Extension));
                });

            SmtpClient Smtp = new SmtpClient();
            Smtp.Host = GetClientEmail(emailCredential.Email);
            Smtp.Port = GetClientPost(emailCredential.Email);
            Smtp.EnableSsl = false;
            Smtp.UseDefaultCredentials = false;
            Smtp.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                Smtp.SendAsync(Mail, new { });
                Mail.Dispose();

                return 1;
            }
            catch (Exception ex)
            {
                ex.Message.ToString();

                return 4;
            }

        }

        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBsody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static int SendBulkEmail
        (EmailCredential emailCredential, string[] To, string Subject, string MessageBsody, ICollection<AttachDocuments> AttachDocuments = null)
        {
            foreach (string to in To)
            {
                if (!VerifyEmail(to))
                    return 2; //Email de destino incorrecto.
            }

            if (!VerifyEmail(emailCredential.Email))
                return 3; //Email incorrecto.

            MailMessage Mail = new MailMessage();
            Mail.From = new MailAddress(emailCredential.Email);
            Mail.Subject = Subject;
            Mail.Body = MessageBsody;
            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachDocuments.Equals(null))
                Parallel.ForEach(AttachDocuments, Document =>
                {
                    Mail.Attachments.Add(new Attachment(Document.Document, Document.Name + Document.Extension));
                });

            SmtpClient Smtp = new SmtpClient();
            Smtp.Host = GetClientEmail(emailCredential.Email);
            Smtp.Port = GetClientPost(emailCredential.Email);
            Smtp.EnableSsl = false;
            Smtp.UseDefaultCredentials = false;
            Smtp.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                foreach (var to in To)
                {
                    Mail.To.Add(new MailAddress(to));

                    Smtp.SendAsync(Mail, new { });
                    Mail.Dispose();
                }

                Smtp.SendAsync(Mail, new { });
                Mail.Dispose();

                return 1;
            }
            catch (Exception ex)
            {
                ex.Message.ToString();

                return 4;
            }
        }

        /// <summary>
        /// Metodo asincorono para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBsody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static async Task<int> SendMailAsync
        (EmailCredential emailCredential, string To, string Subject, string MessageBsody, ICollection<AttachDocuments> AttachDocuments = null)
        {
            if (!VerifyEmail(To))
                return 2; //Email de destino incorrecto.

            if (!VerifyEmail(emailCredential.Email))
                return 3; //Email incorrecto.

            MailMessage Mail = new MailMessage();
            Mail.To.Add(new MailAddress(To));
            Mail.From = new MailAddress(emailCredential.Email);
            Mail.Subject = Subject;
            Mail.Body = MessageBsody;
            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachDocuments.Equals(null))
                Parallel.ForEach(AttachDocuments, Document =>
                {
                    Mail.Attachments.Add(new Attachment(Document.Document, Document.Name + Document.Extension));
                });

            SmtpClient Smtp = new SmtpClient();
            Smtp.Host = GetClientEmail(emailCredential.Email);
            Smtp.Port = GetClientPost(emailCredential.Email);
            Smtp.EnableSsl = false;
            Smtp.UseDefaultCredentials = false;
            Smtp.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                return await Task.Run(() =>
                {
                    Smtp.SendAsync(Mail, new { });
                    Mail.Dispose();

                    return 1;
                });
            }
            catch (Exception ex)
            {
                ex.Message.ToString();

                return 4;
            }

        }

        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBsody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static async Task<int> SendBulkEmailAsync
        (EmailCredential emailCredential, string[] To, string Subject, string MessageBsody, ICollection<AttachDocuments> AttachDocuments = null)
        {
            foreach(string to in To)
            {
                if (!VerifyEmail(to))
                    return 2; //Email de destino incorrecto.
            }
             
            if (!VerifyEmail(emailCredential.Email))
                return 3; //Email incorrecto.

            MailMessage Mail = new MailMessage();
            Mail.From = new MailAddress(emailCredential.Email);
            Mail.Subject = Subject;
            Mail.Body = MessageBsody;
            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachDocuments.Equals(null))
                Parallel.ForEach(AttachDocuments, Document =>
                {
                    Mail.Attachments.Add(new Attachment(Document.Document, Document.Name + Document.Extension));
                });

            SmtpClient Smtp = new SmtpClient();
            Smtp.Host = GetClientEmail(emailCredential.Email);
            Smtp.Port = GetClientPost(emailCredential.Email);
            Smtp.EnableSsl = false;
            Smtp.UseDefaultCredentials = false;
            Smtp.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                return await Task.Run(() =>
                {
                    foreach (var to in To)
                    {
                        Mail.To.Add(new MailAddress(to));

                        Smtp.SendAsync(Mail, new { });
                        Mail.Dispose();
                    }

                    return 1;
                });
            }
            catch (Exception ex)
            {
                ex.Message.ToString();

                return 4;
            }
        }

        private static string GetClientEmail(string clientEmail) =>
            clientEmail.ToLower()[(clientEmail.IndexOf("@") + 1)..] switch
            {
                "live.com" => "smtp.live.com",
                "hotmail.com" => "smtp.hotmail.com",
                "Gmail.com" => "smtp.gmail.com",
                "Yahoo.com" => "smtp.mail.yahoo.com",
                "Outlook.com" => "Smtp.live.com",
                _ => throw new NotImplementedException()
            };

        private static int GetClientPost(string clientEmail) =>
            clientEmail.ToLower()[(clientEmail.IndexOf("@") + 1)..] switch
            {
                "live.com" => 587,
                "hotmail.com" => 587,
                "Gmail.com" => 587,
                "Yahoo.com" => 465,
                "Outlook.com" => 587,
                _ => throw new NotImplementedException()
            };

        private static bool VerifyEmail(String email) =>
            Regex.IsMatch(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*") && 
            (Regex.Replace(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*", String.Empty).Length == 0);

    }
}