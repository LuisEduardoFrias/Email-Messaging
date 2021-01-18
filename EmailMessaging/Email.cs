
namespace EmailMessaging
{
    using System;
    using System.Collections.Generic;
    using System.Net.Mail;
    using System.Threading.Tasks;
    using System.Text.RegularExpressions;
    using System.Text;

    public static class Email
    {
        /// <summary>
        /// Metodo para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static result SendMail
        (
            EmailCredential emailCredential,
            string To,
            string Subject,
            string MessageBody,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            if (!VerifyEmail(To))
                return result.WrongDestinationEmail; //Email de destino incorrecto.

            if (!VerifyEmail(emailCredential.Email))
                return result.WrongEmail; //Email incorrecto.

            MailMessage Mail = new MailMessage();

            Mail.To.Add(new MailAddress(To));
            Mail.From = new MailAddress(emailCredential.Email);

            Mail.Subject = Subject;
            Mail.SubjectEncoding = Encoding.UTF8;


            Mail.Body = MessageBody;
            Mail.BodyEncoding = Encoding.UTF8;

            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachmentsDocuments != null)
                Parallel.ForEach(AttachmentsDocuments, Attachments =>
                {
                    Mail.Attachments.Add(Attachments);
                });

            SmtpClient client = new SmtpClient();

            client.Host = GetClientEmail(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));
            client.Port = GetClientPost(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));

            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                client.Send(Mail);

                return result.Successful;
            }
            catch
            {
                return result.ErrorSendingMessage;
            }
        }

        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static result SendBulkEmail
        (
            EmailCredential emailCredential,
            string[] To,
            string Subject,
            string MessageBody,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            foreach (string to in To)
            {
                if (!VerifyEmail(to))
                    return result.WrongDestinationEmail; //Email de destino incorrecto.
            }

            if (!VerifyEmail(emailCredential.Email))
                return result.WrongEmail; //Email incorrecto.

            MailMessage Mail = new MailMessage();

            Parallel.ForEach(To, to =>
            {
                Mail.To.Add(new MailAddress(to));
            });


            Mail.From = new MailAddress(emailCredential.Email);

            Mail.Subject = Subject;
            Mail.SubjectEncoding = Encoding.UTF8;


            Mail.Body = MessageBody;
            Mail.BodyEncoding = Encoding.UTF8;

            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachmentsDocuments != null)
                Parallel.ForEach(AttachmentsDocuments, Attachments =>
                {
                    Mail.Attachments.Add(Attachments);
                });

            SmtpClient client = new SmtpClient();

            client.Host = GetClientEmail(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));
            client.Port = GetClientPost(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));

            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                client.Send(Mail);

                return result.Successful;
            }
            catch
            {
                return result.ErrorSendingMessage;
            }
        }



        /// <summary>
        /// Metodo asincorono para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static async Task<result> SendMailAsync
        (
            EmailCredential emailCredential,
            string To,
            string Subject,
            string MessageBody,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            if (!VerifyEmail(To))
                return result.WrongDestinationEmail; //Email de destino incorrecto.

            if (!VerifyEmail(emailCredential.Email))
                return result.WrongEmail; //Email incorrecto.

            MailMessage Mail = new MailMessage();

            Mail.To.Add(new MailAddress(To));
            Mail.From = new MailAddress(emailCredential.Email);

            Mail.Subject = Subject;
            Mail.SubjectEncoding = Encoding.UTF8;


            Mail.Body = MessageBody;
            Mail.BodyEncoding = Encoding.UTF8;

            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachmentsDocuments != null)
                Parallel.ForEach(AttachmentsDocuments, Attachments =>
                {
                    Mail.Attachments.Add(Attachments);
                });

            SmtpClient client = new SmtpClient();

            client.Host = GetClientEmail(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));
            client.Port = GetClientPost(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));

            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                client.SendMailAsync(Mail);

                return result.Successful;
            }
            catch
            {
                return result.ErrorSendingMessage;
            }

        }



        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="clientEmail">Cliente email</param>
        /// <param name="AttachDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        /// 
        public static async Task<result> SendBulkEmailAsync
        (
            EmailCredential emailCredential,
            string[] To,
            string Subject,
            string MessageBody,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            foreach (string to in To)
            {
                if (!VerifyEmail(to))
                    return result.WrongDestinationEmail; //Email de destino incorrecto.
            }


            if (!VerifyEmail(emailCredential.Email))
                return result.WrongEmail; //Email incorrecto.

            MailMessage Mail = new MailMessage();

            Parallel.ForEach(To, to =>
            {
                Mail.To.Add(new MailAddress(to));
            });

            Mail.From = new MailAddress(emailCredential.Email);

            Mail.Subject = Subject;
            Mail.SubjectEncoding = Encoding.UTF8;


            Mail.Body = MessageBody;
            Mail.BodyEncoding = Encoding.UTF8;

            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;

            if (AttachmentsDocuments != null)
                Parallel.ForEach(AttachmentsDocuments, Attachments =>
                {
                    Mail.Attachments.Add(Attachments);
                });

            SmtpClient client = new SmtpClient();

            client.Host = GetClientEmail(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));
            client.Port = GetClientPost(emailCredential.Email.Substring(emailCredential.Email.LastIndexOf('@') + 1));

            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password);

            try
            {
                client.SendMailAsync(Mail);

                return result.Successful;
            }
            catch
            {
                return result.ErrorSendingMessage;
            }
        }




        private static string GetClientEmail(string clientEmail) =>
            clientEmail switch
            {
                "live.com" => "smtp.live.com",
                "hotmail.com" => "smtp.live.com",
                "gmail.com" => "smtp-relay.gmail.com",
                "yahoo.com" => "smtp.mail.yahoo.com",
                "outlook.com" => "Smtp.live.com",
                _ => throw new NotImplementedException()
            };



        private static int GetClientPost(string clientEmail) =>
            clientEmail switch
            {
                "live.com" => 587,
                "hotmail.com" => 587,
                "gmail.com" => 587,
                "yahoo.com" => 465,
                "outlook.com" => 587,
                _ => throw new NotImplementedException()
            };



        public enum result
        {
            Successful,
            WrongDestinationEmail,
            WrongEmail,
            ErrorSendingMessage
        }


        public static string GetResulInesDO(result result) =>
            result switch
            {
                result.Successful => "Exitoso",
                result.WrongDestinationEmail => "Email de destino incorrecto",
                result.WrongEmail => "Email incorrecto",
                result.ErrorSendingMessage => "Error al enviar el mensaje puede ser error de coneccion",
                _ => throw new NotImplementedException()
            };


        public static string GetResulIninUS(result result) =>
            result switch
            {
                result.Successful => "Successful",
                result.WrongDestinationEmail => "Wrong destination email",
                result.WrongEmail => "Wrong email",
                result.ErrorSendingMessage => "Error when sending the message may be a connection error",
                _ => throw new NotImplementedException()
            };

        private static bool VerifyEmail(String email) =>
            Regex.IsMatch(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*") &&
            (Regex.Replace(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*", String.Empty).Length == 0);

    }
}