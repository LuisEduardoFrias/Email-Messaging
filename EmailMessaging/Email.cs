using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Text;

namespace EmailMessaging
{
    /// <summary>
    /// 
    /// </summary>
    public static class Email
    {
        /// <summary>
        /// Metodo para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Cc">Email de a copiar.</param>
        /// <param name="Bcc">Email de a copiar.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="ReplyTo">Obtiene o establece la dirección Responder a para el mensaje de correo.</param>
        /// <param name="Sender">Obtiene o establece la dirección del remitente de este mensaje de correo electrónico. </param>
        /// <param name="AttachmentsDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        public static Result SendMail
        (
            EmailCredential emailCredential,
            string To,
            string Subject,
            string MessageBody,
            string ReplyTo = null,
            string Sender = null,
            string Cc = null,
            string Bcc = null,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            if (CheckEmails(new[] { To }, new[] { Cc }, new[] { Bcc }, ReplyTo, Sender, emailCredential.Email) != Result.Successful)
                return Result.WrongDestinationEmail;

            return GenerateSendEmail(emailCredential, InstanceMailMessage(new[] { To }, new[] { Cc }, new[] { Bcc }, ReplyTo, Sender, emailCredential.Email, Subject, MessageBody, AttachmentsDocuments));
        }

        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Cc">Email de a copiar.</param>
        /// <param name="Bcc">Email de a copiar.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="ReplyTo">Obtiene o establece la dirección Responder a para el mensaje de correo.</param>
        /// <param name="Sender">Obtiene o establece la dirección del remitente de este mensaje de correo electrónico. </param>
        /// <param name="AttachmentsDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        public static Result SendBulkEmail
        (
            EmailCredential emailCredential,
            string[] To,
            string Subject,
            string MessageBody,
            string ReplyTo = null,
            string Sender = null,
            string[] Cc = null,
            string[] Bcc = null,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            if (CheckEmails(To, Cc, Bcc, ReplyTo, Sender, emailCredential.Email) == Result.WrongDestinationEmail)
                return Result.WrongDestinationEmail;

            return GenerateSendEmail(emailCredential, InstanceMailMessage(To, Cc, Bcc, ReplyTo, Sender, emailCredential.Email, Subject, MessageBody, AttachmentsDocuments));
        }

        /// <summary>
        /// Metodo asincorono para enviar un mensaje a un correo determinado.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Cc">Email de a copiar.</param>
        /// <param name="Bcc">Email de a copiar.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="ReplyTo">Obtiene o establece la dirección Responder a para el mensaje de correo.</param>
        /// <param name="Sender">Obtiene o establece la dirección del remitente de este mensaje de correo electrónico. </param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="AttachmentsDocuments"></param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        public static async Task<Result> SendMailAsync
        (
            EmailCredential emailCredential,
            string To,
            string Subject,
            string MessageBody,
            string ReplyTo = null,
            string Sender = null,
            string Cc = null,
            string Bcc = null,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {
            if (CheckEmails(new[] { To }, new[] { Cc }, new[] { Bcc }, ReplyTo, Sender, emailCredential.Email) == Result.WrongDestinationEmail)
                return Result.WrongDestinationEmail;

            return await GenerateSendEmailAsync(emailCredential, InstanceMailMessage(new[] { To }, new[] { Cc }, new[] { Bcc }, ReplyTo, Sender, emailCredential.Email, Subject, MessageBody, AttachmentsDocuments));
        }

        /// <summary>
        /// Metodo para enviar un mensaje a varios correos electronicos.
        /// </summary>
        /// <param name="emailCredential">credenciales del Email propietario.</param>
        /// <param name="To">Email de destino.</param>
        /// <param name="Cc">Email de a copiar.</param>
        /// <param name="Bcc">Email de a copiar.</param>
        /// <param name="Subject">Asulto del mesaje.</param>
        /// <param name="MessageBody">Cuerpo del mensaje.</param>
        /// <param name="ReplyTo">Obtiene o establece la dirección Responder a para el mensaje de correo.</param>
        /// <param name="Sender">Obtiene o establece la dirección del remitente de este mensaje de correo electrónico. </param>
        /// <param name="AttachmentsDocuments">Documnetos adjuntos.</param>
        /// <returns>Retorn 
        /// 1: si el mensaje se envio exitosamente; 
        /// 2: El Email de destino es incorrecto; 
        /// 3: El Email es incorrecto;
        /// 4: Si ubo un error al embiar el mensaje.
        /// </returns>
        [Obsolete]
        public static async Task<Result> SendBulkEmailAsync
        (
            EmailCredential emailCredential,
            string[] To,
            string Subject,
            string MessageBody,
            string ReplyTo = null, 
            string Sender = null,
            string[] Cc = null,
            string[] Bcc = null,
            ICollection<Attachment> AttachmentsDocuments = null
        )
        {

            if (CheckEmails(To, Cc, Bcc, ReplyTo, Sender, emailCredential.Email) == Result.WrongDestinationEmail)
                return Result.WrongDestinationEmail;

            return await GenerateSendEmailAsync(emailCredential, InstanceMailMessage(To, Cc, Bcc, ReplyTo, Sender, emailCredential.Email, Subject, MessageBody, AttachmentsDocuments));
        }

        private static Result CheckEmails(string[] To, string[] Cc, string[] Bcc, string ReplyTo, string Sender, string credentialEmail)
        {
            if(To[0] == null)
                return Result.EmailMissing;
            else
                foreach (string to in To)
                {
                    if (!VerifyEmail(to))
                        return Result.WrongDestinationEmail; //Email de destino incorrecto.
                }

            if (Cc[0] != null)
                foreach (string cc in Cc)
                {
                    if (!VerifyEmail(cc))
                        return Result.WrongDestinationEmail; //Email de destino incorrecto.
                }

            if (Bcc[0] != null)
                foreach (string bcc in Bcc)
                {
                    if (!VerifyEmail(bcc))
                        return Result.WrongDestinationEmail; //Email de destino incorrecto.
                }

            if (ReplyTo != null)
                if (!VerifyEmail(ReplyTo))
                    return Result.WrongDestinationEmail; //Email de destino incorrecto.

            if (Sender != null)
                if (!VerifyEmail(Sender))
                    return Result.WrongDestinationEmail; //Email de destino incorrecto.

            if (credentialEmail == null || credentialEmail == "")
                return Result.EmailMissing;
            else
                if (!VerifyEmail(credentialEmail))
                    return Result.WrongEmail; //Email incorrecto.

            return Result.Successful;
        }

        [Obsolete]
        private static MailMessage InstanceMailMessage(string[] To, string[] Cc, string[] Bcc, string ReplyTo, string Sender, string credentialEmail, string Subject, string MessageBody, ICollection<Attachment> AttachmentsDocuments = null)
        {
            MailMessage Mail = new MailMessage();

            Parallel.ForEach(To, to =>
            {
                Mail.To.Add(new MailAddress(to));
            });

            if (Cc[0] != null)
                Parallel.ForEach(Cc, cc =>
                {
                    Mail.CC.Add(new MailAddress(cc));
                });

            if (Bcc[0] != null)
                Parallel.ForEach(Bcc, bcc =>
                {
                    Mail.Bcc.Add(new MailAddress(bcc));
                });

            if (ReplyTo != null && ReplyTo != "")
                Mail.ReplyTo = new MailAddress(ReplyTo);

            if (Sender != null && Sender != "")
                Mail.Sender = new MailAddress(Sender);

            Mail.From = new MailAddress(credentialEmail);

            Mail.Subject = Subject;
            Mail.SubjectEncoding = Encoding.UTF8;


            Mail.Body = MessageBody;
            Mail.BodyEncoding = Encoding.UTF8;

            Mail.IsBodyHtml = true;
            Mail.Priority = MailPriority.High;
            //Mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;

            if (AttachmentsDocuments != null)
                Parallel.ForEach(AttachmentsDocuments, Attachments =>
                {
                    Mail.Attachments.Add(Attachments);
                });

            return Mail;
        }

        private static async Task<Result> GenerateSendEmailAsync(EmailCredential emailCredential, MailMessage Mail)
        {
            return await Task.Run(() => SendEmail(CreateInstanceSmtp(emailCredential), Mail));
        }

        private static Result GenerateSendEmail(EmailCredential emailCredential, MailMessage Mail)
        {
            return SendEmail(CreateInstanceSmtp(emailCredential), Mail);
        }

        private static SmtpClient CreateInstanceSmtp(EmailCredential emailCredential)
        {
            return new SmtpClient
            {
                Host = GetClientEmail(emailCredential.Email[(emailCredential.Email.LastIndexOf('@') + 1)..]),
                Port = GetClientPost(emailCredential.Email[(emailCredential.Email.LastIndexOf('@') + 1)..]),

                UseDefaultCredentials = false,
                EnableSsl = true,
                Credentials = new System.Net.NetworkCredential(emailCredential.Email, emailCredential.Password)
            };
        }

        private static Result SendEmail(SmtpClient client, MailMessage Mail)
        {
            try
            {
                client.Send(Mail);

                System.Collections.Specialized.NameValueCollection Headers = Mail.Headers;

                return Result.Successful;
            }
            catch (Exception ex)
            {
                ex.Message.ToString();

                System.Collections.Specialized.NameValueCollection Headers = Mail.Headers;

                return Result.ErrorSendingMessage;
            }
        }

        private static string GetClientEmail(string clientEmail) =>
            clientEmail switch
            {
                "live.com" => "smtp.live.com",
                "hotmail.com" => "smtp.live.com",
                "outlook.com" => "smtp-mail.outlook.com",
                "gmail.com" => "smtp.gmail.com", //smtp-relay.gmail.com
                "yahoo.com" => "smtp.mail.yahoo.com",
                _ => throw new NotImplementedException("client email not found")
            };

        private static int GetClientPost(string clientEmail) =>
            clientEmail switch
            {
                "live.com" => 465,
                "hotmail.com" => 25,
                "gmail.com" => 465,
                "yahoo.com" => 465,
                "outlook.com" => 587,
                _ => throw new NotImplementedException("client post not found")
            };

        /// <summary>
        /// Tipos de repuestas.
        /// </summary>
        public enum Result
        {
            /// <summary>
            /// repuesta exitosa
            /// </summary>
            Successful,
            /// <summary>
            /// Correo electrónico de destino incorrecto 
            /// </summary>
            WrongDestinationEmail,
            /// <summary>
            /// Email incorrecto
            /// </summary>
            WrongEmail,
            /// <summary>
            /// Error al enviar el mensaje 
            /// </summary>
            ErrorSendingMessage,
            /// <summary>
            /// Falta el correo electrónico 
            /// </summary>
            EmailMissing
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        /// <returns></returns>
        public static string GetResulInesDO(Result result) =>
        result switch
        {
            Result.Successful => "Exitoso",
            Result.WrongDestinationEmail => "Email de destino incorrecto",
            Result.WrongEmail => "Email incorrecto",
            Result.ErrorSendingMessage => "Error al enviar el mensaje puede ser error de coneccion",
            _ => throw new NotImplementedException()
        };

        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        /// <returns></returns>
        public static string GetResulIninUS(Result result) =>
        result switch
        {
            Result.Successful => "Successful",
            Result.WrongDestinationEmail => "Wrong destination email",
            Result.WrongEmail => "Wrong email",
            Result.ErrorSendingMessage => "Error when sending the message may be a connection error",
            _ => throw new NotImplementedException()
        };

        private static bool VerifyEmail(String email) =>
            Regex.IsMatch(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*") &&
            (Regex.Replace(email, "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*", String.Empty).Length == 0);
    }
}