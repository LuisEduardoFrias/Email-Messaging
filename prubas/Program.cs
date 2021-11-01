using EmailMessaging;
using System;

namespace prubas
{
    class Program
    {
        static void Main(string[] args)
        {
            Email.Result resul = Email.SendMail(new EmailCredential("luiseduardofrias27@hotmail.com", "Robert941127"), "luiseduardofrias27@gmail.com", "Pruebas", "Este es un mensaje de prueba.");

            Console.WriteLine(resul.ToString());
        }
    }
}
