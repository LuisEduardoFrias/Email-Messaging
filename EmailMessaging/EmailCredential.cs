
namespace EmailMessaging
{
    public class EmailCredential
    {
        public string Email { get; set; }
        public string Password { get; set; }

        public EmailCredential(string email, string password)
        {
            Email = email;
            Password = password;
        }

        public EmailCredential()
        {
                
        }

    }
}
