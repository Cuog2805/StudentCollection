using System.Text;
using System.Security.Cryptography;

namespace StudentCollection.Models
{
    public class LoginMethod
    {
        public static string salt = "abc!@#zxc";
        public static string HashPassword(string password)
        {
            string hashedPassword = string.Empty;
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password + salt));
                hashedPassword = Convert.ToBase64String(bytes);
            }
            return hashedPassword;
        }
        public static bool VertifyPassword(string password, string storedPassword)
        {
            string hashedPassword = string.Empty;
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password + salt));
                hashedPassword = Convert.ToBase64String(bytes);
            }
            return hashedPassword == storedPassword;
        }
    }
}
