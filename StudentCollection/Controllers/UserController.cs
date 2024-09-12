using Microsoft.AspNetCore.Mvc;
using StudentCollection.Data;
using StudentCollection.Models;

namespace StudentCollection.Controllers
{
    public class UserController : Controller
    {
        private readonly ApplicationDBContext db;
        public UserController(ApplicationDBContext _db) 
        {
            db = _db;
        }
        public IActionResult SignIn()
        {
            try
            {
                User user = new User() { UserName = "", Password = "", FilePath = "" };
                ViewBag.signinError = null;
                return View(user);
            }
            catch
            {
                return RedirectToAction("Error", "Home");
            }
        }
        [HttpPost]
        public IActionResult SignIn(User user, string passwordRepeat)
        {
            //try
            //{
                string signinError = string.Empty;
                if (string.IsNullOrWhiteSpace(user.UserName) || string.IsNullOrWhiteSpace(user.Password))
                {
                    ViewBag.signinError = "Mật khẩu hoặc Username quá ngắn!";
                    return View(user);
                }
                if (user.Password != passwordRepeat)
                {
                    ViewBag.signinError = "Mật khẩu nhập lại không đúng!";
                    return View(user);
                }
                if (db.Users.Any(m => m.UserName == user.UserName.Trim()))
                {
                    ViewBag.signinError = "Username đã tồn tại!";
                    return View(user);
                }
                user.Password = LoginMethod.HashPassword(user.Password);
                db.Users.Add(user);
                db.SaveChanges();
                ViewBag.signinError = "Đăng ký thành công";
                return View(new User());
            //}
            //catch
            //{
            //    return RedirectToAction("Error", "Home");
            //}
        }
        public IActionResult Login(User user)
        {
            return View(user);
        }
        [HttpPost]
        public IActionResult Login(string username, string password)
        {
            try
            {
                HttpContext.Session.Clear();
                ViewBag.loginError = null;
                var user = db.Users.FirstOrDefault(m => m.UserName == username);
                if (user != null && LoginMethod.VertifyPassword(password, user.Password))
                {
                    HttpContext.Session.SetString("user", user.UserName);
                    return RedirectToAction("Index", "Home");
                }
                else
                {
                    ViewBag.loginError = "Sai username hoặc mật khẩu!";
                    return View("Login", new User() { UserName = username, Password = password });
                }
            }
            catch
            {
                return RedirectToAction("Error", "Home");
            }
        }
        public ActionResult Logout()
        {
            HttpContext.Session.Clear();
            return RedirectToAction("Login", "User");
        }
    }
}
