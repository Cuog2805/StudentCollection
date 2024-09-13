using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using StudentCollection.Data;
using StudentCollection.Models;

namespace StudentCollection.ViewComponents
{
    public class StudentMenuViewComponent: ViewComponent
    {
        private readonly ApplicationDBContext _db;
        public StudentMenuViewComponent(ApplicationDBContext db)
        {
            _db = db;
        }
        public IViewComponentResult Invoke(int pageIndex)
        {
            var user = _db.Users.Include(m => m.Students)
                .FirstOrDefault(m => m.UserName == HttpContext.Session.GetString("user"));
            if(user != null)
            {
                var studentList = user.Students.ToList();
                return View(PaginatedList<Student>.Create(studentList.AsQueryable(), pageIndex, 10));
            }
            return View(null);
        }
    }
}
