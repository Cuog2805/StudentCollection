using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using StudentCollection.Data;
using StudentCollection.Models;
using System.Diagnostics;
using System.Globalization;

namespace StudentCollection.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDBContext db;
        public HomeController(ApplicationDBContext _db)
        {
            db = _db;
        }

        public IActionResult Index(string filename)
        {
            if(HttpContext.Session.GetString("user") == null)
            {
                return RedirectToAction("Login", "User");
            }
            string? userNameCurrent = HttpContext.Session.GetString("user");
            var user = db.Users.FirstOrDefault(m => m.UserName == userNameCurrent);
            if (user != null)
            {
                user.FilePath = filename != null ? filename : user.FilePath;
            }
            db.SaveChanges();

            return View(user);
        }
        [HttpPost]
        public IActionResult Index(IFormFile fileInput)
        {
            try
            {
                if (fileInput != null && fileInput.Length > 0)
                {
                    var filename = fileInput.FileName;
                    var fileExtention = Path.GetExtension(filename);
                    if (fileExtention == ".xlsx")
                    {
                        string userNameCurrent = HttpContext.Session.GetString("user");
                        User userCurrent = db.Users.Include(m => m.Students).First(m => m.UserName == userNameCurrent);

                        if (userCurrent.Students != null)
                        {
                            if (userCurrent.Students.Count() > 0)
                            {
                                db.Students.RemoveRange(userCurrent.Students.ToList());
                                db.SaveChanges();
                            }
                        }

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        Stream stream = new MemoryStream();
                        fileInput.CopyTo(stream);
                        using (ExcelPackage excelPackage = new ExcelPackage(stream))
                        {
                            var sheetNum = excelPackage.Workbook.Worksheets.Count;
                            int stt = 1;
                            for (int i = 0; i < sheetNum; i++)
                            {
                                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[i];
                                var rowCount = worksheet.Dimension.End.Row;

                                int headerRow = 6;
                                int endRow = 0;

                                for (int row = 1; row < 10; row++)
                                {
                                    if (worksheet.Cells[row, 1].Text.Trim() == "1")
                                    {
                                        headerRow = row;
                                        break;
                                    }
                                }

                                for (int row = headerRow + 1; row < rowCount; row++)
                                {
                                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                                    {
                                        endRow = row;
                                        break;
                                    }
                                }

                                for (int row = headerRow; row < endRow; row++)
                                {
                                    Student student = new Student();
                                    if (worksheet.Name.ToString()[0] == '6')
                                    {
                                        student.Stt = stt;
                                        stt++;

                                        student.Name = string.Concat(worksheet.Cells[row, 2].Text, " ", worksheet.Cells[row, 3].Text);
                                        student.Class = worksheet.Name.ToString();
                                        //birth
                                        try
                                        {
                                            if (DateTime.TryParseExact(
                                                worksheet.Cells[row, 4].Text,
                                                new string[] { "dd/MM/yyyy", "d/M/yyyy", "MM/dd/yyyy", "M/d/yyyy" },
                                                new CultureInfo("en-US"),
                                                DateTimeStyles.None,
                                                out DateTime bitrh))
                                            {
                                                student.Birth = bitrh;
                                            }
                                            else
                                            {
                                                student.Birth = DateTime.Now;
                                            }
                                        }
                                        catch
                                        {
                                            return View("Error", "Home");
                                        }
                                        student.Gender = worksheet.Cells[row, 5].Text;
                                        student.CurrentResidence = worksheet.Cells[row, 8].Text;
                                        student.PermanentResidece = worksheet.Cells[row, 8].Text;
                                        student.BirthPlace = worksheet.Cells[row, 8].Text;
                                        student.FatherName = worksheet.Cells[row, 6].Text;
                                        student.MotherName = worksheet.Cells[row, 7].Text;
                                        student.PhoneNumber = worksheet.Cells[row, 9].Text.Count() == 9 ? string.Concat("0", worksheet.Cells[row, 9].Text) : worksheet.Cells[row, 9].Text;
                                        student.UserID = userCurrent.UserID;
                                    }
                                    else
                                    {
                                        student.Stt = stt;
                                        stt++;

                                        student.Name = string.Concat(worksheet.Cells[row, 2].Text, " ", worksheet.Cells[row, 3].Text);
                                        student.Class = worksheet.Name.ToString();
                                        //birth
                                        try
                                        {
                                            if (DateTime.TryParseExact(
                                                worksheet.Cells[row, 5].Text,
                                                new string[] { "dd/MM/yyyy", "d/M/yyyy", "MM/dd/yyyy", "M/d/yyyy" },
                                                new CultureInfo("en-US"),
                                                DateTimeStyles.None,
                                                out DateTime bitrh))
                                            {
                                                student.Birth = bitrh;
                                            }
                                            else
                                            {
                                                student.Birth = DateTime.Now;
                                            }
                                        }
                                        catch
                                        {
                                            return View("Error", "Home");
                                        }
                                        student.Gender = worksheet.Cells[row, 4].Text;
                                        student.CurrentResidence = worksheet.Cells[row, 8].Text;
                                        student.PermanentResidece = worksheet.Cells[row, 8].Text;
                                        student.BirthPlace = worksheet.Cells[row, 7].Text;
                                        student.FatherName = worksheet.Cells[row, 9].Text;
                                        student.MotherName = worksheet.Cells[row, 10].Text;
                                        student.PhoneNumber = worksheet.Cells[row, 11].Text.Count() == 9 ? string.Concat("0", worksheet.Cells[row, 11].Text) : worksheet.Cells[row, 11].Text;
                                        student.UserID = userCurrent.UserID;
                                    }

                                    db.Students.Add(student);
                                }
                            }
                        }
                        db.SaveChanges();

                        return RedirectToAction("Index", "Home", new { fileInput.FileName });
                    }
                    else
                    {
                        return RedirectToAction("Index", "Home", null);
                    }
                }
                else
                {
                    return RedirectToAction("Index", "Home", null);
                }
            }
            catch
            {
                return View("Error");
            }
        }
        public ActionResult Export()
        {
            try
            {
                var userNameCurrent = HttpContext.Session.GetString("user");
                User userCurrent = db.Users.Include(m => m.Students).First(m => m.UserName == userNameCurrent);

                Stream excelfile = new MemoryStream();
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(excelfile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    worksheet.Cells[1, 1].Value = "STT";
                    worksheet.Cells[1, 2].Value = "Họ và tên";
                    worksheet.Cells[1, 3].Value = "Lớp";
                    worksheet.Cells[1, 4].Value = "Ngày sinh";
                    worksheet.Cells[1, 5].Value = "Giới tính";
                    worksheet.Cells[1, 6].Value = "Chỗ ở hiện nay";
                    worksheet.Cells[1, 7].Value = "Hộ khẩu thường trú";
                    worksheet.Cells[1, 8].Value = "Nơi sinh";
                    worksheet.Cells[1, 9].Value = "Tên cha";
                    worksheet.Cells[1, 10].Value = "Tên mẹ";
                    worksheet.Cells[1, 11].Value = "Điện thoại";

                    worksheet.Cells["A1:J1"].Style.Font.Bold = true;

                    int row = 2;
                    foreach (var student in userCurrent.Students.ToList())
                    {
                        worksheet.Cells[row, 1].Value = student.Stt;
                        worksheet.Cells[row, 2].Value = student.Name;
                        worksheet.Cells[row, 3].Value = student.Class;

                        var studentAge = DateTime.Now.Year - student.Birth.Year;
                        //
                        if (studentAge > 18 || studentAge <= 5 || studentAge != (int.Parse(student.Class[0].ToString()) + 5))
                        {
                            worksheet.Cells[row, 4].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 4].Value = student.Birth.ToString("dd/MM/yyyy");
                        //
                        if (string.IsNullOrWhiteSpace(student.Gender))
                        {
                            worksheet.Cells[row, 5].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 5].Value = student.Gender;
                        //
                        if (string.IsNullOrWhiteSpace(student.CurrentResidence))
                        {
                            worksheet.Cells[row, 6].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 6].Value = student.CurrentResidence;
                        //
                        if (string.IsNullOrWhiteSpace(student.PermanentResidece))
                        {
                            worksheet.Cells[row, 7].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 7].Value = student.PermanentResidece;
                        //
                        if (string.IsNullOrWhiteSpace(student.BirthPlace))
                        {
                            worksheet.Cells[row, 8].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 8].Value = student.BirthPlace;
                        //
                        if (string.IsNullOrWhiteSpace(student.FatherName))
                        {
                            worksheet.Cells[row, 9].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 9].Value = student.FatherName;
                        //
                        if (string.IsNullOrWhiteSpace(student.MotherName))
                        {
                            worksheet.Cells[row, 10].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 10].Value = student.MotherName;
                        //
                        if (string.IsNullOrWhiteSpace(student.PhoneNumber))
                        {
                            worksheet.Cells[row, 11].Style.Fill.SetBackground(System.Drawing.Color.Red);
                        }
                        worksheet.Cells[row, 11].Value = student.PhoneNumber;

                        row++;
                    }
                    worksheet.Cells.AutoFitColumns();

                    package.Save();
                }
                excelfile.Position = 0;
                return File(excelfile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "student_data.xlsx");
            }
            catch
            {
                return View("Error");
            }
        }
        public IActionResult Error()
        {
            return View();
        }
    }
}