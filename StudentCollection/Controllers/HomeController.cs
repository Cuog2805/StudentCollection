using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using StudentCollection.Data;
using StudentCollection.Models;
using System.Diagnostics;
using System.Globalization;
using static StudentCollection.Controllers.HomeController;

namespace StudentCollection.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDBContext db;
        public HomeController(ApplicationDBContext _db)
        {
            db = _db;
        }

        public IActionResult Index(string filename, int? pageIndex=1)
        {
            if(HttpContext.Session.GetString("user") == null)
            {
                return RedirectToAction("Login", "User");
            }
            string? userNameCurrent = HttpContext.Session.GetString("user");
            var user = db.Users.Include(m => m.Students).FirstOrDefault(m => m.UserName == userNameCurrent);
            if (user != null)
            {
                user.FilePath = filename != null ? filename : user.FilePath;
            }
            db.SaveChanges();

            ViewData["pageIndex"] = pageIndex;
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

                        //Xóa student list trước
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

                                    //tạo lớp chứa tên các cột tương ứng với từng sheet
                                    StudentCol studentCol1 = new StudentCol(row.ToString(), new string[] { "6" }, null, "B", "C", null, "D", "E", "H", "H", "H", "F", "G", "I");
                                    StudentCol studentCol2 = new StudentCol(row.ToString(), new string[] { "7", "8", "9" }, null, "B", "C", null, "E", "D", "H", "H", "G", "I", "J", "K");

                                    //gán giá trị vào model
                                    var sheetName = worksheet.Name.ToString()[0].ToString();
                                    if (studentCol1.SheetType.Contains(sheetName))
                                    {
                                        student = AssignStudent(studentCol1, worksheet);
                                        student.Stt = stt;
                                        stt++;
                                        student.UserID = userCurrent.UserID;
                                    }
                                    if (studentCol2.SheetType.Contains(sheetName))
                                    {
                                        student = AssignStudent(studentCol2, worksheet);
                                        student.Stt = stt;
                                        stt++;
                                        student.UserID = userCurrent.UserID;
                                    }

                                    db.Students.Add(student);
                                }
                            }
                        }
                        db.SaveChanges();
                        return RedirectToAction("Index", "Home", new { filename = fileInput.FileName });
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

                    //create header
                    worksheet.Cells["A1:H1"].Merge = true;
                    worksheet.Cells["A2:H2"].Merge = true;
                    worksheet.Cells["I1:V1"].Merge = true;
                    worksheet.Cells["I2:V2"].Merge = true;
                    worksheet.Cells["A3:V3"].Merge = true;
                    worksheet.Cells["A3:V3"].Merge = true;
                    worksheet.Cells["A1:V3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A1:V3"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["A1"].Value = "PHÒNG GD&ĐT THÀNH PHỐ PHỦ LÝ";
                    worksheet.Cells["I1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                    worksheet.Cells["A2"].Value = "TRƯỜNG THCS TRẦN QUỐC TOẢN";
                    worksheet.Cells["I2"].Value = "Độc lập - Tự do - Hạnh phúc";
                    worksheet.Cells["A3"].Value = "DANH SÁCH HỌC SINH";
                    worksheet.Cells["A3"].Style.Font.Size = 20;
                    worksheet.Row(3).Height = 50;

                    worksheet.Cells[7, 1].Value = "STT";
                    worksheet.Cells[7, 2].Value = "Lớp học";
                    worksheet.Cells[7, 3].Value = "Mã học sinh";
                    worksheet.Cells[7, 4].Value = "Mã MOET";
                    worksheet.Cells[7, 5].Value = "Mã VEMIS";
                    worksheet.Cells[7, 6].Value = "Sổ đăng bộ";
                    worksheet.Cells[7, 7].Value = "Sổ định danh cá nhân";
                    worksheet.Cells[7, 8].Value = "Họ tên";
                    worksheet.Cells[7, 9].Value = "Ngày sinh";
                    worksheet.Cells[7, 10].Value = "Giới tính";
                    worksheet.Cells[7, 11].Value = "Chỗ ở hiện nay";
                    worksheet.Cells[7, 12].Value = "Hộ khẩu thường trú";
                    worksheet.Cells[7, 13].Value = "Nơi sinh";
                    worksheet.Cells[7, 14].Value = "Quê quán";
                    worksheet.Cells[7, 15].Value = "Chứng minh thư";
                    worksheet.Cells[7, 16].Value = "Ngày cấp CMT";
                    worksheet.Cells[7, 17].Value = "Nơi cấp CMT";
                    worksheet.Cells[7, 18].Value = "Dân tộc";
                    worksheet.Cells[7, 19].Value = "Tôn giáo";
                    worksheet.Cells[7, 20].Value = "Diện chính sách";
                    worksheet.Cells[7, 21].Value = "Diện khuyết tật";
                    worksheet.Cells[7, 22].Value = "Cận nghèo";
                    worksheet.Cells[7, 23].Value = "Đoàn viên";
                    worksheet.Cells[7, 24].Value = "Đội viên";
                    worksheet.Cells[7, 25].Value = "Con giáo viên";
                    worksheet.Cells[7, 26].Value = "Tên cha";
                    worksheet.Cells[7, 27].Value = "Nghề nghiệp cha";
                    worksheet.Cells[7, 28].Value = "Năm sinh cha";
                    worksheet.Cells[7, 29].Value = "Tên mẹ";
                    worksheet.Cells[7, 30].Value = "Nghề nghiệp mẹ";
                    worksheet.Cells[7, 31].Value = "Năm sinh mẹ";
                    worksheet.Cells[7, 32].Value = "Điện thoại DĐ";
                    worksheet.Cells[7, 33].Value = "Email";
                    worksheet.Cells[7, 34].Value = "Điện thoại bố";
                    worksheet.Cells[7, 35].Value = "Điện thoại mẹ";
                    worksheet.Cells[7, 36].Value = "Ghi chú";
                    worksheet.Cells[7, 37].Value = "Điện thoại học sinh";
                    worksheet.Cells[7, 38].Value = "Ngày vào trường";

                    worksheet.Cells["A2:AL7"].Style.Font.Bold = true;
                    //create body
                    int row = 8;
                    foreach (var student in userCurrent.Students.ToList())
                    {
                        AsignValue(worksheet.Cells["A" + row.ToString()], student.Stt.ToString());
                        AsignValue(worksheet.Cells["B" + row.ToString()], student.Class);
                        AsignValue(worksheet.Cells["H" + row.ToString()], student.Name);

                        //check tuoi
                        var studentAge = DateTime.Now.Year - student.Birth.Year;
                        //
                        if (studentAge > 18 || studentAge <= 5 || studentAge != (int.Parse(student.Class[0].ToString()) + 5))
                        {
                            worksheet.Cells["I" + row.ToString()].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        }
                        worksheet.Cells["I" + row.ToString()].Value = student.Birth.ToString("dd/MM/yyyy");
                        //
                        AsignValue(worksheet.Cells["J" + row.ToString()], student.Gender);
                        AsignValue(worksheet.Cells["K" + row.ToString()], student.CurrentResidence);
                        AsignValue(worksheet.Cells["L" + row.ToString()], student.PermanentResidece);
                        AsignValue(worksheet.Cells["M" + row.ToString()], student.BirthPlace);
                        AsignValue(worksheet.Cells["Z" + row.ToString()], student.FatherName);
                        AsignValue(worksheet.Cells["AC" + row.ToString()], student.MotherName);
                        AsignValue(worksheet.Cells["AF" + row.ToString()], student.PhoneNumber);

                        row++;
                    }
                    worksheet.Cells.AutoFitColumns();
                    package.Save();
                    //
                }
                excelfile.Position = 0;
                return File(excelfile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "student_data.xlsx");
            }
            catch
            {
                return View("Error");
            }
        }
        public static void AsignValue(ExcelRange worksheet, string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                worksheet.Style.Fill.SetBackground(System.Drawing.Color.Red);
            }
            worksheet.Value = value;
        }
        public class StudentCol
        {
            public string RowIndex { get; set; }
            public string? SttCol { get; set; }
            public string? FirstNameCol { get; set; }
            public string? LastNameCol { get; set; }
            public string? ClassCol { get; set; }
            public string? BirthCol { get; set; }
            public string? GenderCol { get; set; }
            public string? CurrentResidenceCol { get; set; }
            public string? PermanentResideceCol { get; set; }
            public string? BirthPlaceCol { get; set; }
            public string? FatherNameCol { get; set; }
            public string? MotherNameCol { get; set; }
            public string? PhoneNumberCol { get; set; }
            public string[] SheetType { get; set; }
            public StudentCol(string rowIndex, string[] sheetType, string? stt, string? firstName, 
                string? lastName, string? _class, 
                string? birth, string? gender, string? currentResidence, string? permanentResidece, 
                string? birthPlace, string? fatherName, string? motherName, string? phoneNumber)
            {
                RowIndex = rowIndex;
                SheetType = sheetType;
                SttCol = stt + RowIndex;
                FirstNameCol = firstName + RowIndex;
                LastNameCol = lastName + RowIndex;
                ClassCol = _class + RowIndex;
                BirthCol = birth + RowIndex;
                GenderCol = gender + RowIndex;
                CurrentResidenceCol = currentResidence + RowIndex;
                PermanentResideceCol = permanentResidece + RowIndex;
                BirthPlaceCol = birthPlace + RowIndex;
                FatherNameCol = fatherName + RowIndex;
                MotherNameCol = motherName + RowIndex;
                PhoneNumberCol = phoneNumber + RowIndex;
            }
        }
        public static Student AssignStudent(StudentCol studentCol, ExcelWorksheet worksheet)
        {
            Student student = new Student();

            student.Name = string.Concat(worksheet.Cells[studentCol.FirstNameCol].Text, " ", worksheet.Cells[studentCol.LastNameCol].Text);
            student.Class = worksheet.Name.ToString();
            //birth
            try
            {
                if (DateTime.TryParseExact(
                    worksheet.Cells[studentCol.BirthCol].Text,
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
                student.Birth = DateTime.Now;
            }
            student.Gender = worksheet.Cells[studentCol.GenderCol].Text;
            student.CurrentResidence = worksheet.Cells[studentCol.CurrentResidenceCol].Text;
            student.PermanentResidece = worksheet.Cells[studentCol.PermanentResideceCol].Text;
            student.BirthPlace = worksheet.Cells[studentCol.BirthPlaceCol].Text;
            student.FatherName = worksheet.Cells[studentCol.FatherNameCol].Text;
            student.MotherName = worksheet.Cells[studentCol.MotherNameCol].Text;
            student.PhoneNumber = worksheet.Cells[studentCol.PhoneNumberCol].Text.Count() == 9 ? string.Concat("0", worksheet.Cells[studentCol.PhoneNumberCol].Text) : worksheet.Cells[studentCol.PhoneNumberCol].Text;
            return student;
        }
        public IActionResult Error()
        {
            return View();
        }
    }
}