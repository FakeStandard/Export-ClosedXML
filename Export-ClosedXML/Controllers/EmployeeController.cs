using Export_ClosedXML.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;

namespace Export_ClosedXML.Controllers
{
    public class EmployeeController : Controller
    {
        private List<Employee> employees;

        public EmployeeController()
        {
            if (employees == null)
            {
                employees = new List<Employee>()
                {
                    new Employee{ ID = 1, LastName = "Davolio", FirstName = "Nancy", Title = "Sales Representative", City = "Seattle" },
                    new Employee{ ID = 2, LastName = "Fuller", FirstName = "Andrew", Title = "Vice President, Sales", City = "Tacoma" },
                    new Employee{ ID = 3, LastName = "Leverling", FirstName = "Janet", Title = "Sales Representative", City = "Kirkland" },
                    new Employee{ ID = 4, LastName = "Peacock", FirstName = "Margaret", Title = "Sales Representative", City = "Redmond" },
                    new Employee{ ID = 5, LastName = "Buchanan", FirstName = "Steven", Title = "Sales Manager", City = "London" },
                    new Employee{ ID = 6, LastName = "Suyama", FirstName = "Michael", Title = "Sales Representative", City = "London" },
                    new Employee{ ID = 7, LastName = "King", FirstName = "Robert", Title = "Sales Representative", City = "London" },
                    new Employee{ ID = 8, LastName = "Callahan", FirstName = "Laura", Title = "Inside Sales Coordinator", City = "Seattle" },
                    new Employee{ ID = 9, LastName = "Dodsworth", FirstName = "Anne", Title = "Sales Representative", City = "London" }
                };
            }
        }
        public IActionResult Index()
        {
            return View(employees);
        }

        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult Export()
        {
            try
            {
                // 取得類別的欄位名稱
                var headerList = typeof(Employee).GetProperties().Select(m => m.Name).ToList();

                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //string contentType = "application/vnd.ms-excel";
                string fileName = "EmployeeExport.xlsx";

                // 建立工作簿
                IXLWorkbook wb = new XLWorkbook();

                // 建立工作表
                IXLWorksheet sheet = wb.Worksheets.Add("Employee");

                // 合併儲存格
                sheet.Range(1, 1, 1, headerList.Count()).Merge();
                sheet.Cell(1, 1).Value = "Employee Report";

                // 樣式-背景色
                sheet.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.AppleGreen);
                // 字體大小
                sheet.Cell(1, 1).Style.Font.SetFontSize(12);
                // 粗體
                sheet.Cell(1, 1).Style.Font.SetBold();

                // 水平垂直對齊方式
                sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // 標題列寫入 / ClosedXML 的 row 或 cell 都是從 1 開始
                for (int i = 1; i <= headerList.Count(); i++)
                {
                    sheet.Cell(2, i).Value = headerList[i - 1];

                    // 上框線
                    sheet.Cell(2, i).Style.Border.SetTopBorder(XLBorderStyleValues.Double);
                }

                // 內容寫入
                for (int i = 1; i <= employees.Count(); i++)
                {
                    sheet.Cell(i + 2, 1).Value = employees[i - 1].ID;
                    sheet.Cell(i + 2, 2).Value = employees[i - 1].FirstName;
                    sheet.Cell(i + 2, 3).Value = employees[i - 1].LastName;
                    sheet.Cell(i + 2, 4).Value = employees[i - 1].Title;
                    sheet.Cell(i + 2, 5).Value = employees[i - 1].City;
                }

                // 自適應欄寬
                sheet.Columns().AdjustToContents();

                using (MemoryStream ms = new MemoryStream())
                {
                    // 將檔案存入記憶流
                    wb.SaveAs(ms);

                    // 記憶流轉換成 byte[]
                    var content = ms.ToArray();

                    return File(content, contentType, fileName);
                }
            }
            catch (Exception ex)
            {
                return File(System.Text.Encoding.Unicode.GetBytes(ex.Message), "application/x-unknown", "error.txt");
            }
        }
    }
}
