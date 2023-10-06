using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace saveText_to_Excel.AllClass
{
    internal static class ExcelHelper
    {
        // Read existing data from the Excel file
        // อ่านข้อมูลที่มีอยู่จากไฟล์ Excel
        public static List<(string Date, string Task, string Total)> ReadExistingData(string filePath)
        {
            List<(string Date, string Task, string Total)> existingData = new List<(string Date, string Task, string Total)>();

            // Load the Excel package from the specified file
            // โหลด Excel package จากไฟล์ที่ระบุ
            using (ExcelPackage existingExcelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet existingWorksheet = existingExcelPackage.Workbook.Worksheets[0];

                int rows = existingWorksheet.Dimension.Rows;

                // Read data from each row starting from the second row (header is in the first row)
                // อ่านข้อมูลจากแต่ละแถว เริ่มจากแถวที่สอง (หัวเรื่องอยู่ในแถวแรก)
                for (int i = 2; i <= rows; i++)
                {
                    // Read date, task, and total from each row
                    // อ่านวันที่, งาน, และยอดรวมจากแต่ละแถว
                    string date = existingWorksheet.Cells[i, 1].Value?.ToString() ?? "";
                    string task = existingWorksheet.Cells[i, 2].Value?.ToString() ?? "";
                    string total = existingWorksheet.Cells[i, 3].Value?.ToString() ?? "";

                    // Add the read data to the list
                    // เพิ่มข้อมูลที่อ่านลงในรายการ
                    existingData.Add((date, task, total));
                }
            }

            return existingData;
        }
    }
}
