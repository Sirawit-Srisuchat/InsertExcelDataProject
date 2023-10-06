using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace saveText_to_Excel.AllClass
{
    internal static class ExcelWriter
    {
        // Write data to an Excel file
        // เขียนข้อมูลลงในไฟล์ Excel
        public static void WriteDataToExcel(string filePath, List<(string Date, string Task, string Total)> data)
        {
            // Create a new Excel package
            // สร้าง Excel package ใหม่
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a worksheet to the Excel package
                // เพิ่มแผ่นงานใน Excel package
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Set the header row
                // กำหนดหัวแถว
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Task";
                worksheet.Cells[1, 3].Value = "Total";

                // Write the data to the worksheet
                // เขียนข้อมูลลงในแผ่นงาน
                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = data[i].Date;
                    worksheet.Cells[i + 2, 2].Value = data[i].Task;
                    worksheet.Cells[i + 2, 3].Value = data[i].Total;
                }

                // Save the Excel package to the specified file path
                // บันทึก Excel package ลงในที่ระบุ
                FileInfo fileInfo = new FileInfo(filePath);
                excelPackage.SaveAs(fileInfo);
            }
        }
    }
}
