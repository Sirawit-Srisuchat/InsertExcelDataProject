using OfficeOpenXml;
using saveText_to_Excel.AllClass;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace saveText_to_Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            // ที่อยู่ไฟล์ Excel
            string filePath = "D:\\InsertDataToExcel.xlsx";

            // Initialize a list to hold existing data
            // สร้างรายการเพื่อเก็บข้อมูลที่มีอยู่
            List<(string Date, string Task, string Total)> existingData = System.IO.File.Exists(filePath)
                ? ExcelHelper.ReadExistingData(filePath)
                : new List<(string Date, string Task, string Total)>();

            // Prompt the user for the number of new entries to add
            // ร้องขอจำนวนของรายการใหม่ที่ต้องการเพิ่ม
            Console.Write("Please enter the number of rows you want to save into the Excel file: ");
            int numNewEntries = int.Parse(Console.ReadLine());

            // Collect data for each new entry
            // รวบรวมข้อมูลสำหรับแต่ละรายการใหม่
            List<(string Date, string Task, string Total)> newData = new List<(string Date, string Task, string Total)>();
            for (int i = 0; i < numNewEntries; i++)
            {
                // Use the current date for each new entry
                // ใช้วันที่ปัจจุบันสำหรับแต่ละรายการใหม่
                string date = DateTime.Now.ToString("dd-MMM-yyyy");

                // Prompt the user for the task for this entry
                // ร้องของผู้ใช้สำหรับงานของรายการนี้
                Console.Write($"Please input Task into row {i + 1}: ");
                string task = Console.ReadLine();

                // Prompt the user for the total amount spent for this entry
                // ร้องของผู้ใช้สำหรับยอดรวมที่ใช้สำหรับรายการนี้
                Console.Write($"Please input Total money spent into row {i + 1}: ");
                string total = Console.ReadLine();

                // Add the new entry to the list
                // เพิ่มรายการใหม่เข้าไปในรายการ
                newData.Add((date, task, total));
            }

            // Check if the Excel file is currently open by another application
            // ตรวจสอบว่าไฟล์ Excel ถูกเปิดอยู่โดยแอปพลิเคชันอื่นหรือไม่
            if (FileHelper.IsFileOpen(filePath))
            {
                // If the file is open, notify the user to close it and try again
                // หากไฟล์ถูกเปิดอยู่ แจ้งให้ผู้ใช้ปิดไฟล์และลองอีกครั้ง
                Console.WriteLine("The Excel file is currently open. Please close it and try again.");
            }
            else
            {
                // Merge existing data with new data
                // ผสานข้อมูลที่มีอยู่กับข้อมูลใหม่
                List<(string Date, string Task, string Total)> mergedData = new List<(string Date, string Task, string Total)>(existingData);
                mergedData.AddRange(newData);

                // Write the merged data to the Excel file
                // เขียนข้อมูลที่ผสานเข้าไฟล์ Excel
                ExcelWriter.WriteDataToExcel(filePath, mergedData);

                // Confirm successful data writing
                // ยืนยันการเขียนข้อมูลสำเร็จ
                Console.WriteLine("Data merged and written to Excel successfully at the specified location: " + filePath);
            }
        }
    }
}