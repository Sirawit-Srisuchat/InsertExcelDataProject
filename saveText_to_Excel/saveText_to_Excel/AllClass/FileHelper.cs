using System;
using System.Collections.Generic;
using System.IO;

namespace saveText_to_Excel.AllClass
{
    internal static class FileHelper
    {
        // Check if a file is currently open by another application
        // ตรวจสอบว่าไฟล์ถูกเปิดอยู่โดยแอปพลิเคชันอื่นหรือไม่
        public static bool IsFileOpen(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // The file is not in use
                    // ไฟล์ไม่ถูกใช้งาน
                }
                return false;
            }
            catch (IOException)
            {
                // The file is in use
                // ไฟล์ถูกใช้งาน
                return true;
            }
        }
    }
}
