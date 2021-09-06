using System;
using System.IO;
using OfficeOpenXml;
namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Users\MauricioIvanSalazarM\Documents\C#Tests\CORTES.xlsx");

            using var package = new ExcelPackage(file);
            var worksheet = package.Workbook.Worksheets[0];
            Console.WriteLine("Worksheet name: " + worksheet.ToString());
            //Console.WriteLine("Value in cell G4: " + worksheet.Cells["G19"].Value.ToString());

            string textInColumn = "test";
            int indexOfLastRowWithText = 1;
            
            while (true)
            {
                var temp = worksheet.Cells["G" + indexOfLastRowWithText.ToString()].Value;
                if (temp != null)
                {
                    textInColumn = temp.ToString();
                    Console.WriteLine(textInColumn);
                    indexOfLastRowWithText++;
                    Console.WriteLine(indexOfLastRowWithText);
                }
                else
                {
                    indexOfLastRowWithText--;
                    break;
                }
            }
            Console.WriteLine("Last row with text: " + indexOfLastRowWithText);
            //using (StreamWriter sw = File.AppendText(@"C:\Users\MauricioIvanSalazarM\Documents\Tickets de Cortes 2.txt"))
            //{
            //    sw.WriteLine("\n------------------------------------ " + DateTime.Today.ToString("dddd, dd MMM y"));
            //    sw.WriteLine("This is the new text");
            //}

        }
    }
}
