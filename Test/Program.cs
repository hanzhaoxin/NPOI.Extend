using NPOI.Extend;
using System;
using System.Drawing;
using System.IO;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var w = NPOIHelper.LoadWorkbook(@"files\test.xlsx");
            w.GetSheet("Sheet1").GetRow(0).GetCell(2).SetValue(Image.FromFile(@"image\C#高级编程.jpg"));
            using (FileStream fs = File.OpenWrite(@"files\out.xlsx"))
            {
                byte[] buffer = w.SaveToBuffer();
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
            Console.WriteLine("Hello World!");
        }
    }
}
