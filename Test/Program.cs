using NPOI.Extend;
using NPOI.Extend.Extend.CellExtend;
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
            w.GetSheet("Sheet1").GetRow(0).GetCell(0).Merge(w.GetSheet("Sheet1").GetRow(0).GetCell(2));

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
