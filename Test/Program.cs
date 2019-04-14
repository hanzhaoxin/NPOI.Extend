using NPOI.Extend;
using NPOI.Extend.Extend.CellExtend;
using NPOI.Extend.Formula;
using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var w = NPOIHelper.LoadWorkbook(@"files\test.xlsx");
                w.GetSheet("Sheet1").CopyRows(0, 4);

                using (FileStream fs = File.OpenWrite(@"files\out.xlsx"))
                {
                    byte[] buffer = w.SaveToBuffer();
                    fs.Write(buffer, 0, buffer.Length);
                    fs.Flush();
                }

                //string str = w.GetSheet("Sheet1").GetRow(0).GetCell(3).CellFormula;
                //Console.WriteLine(str);
                //FormulaContext context = new FormulaContext(str);
                //context.Parse();
                //foreach (var part in context.Formula)
                //{
                //    if (part.Type.Equals(PartType.Formula))
                //    {
                //        Regex r = new Regex(@"([A-Z]+)(\d+)");
                //        var ms = r.Matches(part.ToString());
                //        foreach (Match m in ms)
                //        {
                //            Console.WriteLine($"row:{m.Groups[2]},col:{m.Groups[1]}");
                //        }
                //        Console.WriteLine(part);
                //        Console.WriteLine(r.Replace(part.ToString(), (m)=>  $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value) + 5}"));
                //    }




                //}
                //Console.WriteLine(context.Formula.ToString());
                //Console.WriteLine(context.Formula.ToString(part => {
                //    if (part.Type.Equals(PartType.Formula))
                //    {
                //        Regex regex = new Regex(@"([A-Z]+)(\d+)");
                //        return regex.Replace(part.ToString(), (m) => $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value) + 5}");
                //    }
                //    else
                //    {
                //        return part.ToString();
                //    }
                //}));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.WriteLine("Hello World!");
            Console.ReadKey();
        }
    }
}
