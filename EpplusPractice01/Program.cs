using System;
using System.IO;
using System.Reflection;

namespace EpplusPractice01
{
    class Program
    {
        static void Main(string[] args)
        {
            var xlsxFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                            "xlsxs",
                                            "1+2xlsx.xlsx");

            var dtos = new ParseExcelService().Parse(xlsxFilePath);

            foreach (var dto in dtos)
            {
                Console.WriteLine(dto.ToString());
            }
        }
    }
}
