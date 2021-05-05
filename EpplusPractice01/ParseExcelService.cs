using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace EpplusPractice01
{
    public class ParseExcelService
    {
        public List<Dto> Parse(string xlsxFilePath)
        {
            var fileInfo = new FileInfo(xlsxFilePath);
            using (var excel = new ExcelPackage(fileInfo))
            {
                var sheet = excel.Workbook.Worksheets.FirstOrDefault();


                var result = ParseSheet(sheet);

                return result;
            }
        }

        private List<Dto> ParseSheet(ExcelWorksheet sheet)
        {
            var lastRow = sheet.Dimension.End.Row;
            var result  = new List<Dto>();

            // 扣掉 Title
            for (int rowIndex = 2; rowIndex < lastRow; rowIndex++)
            {
                var item = new Dto();

                item.BB    = sheet.Cells[rowIndex, 1].GetValue<string>();
                item.plant = sheet.Cells[rowIndex, 2].GetValue<string>();
                item.Model = sheet.Cells[rowIndex, 3].GetValue<string>();
                item.part  = sheet.Cells[rowIndex, 4].GetValue<string>();
                item.pm    = sheet.Cells[rowIndex, 5].GetValue<string>();
                item.MRP   = sheet.Cells[rowIndex, 6].GetValue<string>();
                item.NG    = sheet.Cells[rowIndex, 8].GetValue<string>();
                item.Total = sheet.Cells[rowIndex, 10].GetValue<string>();
                item.type  = sheet.Cells[rowIndex, 11].GetValue<string>();
                item.May   = sheet.Cells[rowIndex, 12].GetValue<string>();
                item.MTDA  = sheet.Cells[rowIndex, 13].GetValue<string>();
                item.Jun   = sheet.Cells[rowIndex, 14].GetValue<string>();
                result.Add(item);
            }

            return result;
        }
    }
}
