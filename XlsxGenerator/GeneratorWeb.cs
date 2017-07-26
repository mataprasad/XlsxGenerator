using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace XlsxGenerator
{
    public static class GeneratorWeb
    {
        public static byte[] Generate(List<string> columnList, DataTable dataToWrite, bool writeHeader)
        {
            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add(Guid.NewGuid().ToString());
            int increment = 1;
            if (writeHeader)
            {
                increment = 2;
                ws.Row(1).Style.Font.Bold = true;
                for (int j = 1; j < columnList.Count + 1; j++)
                {
                    ws.SetValue(1, j, columnList[j - 1]);
                }
            }

            for (int i = 0; i < dataToWrite.Rows.Count; i++)
            {
                for (int j = 1; j < columnList.Count + 1; j++)
                {
                    ws.SetValue(i + increment, j, Convert.ToString(dataToWrite.Rows[i][columnList[j - 1]]));
                }

            }

            return pck.GetAsByteArray();
        }
    }
}
