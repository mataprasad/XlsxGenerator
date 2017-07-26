using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;

namespace XlsxGenerator
{
    public class Generator
    {
        public static void Generate(string tempTemplateFilePath, List<string> columnList, DataTable dataToWrite, bool writeHeader)
        {
            //1. Create temp copy of template

            using (Stream s = Helper.GetEmbeddedResourceAsFileStream(Helper.TEMPLATE_RESOURCE_NAME))
            {
                using (FileStream fs = File.Create(tempTemplateFilePath))
                {
                    Helper.CopyStream(s, fs);
                }
            }

            //2. Generate sheet1.xml with the data

            string tempSheetFilePath = Path.Combine(Path.GetDirectoryName(tempTemplateFilePath), Guid.NewGuid().ToString() + ".xml");
            StreamWriter sw = File.CreateText(tempSheetFilePath);
            sw.AutoFlush = true;
            sheetData sheet = new sheetData();

            using (sw)
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sw.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
                sw.Write("<sheetPr/>");
                sw.Write("<dimension ref=\"A1:CN2\"/>");
                sw.Write("<sheetViews>");
                sw.Write("<sheetView tabSelected=\"1\" workbookViewId=\"0\">");
                sw.Write("<selection activeCell=\"A1\" sqref=\"A1\"/>");
                sw.Write("</sheetView>");
                sw.Write("</sheetViews>");
                sw.Write("<sheetFormatPr defaultColWidth=\"9.14285714285714\" defaultRowHeight=\"13\" customHeight=\"1\" outlineLevelRow=\"1\"/>");
                
                if (writeHeader)
                {
                    sheet.row = new sheetDataRow[dataToWrite.Rows.Count+1];
                    sheet.row[0] = new sheetDataRow();
                    sheet.row[0].c = new sheetDataRowC[columnList.Count];
                    for (int j = 0; j < columnList.Count; j++)
                    {

                        sheet.row[0].c[j] = new sheetDataRowC();
                        sheet.row[0].c[j].v = columnList[j];
                    }
                }else
                {
                    sheet.row = new sheetDataRow[dataToWrite.Rows.Count];
                }
                for (int i = 1; i < sheet.row.Length-1; i++)
                {

                    sheet.row[i] = new sheetDataRow();
                    sheet.row[i].c = new sheetDataRowC[columnList.Count];
                    for (int j = 0; j < columnList.Count; j++)
                    {

                        sheet.row[i].c[j] = new sheetDataRowC();
                        sheet.row[i].c[j].v = Convert.ToString(dataToWrite.Rows[i][columnList[j]]);
                    }
                }
                sw.Write(Helper.SerializeToXML<sheetData>(sheet));
                sw.Write("<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.511805555555556\" footer=\"0.511805555555556\"/>");
                sw.Write("<headerFooter/>");
                sw.Write("</worksheet>");
            }

            //3. Add the generated sheet.xml file to existing .xlsx template file

            Helper.AddFileToExistingZip(tempTemplateFilePath, tempSheetFilePath);

            //3. Delete the temp sheet.xml
            File.Delete(tempSheetFilePath);
        }
    }
}
