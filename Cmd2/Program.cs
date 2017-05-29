using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cmd2
{
    class Program
    {
        private static List<List<object>> _data = GetData();

        static void Main(string[] args)
        {
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            string _basePath = Path.Combine(tempBase, Guid.NewGuid().ToString()+".xlsx");
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_basePath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            foreach (var row in _data)
            {
                Row newRow = new Row();

                foreach (var value in row)
                {
                    Cell cell = new Cell();
                    if (value is DateTime)
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
                    else if (value is decimal || value is int || value is double || value is long)
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    else
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;

                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.ToString());
                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }
            
            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

        }


        private static List<List<object>> GetData()
        {
            List<List<object>> data = new List<List<object>>();

            var random = new Random();
            for (int i = 0; i < 100; i++)
            {
                var x = random.Next(0, 10);
                var y = Math.Round(random.NextDouble(), 2, MidpointRounding.AwayFromZero);
                data.Add(new List<object>
                {
                    x,"CD BLABLA BLA BLA BLA "+i, x,"REDE BLA BLA BLA BLA"+i,x,"BRAHMA LATA 350 BLA BLA BLACX12"+i,y,y,y,y,y,y
                });
            }

            return data;
        }
    }
}
