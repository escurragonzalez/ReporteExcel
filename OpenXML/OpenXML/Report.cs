using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXML
{
    public class Report
    {
        public void CreateExcelDoc(string file, List<Employee> employees)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(file, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Employees" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();


                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String),
                    ConstructCell("Name", CellValues.String),
                    ConstructCell("Birth Date", CellValues.String),
                    ConstructCell("Salary", CellValues.String));

                sheetData.AppendChild(row);

                foreach (var employee in employees)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(employee.Id.ToString(), CellValues.Number),
                        ConstructCell(employee.Name, CellValues.String),
                        ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String),
                        ConstructCell(employee.Salary.ToString(), CellValues.Number));

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}