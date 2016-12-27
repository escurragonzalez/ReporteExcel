using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OpenXML
{
    public class Report
    {
        public byte[] CreateExcelDoc(List<InscritoDTO> inscriptos)
        {            
            using (MemoryStream mem = new MemoryStream())
            {
                SpreadsheetDocument document = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Inscriptos" };

                sheets.Append(sheet);
                worksheetPart.Worksheet = new Worksheet();
                workbookPart.Workbook.Save();

                SheetData sheetData = new SheetData();

                worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:I1" };

                SheetView sheetView = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
                SheetFormatProperties sheetFormatoperties = worksheetPart.Worksheet.AppendChild(new SheetFormatProperties() { BaseColumnWidth = (UInt32Value)10U, DefaultRowHeight = 15D });

                #region Cabecera

                #region Definicion de Columnas
                Columns columns1 = worksheetPart.Worksheet.AppendChild(new Columns());

                Column column1 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 21D, CustomWidth = true };
                Column column2 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 21D, CustomWidth = true };
                Column column3 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 25D, CustomWidth = true };
                Column column4 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15D, CustomWidth = true };
                Column column5 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 18D, CustomWidth = true };
                Column column6 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 16D, CustomWidth = true };

                columns1.Append(column1);
                columns1.Append(column2);
                columns1.Append(column3);
                columns1.Append(column4);
                columns1.Append(column5);
                columns1.Append(column6);
                #endregion

                Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:9" } };

                Cell cell1 = new Cell() { CellReference = "A1", DataType = CellValues.String };
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "Id";

                cell1.Append(cellValue1);

                Cell cell2 = new Cell() { CellReference = "B1", DataType = CellValues.String };
                CellValue cellValue2 = new CellValue();
                cellValue2.Text = "Nro de Socio";

                cell2.Append(cellValue2);

                Cell cell3 = new Cell() { CellReference = "C1", DataType = CellValues.String };
                CellValue cellValue3 = new CellValue();
                cellValue3.Text = "DNI";

                cell3.Append(cellValue3);

                Cell cell4 = new Cell() { CellReference = "D1", DataType = CellValues.String };
                CellValue cellValue4 = new CellValue();
                cellValue4.Text = "Nombre";

                cell4.Append(cellValue4);

                Cell cell5 = new Cell() { CellReference = "E1", DataType = CellValues.String };
                CellValue cellValue5 = new CellValue();
                cellValue5.Text = "Email";

                cell5.Append(cellValue5);

                Cell cell6 = new Cell() { CellReference = "F1", DataType = CellValues.String };
                CellValue cellValue6 = new CellValue();
                cellValue6.Text = "Telefono";

                cell6.Append(cellValue6);

                Cell cell7 = new Cell() { CellReference = "G1", DataType = CellValues.String };
                CellValue cellValue7 = new CellValue();
                cellValue7.Text = "Fecha de Registro";

                cell7.Append(cellValue7);

                Cell cell8 = new Cell() { CellReference = "H1", DataType = CellValues.String };
                CellValue cellValue8 = new CellValue();
                cellValue8.Text = "Lugar";

                cell8.Append(cellValue8);

                Cell cell9 = new Cell() { CellReference = "I1", DataType = CellValues.String };
                CellValue cellValue9 = new CellValue();
                cellValue9.Text = "Horario";

                cell9.Append(cellValue9);

                row1.Append(cell1);
                row1.Append(cell2);
                row1.Append(cell3);
                row1.Append(cell4);
                row1.Append(cell5);
                row1.Append(cell6);
                row1.Append(cell7);
                row1.Append(cell8);
                row1.Append(cell9);

                #endregion

                sheetData.AppendChild(row1);
                cargarDatos(inscriptos, sheetData);
                worksheetPart.Worksheet.AppendChild(sheetData);
                PageSetup pageSetup1 = worksheetPart.Worksheet.AppendChild(new PageSetup() { Orientation = OrientationValues.Portrait, Id = "rId1" });                
                worksheetPart.Worksheet.Save();
                mem.Seek(0, SeekOrigin.Begin);
                document.Close();
                return mem.ToArray();
            }
        }

        private void cargarDatos(List<InscritoDTO> inscriptos, SheetData sheetData)
        {
            Row row = null;
            foreach (var inscripto in inscriptos)
            {
                row = new Row();

                row.Append(
                    ConstructCell(inscripto.Id.ToString(), CellValues.Number),
                    ConstructCell(inscripto.Num_socio.ToString(), CellValues.Number),
                    ConstructCell(inscripto.Dni.ToString(), CellValues.Number),
                    ConstructCell(inscripto.Nombre.ToString(), CellValues.String),
                    ConstructCell(inscripto.Email.ToString(), CellValues.String),
                    ConstructCell(inscripto.Telefono.ToString(), CellValues.String),
                    ConstructCell(inscripto.Fecha_reg.ToString(), CellValues.String),
                    ConstructCell(inscripto.Lugar.ToString(), CellValues.String),
                    ConstructCell(inscripto.Horario.ToString(), CellValues.String));

                sheetData.AppendChild(row);
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
