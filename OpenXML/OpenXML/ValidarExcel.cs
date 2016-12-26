using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics;

namespace OpenXML
{
    public class ValidarExcel
    {
        public static void validarExcel(string path)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            using (SpreadsheetDocument excelFile = SpreadsheetDocument.Open(path, true))
            {
                try
                {
                    int count = 0;
                    foreach (ValidationErrorInfo error in validator.Validate(excelFile))
                    {
                        count++;
                        Trace.WriteLine("Error " + count);
                        Trace.WriteLine("Description: " + error.Description);
                        Trace.WriteLine("Path: " + error.Path.XPath);
                        Trace.WriteLine("Part: " + error.Part.Uri);
                        Trace.WriteLine("-------------------------------------");
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
                excelFile.Close();
            }
        }
    }
}
