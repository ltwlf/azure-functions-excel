using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class ExcelService : IDisposable
{
    private SpreadsheetDocument spreadsheet;
    private SharedStringTablePart sharedStringTable;
    private WorkbookPart workbookPart;
    private MemoryStream stream;

    public ExcelService(string excelAsBase64)
    {
        var binExcel = Convert.FromBase64String(excelAsBase64);

        stream = new MemoryStream(binExcel);

        spreadsheet = SpreadsheetDocument.Open(stream, true);

        sharedStringTable = null;
        if (spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            sharedStringTable = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }

        workbookPart = spreadsheet.WorkbookPart;
    }

    public string ReplaceTokens(string template)
    {
        var tokenRegex = new Regex(@"\$\{(.*?)!(.*?)\}");
        var tokenMatches = tokenRegex.Matches(template);

        foreach (Match token in tokenMatches)
        {
            var tokenString = token.Groups[0].Value;
            var sheetName = token.Groups[1].Value;
            var cellRef = token.Groups[2].Value;

            var sheet = workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == sheetName).FirstOrDefault();

            if(sheet != null)
            {
                var wsPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                var value = GetCellValue(wsPart, cellRef);

                template = template.Replace(tokenString, value);
            }
        }

        return template;
    }

    private string GetCellValue(WorksheetPart wsPart, string cellRef)
    {
        var cell = wsPart.Worksheet.Descendants<Cell>().
                            Where(c => c.CellReference == cellRef).FirstOrDefault();

        var value = "";
        // If the cell does not exist, return an empty string.
        if (cell != null)
        {
            value = cell.InnerText;

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:

                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (sharedStringTable != null)
                        {
                            value =
                                sharedStringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
        }

        return value;
    }

    public void Dispose()
    {
        if(spreadsheet != null)
        {
            spreadsheet.Dispose();
        }

        if(stream != null)
        {
            stream.Dispose();
        }
    }
}