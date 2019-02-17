using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
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

        stream = new MemoryStream();
        stream.WriteAsync(binExcel);

        spreadsheet = SpreadsheetDocument.Open(stream, true);

        sharedStringTable = null;
        if (spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        {
            sharedStringTable = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        }

        workbookPart = spreadsheet.WorkbookPart;
    }

    public string GetCellData(string template)
    {
        var tokenRegex = new Regex(@"\$\{(.*?)!(.*?)\}");
        var tokenMatches = tokenRegex.Matches(template);

        foreach (Match token in tokenMatches)
        {
            var tokenString = token.Groups[0].Value;
            var sheetName = token.Groups[1].Value;
            var cellRef = token.Groups[2].Value;

            var sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

            if (sheet != null)
            {
                var wsPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                var value = GetCellValue(wsPart, cellRef);

                template = template.Replace(tokenString, HttpUtility.JavaScriptStringEncode(value));
            }
        }

        return template;
    }

    public string WriteCellData(IDictionary<string, object> data)
    {
        var cellRefRegex = new Regex(@"^([a-zA-Z]*)(\d*)$");

        if (sharedStringTable == null)
        {
            sharedStringTable = workbookPart.AddNewPart<SharedStringTablePart>();
        }

        foreach (var key in data.Keys)
        {
            var sheetName = key.Split("!")[0];
            var cellRef = key.Split("!")[1];

            WorksheetPart wsPart = null;
            var sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (sheet == null)
            {
                wsPart = InsertWorksheet(workbookPart, sheetName);
            }
            else
            {
                wsPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
            }

            var cellRefMatch = cellRefRegex.Matches(cellRef).FirstOrDefault();
            var columnName = cellRefMatch.Groups[1].Value;
            var rowIndex = (uint)Int32.Parse(cellRefMatch.Groups[2].Value);

            var cell = InsertCellInWorksheet(columnName, rowIndex, wsPart);

            var value = data[key];
            int strIndex = -1;
            if (value is string)
            {
                strIndex = InsertSharedStringItem(value.ToString());
                cell.CellValue = new CellValue(strIndex.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            else
            {
                cell.CellValue = new CellValue(value.ToString());
            }
            wsPart.Worksheet.Save();
        }

        spreadsheet.WorkbookPart.Workbook.Save();
        spreadsheet.Save();
        spreadsheet.Close();
        spreadsheet.Dispose();

        return Convert.ToBase64String(stream.ToArray());

    }

    private WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string name)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new sheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        string sheetName = name;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();

        return newWorksheetPart;
    }


    private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }

    private int InsertSharedStringItem(string text)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (sharedStringTable.SharedStringTable == null)
        {
            sharedStringTable.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in sharedStringTable.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        sharedStringTable.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        sharedStringTable.SharedStringTable.Save();

        return i;
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
        if (spreadsheet != null)
        {
            spreadsheet.Dispose();
        }

        if (stream != null)
        {
            stream.Dispose();
        }
    }
}