using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.ComponentModel;

namespace ExcelImportExport;

public class ExcelSerializer
{
    private const string DefaultSheetName = "Sheet1";

    public List<T> ImportFromExcel<T>(string filepath) where T : new()
    {
        using var stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Write);
        var result = ImportFromExcel<T>(stream);
        return result;
    }

    public List<T> ImportFromExcel<T>(Stream stream) where T : new()
    {
        using var document = SpreadsheetDocument.Open(stream, false);

        var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
        var sheet = worksheetPart.Worksheet;

        var header = sheet.Descendants<Row>().FirstOrDefault();
        var rows = sheet.Descendants<Row>().Skip(1);

        if (header == null) throw new Exception("No header row found in the Excel file.");

        var headerCells = header.Descendants<Cell>().ToList();
        var result = new List<T>();

        foreach (var row in rows)
        {
            var item = new T();
            var cells = row.Elements<Cell>().ToList();

            for (int i = 0; i < Math.Min(cells.Count, headerCells.Count); i++)
            {
                var propertyName = headerCells[i].CellValue?.Text;
                if (string.IsNullOrWhiteSpace(propertyName)) continue;

                var property = typeof(T).GetProperty(propertyName);
                if (property == null || !property.CanWrite) continue;

                var value = cells[i].CellValue?.Text;
                if (value == null) continue;

                try
                {
                    var converter = TypeDescriptor.GetConverter(property.PropertyType);
                    var convertedValue = converter.ConvertFromInvariantString(value);
                    property.SetValue(item, convertedValue);
                }
                catch
                {
                    throw new FormatException($"Unable to convert value '{value}' to property '{propertyName}' of type '{property.PropertyType.Name}'.");
                }
            }

            result.Add(item);
        }

        return result;
    }

    public void ExportToExcel<T>(List<T> list, string filePath, List<string>? customHeaders = null)
    {
        //using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        using var stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read);
        ExportToExcel(list, stream, customHeaders);
    }

    public void ExportToExcel<T>(List<T> list, Stream stream, List<string>? customHeaders = null)
    {
        if (list == null || !list.Any()) throw new ArgumentException("The list is empty.", nameof(list));

        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = DefaultSheetName };
        sheets.Append(sheet);

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) throw new Exception("Failed to initialize sheet data.");

        var properties = typeof(T).GetProperties().Where(p => p.CanRead).ToList();

        // Header row
        var headerRow = new Row();
        foreach (var header in customHeaders ?? properties.Select(p => p.Name))
        {
            headerRow.Append(CreateCell(header));
        }
        sheetData.Append(headerRow);

        // Data rows
        foreach (var item in list)
        {
            var dataRow = new Row();
            foreach (var prop in properties)
            {
                var value = prop.GetValue(item)?.ToString() ?? string.Empty;
                dataRow.Append(CreateCell(value));
            }
            sheetData.Append(dataRow);
        }
    }

    private static Cell CreateCell(string value)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(value)
        };
    }
}

public class ExcelColumnHelper
{
    public static int ColumnIndex(string reference)
    {
        int index = 0;
        reference = reference.ToUpper();
        for (int i = 0; i < reference.Length && reference[i] >= 'A'; i++)
            index = (index * 26) + (reference[i] - 'A' + 1);
        return index;
    }

    public static string GetColumnName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var name = string.Empty;
        while (index > 0)
        {
            index--;
            name = letters[index % 26] + name;
            index /= 26;
        }
        return name;
    }
}