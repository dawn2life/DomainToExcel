using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DomainToExcel.Domain;

var domainModels = new List<Type>()
{
    typeof(Audit),
    typeof(Status),
    typeof(Source),
    typeof(Base),
    typeof(Domain),
    typeof(DirectionReference),
    typeof(AddressReference),
    typeof(Hours)
};

var workBook = new XLWorkbook();
var workSheet = workBook.Worksheets.Add("Domain Schema");

workSheet.Cell(1, 1).Value = "Class Name";
workSheet.Cell(1, 2).Value = "Property Name";
workSheet.Cell(1, 3).Value = "Property Type";
workSheet.Row(1).Style.Font.Bold = true;

int row = 2;

foreach (var type in domainModels) 
{
    var properties = type.GetProperties();
    foreach (var property in properties) 
    {
        workSheet.Cell(row, 1).Value = type.Name;
        workSheet.Cell(row, 2).Value = property.Name;
        workSheet.Cell(row, 3).Value = GetFriendlyTypeName(property.PropertyType);
        row++;
    }
}

var range = workSheet.Range(1, 1, row - 1, 3);
range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
workSheet.Columns().AdjustToContents();

workBook.SaveAs("DomainSchema.xlsx");
Console.WriteLine("Excel file created successfully");

#region Helper Methods

static string GetFriendlyTypeName(Type type)
{
    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
    {
        return $"List<{type.GetGenericArguments()[0].Name}>"; // For List<T>
    }

    return type.Name; // For other types
}

#endregion