// See https://aka.ms/new-console-template for more information

using System.Data;
using ClosedXML.Excel;
using OfficeOpenXml;

bool exit = false;
do 
{
    Console.WriteLine("Enter Flex Localization directory: ");

    string directory = Console.ReadLine();

    if (!Directory.Exists(directory))
    {
        Console.WriteLine("Invalid directory");
        continue;
    }

    var files = Directory.GetFiles(directory, "*.properties");

    if (files.Length <= 0)
    {
        Console.WriteLine("No *.properties files found");
        continue;
    }
    
    DataTable data = new DataTable();
    data.Columns.Add("Source", typeof(string));
    data.Columns.Add("Key", typeof(string));
    data.Columns.Add("Value", typeof(string));

    foreach (var file in files)
    {
        var info = new FileInfo(file);
        var lines = File.ReadAllLines(file);
        
        lines = lines
            .Skip(1)
            .Where(x => !String.IsNullOrWhiteSpace(x)
                && !x.TrimStart().StartsWith("#"))
            .ToArray();

        string source = info.Name.Replace(".properties", "");

        foreach (var line in lines)
        {
            string[] values = line.Split("=");

            if (values.Length < 2)
                continue;

            var row = data.NewRow();
            row["Source"] = source;
            row["Key"] = values[0].Trim();
            row["Value"] = values[1].Trim();
            data.Rows.Add(row);
        }
    }

    using XLWorkbook wb = new XLWorkbook();
    wb.Worksheets.Add(data,"en_US");
    wb.SaveAs($"FlexLocale_{ DateTime.Now.ToString("yyyyMMddhhmmss") }.xlsx");

    wb.Dispose();

    // using var stream = new MemoryStream();
    // using var package = new ExcelPackage(stream); 


    // var workSheet = package.Workbook.Worksheets.Add("sheetName");
    // workSheet.Cells.LoadFromDataTable(data, true);
    // package.Save();
    // stream.Position = 0;

    // package.SaveAs($"FlexLocale_{ DateTime.Now.ToString("yyyyMMddhhmmss") }");

    // package.Dispose(); 
    // stream.Dispose();
    //C:\Workspace\TKS\AML\LIA_FLEX\LIA-Core\src\Locale\en_US
}
while (exit != true);