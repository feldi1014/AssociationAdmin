using System.Data;
using ClosedXML.Excel;
using ExcelHelper;

Console.WriteLine("Split orders");
Console.WriteLine("Usage Examle Bestellungen.xlsx Mitglieder.xlsx");
Console.WriteLine("Usage Examle Bestellungen.xlsx Mitglieder.xlsx 2024_Mitgliederkartenvergabe_final.xls");

if (args.Length == 0)
{
    Console.WriteLine("Missing excel file!");
    return;
}

var orderFileName = new FileInfo(args[0]);
if (!orderFileName.Exists)
{
    Console.WriteLine($"Missing excel file at {orderFileName.FullName}!");
    return;
}

var helperOrders = new DataSetHelper(orderFileName);


if (!helperOrders.GetDataSet())
{
    Console.WriteLine("Getting dataset Bestellungen fehlgeschlagen!");
    return;
}

helperOrders.UpdateRowContent();

helperOrders.GetCardsAndGroups();

helperOrders.OptimizeOrders();

helperOrders.ValidateOrders();

if (args.Length >= 2)
{
    AppendMemberData(new FileInfo(args[1]), helperOrders);
}

if (args.Length >= 3)
{
    AppendLastYearData(new FileInfo(args[2]), helperOrders);
}


using (var workbook = new XLWorkbook())
{
    Console.WriteLine();
    Console.WriteLine("Bestellungen gesplittet:");
    var builder = new WorksheetBuilder(helperOrders.Data, workbook);

    builder.ExportOrderSplit();

    workbook.SaveAs("SplitBestellungen.xlsx");
}

Console.Write("Press any key");
Console.ReadKey();



static void AppendMemberData(FileInfo memberFileName, DataSetHelper helperOrders)
{
    DataSetHelper helperMembers = null;
    if (!memberFileName.Exists)
    {
        Console.WriteLine($"Missing excel file at {memberFileName.FullName}!");
    }
    else
    {
        helperMembers = new DataSetHelper(memberFileName);
        if (!helperMembers.GetDataSet())
        {
            Console.WriteLine("Getting dataset Mitglieder fehlgeschlage!");
        }
        else
        {
            foreach (DataTable dt in helperMembers.Data.Tables)
            {
                helperOrders.Data.Tables.Add(dt.Copy());
            }
            helperMembers.SetPrimaryKey("Mitglieder", "angelflix ID");
            helperOrders.AppendColumnMitgliedsJahre("Mitglieder");
            helperOrders.AppendColumnMitgliedsJahre("Bestellungen");
        }
    }
}

static void AppendLastYearData(FileInfo lastYearFileName, DataSetHelper helperOrders)
{
    DataSetHelper helperLastYear = null;
    if (!lastYearFileName.Exists)
    {
        Console.WriteLine($"Missing excel file at {lastYearFileName.FullName}!");
    }
    else
    {
        helperLastYear = new DataSetHelper(lastYearFileName);
        if (!helperLastYear.GetDataSet())
        {
            Console.WriteLine("Getting dataset last year fehlgeschlage!");
        }
        else
        {
            foreach (DataTable dt in helperLastYear.Data.Tables)
            {
                var copyTable = dt.Copy();
                copyTable.TableName = "LastYear";

                helperOrders.Data.Tables.Add(copyTable);
            }
            helperOrders.AppendColumnsForLastYear();
        }
    }
}
