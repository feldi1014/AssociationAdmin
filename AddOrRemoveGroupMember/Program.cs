using AddOrRemoveGroupMember;
using ClosedXML.Excel;
using ExcelHelper;
using System.Data;

Console.WriteLine("Add or remove group member");
Console.WriteLine("Usage: AddOrRemove <members excel file> -MG:Moos -By:SplitBestellungen.xlsx -ByTable:\"Erwachsene Moos\" -ByField:Wunschkarte -ByContent:Moos");
Console.WriteLine("Usage: AddOrRemove <members excel file> --fees -F:Bestellkostenerstattung -By:SplitBestellungen.xlsx -ByTable:\"Gesamt\"");

if (args.Length == 0)
{
    Console.WriteLine("Missing commandline arguments!");
    return;
}

var settings = new CommandArgs(args);

var memberFileName = new FileInfo(settings.MemberFileName);
if (!memberFileName.Exists)
{
    Console.WriteLine($"Missing excel file at {memberFileName.FullName}!");
    return;
}

var helperMembers = new DataSetHelper(memberFileName);

if (!helperMembers.GetDataSet())
{
    Console.WriteLine("Getting dataset Mitglieder fehlgeschlagen!");
    return;
}

DataTable? tableMembers = null;
if (helperMembers.Data.Tables[0].TableName == "Mitglieder")
{
    tableMembers = helperMembers.Data.Tables[0];
}
else
{
    Console.WriteLine("Missing table Mitglieder!");
    return;
}

if (!string.IsNullOrEmpty(settings.ByExcelFileName))
{
    var byExcelFile = new FileInfo(settings.ByExcelFileName);
    if (byExcelFile.Exists)
    {
        var helperByData = new DataSetHelper(byExcelFile);
        helperByData.GetDataSet();
        if (settings.Command == ToolCommandType.MemberGroup)
        {
            RemoveGroup(settings, tableMembers, helperByData.Data);
            AddGroup(settings, tableMembers, helperByData.Data);
        }
        else if (settings.Command == ToolCommandType.Fees)
        {
            RemoveFees(settings, tableMembers, helperByData.Data);
            AddFees(settings, tableMembers, helperByData.Data);
        }
    }
    else
    {
        Console.WriteLine($"Missing excel file at {byExcelFile.FullName}!");
        return;
    }
}

var changesTable = tableMembers.GetChanges();
if (changesTable == null)
{
    Console.WriteLine("No changes!");
    return;
}

changesTable.TableName = "Mitglieder";
var changesDataSet = new DataSet();
changesDataSet.Tables.Add(changesTable);

using (var workbook = new XLWorkbook())
{
    Console.WriteLine();
    Console.WriteLine("Mitglieder");

    var builder = new WorksheetBuilder(changesDataSet, workbook);

    builder.ExportMitglieder();

    workbook.SaveAs($"{DateTime.Now.ToString("s")}_Mitglieder.xlsx".Replace(":",string.Empty));
}

Console.Write("Press any key");
Console.ReadKey();


static void RemoveFees(CommandArgs settings, DataTable tableMembers, DataSet byData)
{
    foreach (DataRow row in tableMembers.Rows)
    {
        var currentFees = row.Field<string>("Gebühren / Boni");
        var memberNumber = row.Field<string>("Mitgliedsnummer *");

        if (!string.IsNullOrEmpty(currentFees) && currentFees.Contains(settings.Fees))
        {
            foreach (DataTable table in byData.Tables)
            {
                if (table.TableName == settings.ByTableName)
                {
                    DataRow? tableRow = null;
                    if (table.Columns.Contains("Mitgliedsnummer"))
                    {
                        tableRow
                            = table.AsEnumerable().FirstOrDefault(r => r.Field<string>("Mitgliedsnummer") == memberNumber);
                    }
                    if (table.Columns.Contains("Mitgliedsnummer *"))
                    {
                        tableRow
                            = table.AsEnumerable().FirstOrDefault(r => r.Field<string>("Mitgliedsnummer *") == memberNumber);
                    }
                    if (tableRow != null)
                    {
                        row["Gebühren / Boni"] = currentFees.Replace(settings.Fees, string.Empty).Replace(", ,", ", ").Trim();
                    }
                }
            }
        }
    }
}


static void RemoveGroup(CommandArgs settings, DataTable tableMembers, DataSet byData)
{
    foreach (DataRow row in tableMembers.Rows)
    {
        var memberGroups = row.Field<string>("Gruppe(n) *");
        var memberNumber = row.Field<string>("Mitgliedsnummer *");

        if (!string.IsNullOrEmpty(memberGroups) && memberGroups.Contains(settings.MemberGroupName))
        {
            foreach (DataTable table in byData.Tables)
            {
                if (table.TableName == settings.ByTableName)
                {
                    var order
                        = table.AsEnumerable().FirstOrDefault(r => r.Field<string>("Mitgliedsnummer") == memberNumber);
                    if (order != null)
                    {
                        if (order.Field<string>(settings.ByFieldName) == settings.ByContent)
                        {
                            row["Gruppe(n) *"] = memberGroups.Replace(settings.MemberGroupName, string.Empty).Replace(", ,", ", ").Trim();
                        }
                    }
                }
            }
        }
    }
}


static void AddFees(CommandArgs settings, DataTable tableMembers, DataSet byData)
{
    foreach (DataTable table in byData.Tables)
    {
        if (table.TableName == settings.ByTableName)
        {
            foreach (DataRow row in table.Rows)
            {
                var memberNumber = row.Field<string>("Mitgliedsnummer");
                var member = tableMembers.AsEnumerable().FirstOrDefault(r => r.Field<string>("Mitgliedsnummer *") == memberNumber);
                if (member != null)
                {
                    var currentFees = member.Field<string>("Gebühren / Boni");
                    if (string.IsNullOrEmpty(currentFees))
                    {
                        member["Gebühren / Boni"] = settings.Fees;
                    }
                    else
                    {
                        if (!currentFees.Contains(settings.Fees))
                        {
                            member["Gebühren / Boni"] = $"{currentFees}, {settings.Fees}";
                        }
                    }
                }
                else
                {
                    // Darf passieren, da Erwachsene und Senioren getrennt sind
                }
            }
            return;
        }
    }
}


    static void AddGroup(CommandArgs settings, DataTable tableMembers, DataSet byData)
{
    foreach (DataTable table in byData.Tables)
    {
        if (table.TableName == settings.ByTableName)
        {
            foreach (DataRow row in table.Rows)
            {
                var memberNumber = row.Field<string>("Mitgliedsnummer");
                var member = tableMembers.AsEnumerable().FirstOrDefault(r => r.Field<string>("Mitgliedsnummer *") == memberNumber);
                if (member != null)
                {
                    var memberGroups = member.Field<string>("Gruppe(n) *");
                    if (string.IsNullOrEmpty(memberGroups))
                    {
                        member["Gruppe(n) *"] = settings.MemberGroupName;
                    }
                    else
                    {
                        if (!memberGroups.Contains(settings.MemberGroupName))
                        {
                            member["Gruppe(n) *"] = $"{memberGroups}, {settings.MemberGroupName}";
                        }
                    }
                }
                else
                {
                    // Darf passieren, da Erwachsene und Senioren getrennt sind
                }
            }
            return;
        }
    }

    Console.WriteLine($"Missing table {settings.ByTableName}!");
}