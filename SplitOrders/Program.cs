using System.Data;
using ClosedXML.Excel;
using SplitOrders;

if (args.Length == 0)
{
    Console.WriteLine("Missing excel file!");
    return;
}

var filePath = new FileInfo(args[0]);
if (!filePath.Exists)
{
    Console.WriteLine($"Missing excel file at {filePath.FullName}!");
    return;
}

var helper = new DataSetHelper(filePath);


if (!helper.GetDataSet())
{
    return;
}

helper.UpdateRowContent();

helper.GetCardsAndGroups();

DataTable bestellungen = helper.Data.Tables["Bestellungen"];

var gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") && row.Field<string>("Wunschkarte").Contains("JUGEND")).ToList(); 
// Gefilterte Ergebnisse anzeigen
foreach (var row in gefilterteBestellungen) 
{
    Console.WriteLine($"{row["Mitgliedsnummer"]} {row["Nachname"]} {row["Mitgliedsgruppen"]}, Wunschkarte: {row["Wunschkarte"]}"); 
}


using (var workbook = new XLWorkbook())
{

    var builder = new WorksheetBuilder();

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) && row.Field<string>("Wunschkarte").Contains("Moos")).ToList();
    builder.Fill(workbook.Worksheets.Add("Erwachsene Moos"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) && row.Field<string>("Wunschkarte").Contains("Günzried")).ToList();
    builder.Fill(workbook.Worksheets.Add("Erwachsene GW"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) && row.Field<string>("Wunschkarte").Contains("Weißried")).ToList();
    builder.Fill(workbook.Worksheets.Add("Erwachsene WR"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) && row.Field<string>("Wunschkarte").Contains("Donau")).ToList();
    builder.Fill(workbook.Worksheets.Add("Erwachsene Donau"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) && row.Field<string>("Wunschkarte").Contains("Kammel")).ToList();
    builder.Fill(workbook.Worksheets.Add("Erwachsene Kammel"), gefilterteBestellungen, bestellungen);

    // Jugend
    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) && row.Field<string>("Wunschkarte").Contains("Moos")).ToList();
    builder.Fill(workbook.Worksheets.Add("Jugend Moos"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) && row.Field<string>("Wunschkarte").Contains("Günzried")).ToList();
    builder.Fill(workbook.Worksheets.Add("Jugend GW"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) && row.Field<string>("Wunschkarte").Contains("Weißried")).ToList();
    builder.Fill(workbook.Worksheets.Add("Jugend WR"), gefilterteBestellungen, bestellungen);

    gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) && row.Field<string>("Wunschkarte").Contains("Donau")).ToList();
    builder.Fill(workbook.Worksheets.Add("Jugend Donau"), gefilterteBestellungen, bestellungen);

    //gefilterteBestellungen = bestellungen.AsEnumerable().Where(row => (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) && row.Field<string>("Wunschkarte").Contains("Kammel")).ToList();
    //builder.Fill(workbook.Worksheets.Add("Jugend Kammel"), gefilterteBestellungen, bestellungen);

    workbook.SaveAs("SplitBestellungen.xlsx");
}



