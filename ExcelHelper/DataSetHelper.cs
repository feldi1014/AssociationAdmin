using ExcelDataReader;
using System.Data;
using System.Text;

namespace ExcelHelper;


public class DataSetHelper
{
    private FileInfo _filePath;

    public DataSetHelper(FileInfo filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        this._filePath = filePath;
    }

    public DataSet Data { get; private set; } = new DataSet();

    public Dictionary<string, List<MemberGroupCards>> MemberGroupCards { get; private set; } = new Dictionary<string, List<MemberGroupCards>>();

    public bool GetDataSet()
    {
        try
        {
            using (var stream = File.Open(_filePath.FullName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    //var result = reader.AsDataSet();

                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    // The result of each spreadsheet is in result.Tables
                    Data = result;
                    return true;
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }

        return false;
    }



    public void UpdateRowContent()
    {
        Console.WriteLine("Updating row content");
        foreach (DataTable table in Data.Tables)
        {
            if (table.TableName == "Bestellungen")
            {
                Console.WriteLine($"{table.TableName} {table.Rows.Count}");
                List<string> columns = new List<string>() { "Bestelldatum", "Vorname", "Nachname", "Strasse", "PLZ", "Ort", "Land", "Mitgliedsnummer", "Mitgliedsgruppen" };
                DataRow? lastRow = null;
                foreach (DataRow row in table.Rows)
                {
                    string? bestDateText = row["Bestelldatum"] as string;
                    if (!string.IsNullOrEmpty(bestDateText))
                    {
                        lastRow = row;
                    }
                    else
                    {
                        if (lastRow == null)
                        {
                            throw new InvalidOperationException("Missing last row!");
                        }

                        foreach (string columnName in columns)
                        {
                            row[columnName] = lastRow[columnName];
                        }
                    }
                }
            }
        }
    }

    public void GetCardsAndGroups()
    {
        Console.WriteLine("Mitgliedergruppen und Wunschkarten");
        var dict = new Dictionary<string, List<MemberGroupCards>>();

        foreach (DataTable table in Data.Tables)
        {
            if (table.TableName == "Bestellungen")
            {
                foreach (DataRow row in table.Rows)
                {
                    string? group = row["Mitgliedsgruppen"] as string;
                    string? karte = row["Wunschkarte"] as string;
                    List<MemberGroupCards> list;
                    if (!dict.ContainsKey(group))
                    {
                        list = new List<MemberGroupCards>();
                        dict[group] = list;
                    }
                    else
                    {
                        list = dict[group];
                    }

                    if (!string.IsNullOrEmpty(karte))
                    {
                        var memberGroupCard = new MemberGroupCards(group, karte);
                        if (!list.Contains(memberGroupCard))
                        {
                            list.Add(memberGroupCard);
                        }
                        else
                        {
                            int i = list.IndexOf(memberGroupCard);
                            list[i].Count++;
                        }
                    }
                }
            }
        }
        MemberGroupCards = dict;
        if (MemberGroupCards.Count == 0)
        {
            Console.WriteLine("Keine Karten verkauft!");
            return;
        }
        foreach(var pair in dict)
        {
            foreach(var m in pair.Value) 
            {
                Console.WriteLine(m.ToString());
            }
        }
    }

    public void OptimizeOrders()
    {
        DataTable bestellungen = Data.Tables["Bestellungen"];
        if (bestellungen == null)
        {
            return;
        }

        var gruppierteBestellungen =
            (from row in bestellungen.AsEnumerable()
                group row by new
                {
                    Bestelldatum = row.Field<string>("Bestelldatum"),
                    Mitgliedsnummer = row.Field<string>("Mitgliedsnummer"),
                    Vorname = row.Field<string>("Vorname"),
                    Nachname = row.Field<string>("Nachname")
                } into g
                select new MemberOrderPositions(g.Key.Bestelldatum, g.Key.Mitgliedsnummer)
                {
                    Vorname = g.Key.Vorname,
                    Nachname = g.Key.Nachname,
                    BestellPositionen = g.Count()
                }).Distinct().ToList();

        var anzahlBestellungen =
            (from row in gruppierteBestellungen
                group row by row.Mitgliedsnummer
            into g
                select new
                {
                    g.Key,
                    Bestellungen = g.Count()
                }).ToList();

        var doppeltBestellungen = anzahlBestellungen.Where(x => x.Bestellungen > 1).ToList();

        if (doppeltBestellungen.Any())
        {
            Console.WriteLine($"Mitglieder mit doppelten Bestellungen {doppeltBestellungen.Count()}:");
            foreach (var item in doppeltBestellungen)
            {
                var list = gruppierteBestellungen.Where(x => x.Mitgliedsnummer == item.Key);
                foreach (var item2 in list)
                {
                    Console.WriteLine(item2);
                }
            }
        }

    }

    public void ValidateOrders()
    {
        DataTable bestellungen = Data.Tables["Bestellungen"];
        if (bestellungen == null)
        {
            return;
        }

        var gefilterteBestellungen = bestellungen.AsEnumerable().Where(row =>
        (row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")) &&
            row.Field<string>("Wunschkarte").Contains("JUGEND")).ToList();
        // Gefilterte Ergebnisse anzeigen
        if (gefilterteBestellungen.Any())
        {
            Console.WriteLine("Falsche Kartenbestellung Erwachsene:");
            foreach (var row in gefilterteBestellungen)
            {
                Console.WriteLine($"{row["Mitgliedsnummer"]} {row["Nachname"]} {row["Mitgliedsgruppen"]}, Wunschkarte: {row["Wunschkarte"]}");
            }
        }
        gefilterteBestellungen = bestellungen.AsEnumerable().Where(row =>
            (row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")) &&
            !row.Field<string>("Wunschkarte").Contains("JUGEND")).ToList();
        // Gefilterte Ergebnisse anzeigen
        if (gefilterteBestellungen.Any())
        {
            Console.WriteLine("Falsche Kartenbestellung Jugend:");
            foreach (var row in gefilterteBestellungen)
            {
                Console.WriteLine($"{row["Mitgliedsnummer"]} {row["Nachname"]} {row["Mitgliedsgruppen"]}, Wunschkarte: {row["Wunschkarte"]}");
            }
        }
    }

    public void AppendColumnMitgliedsJahre(string tableName)
    {
        var today = DateTime.Today;
        foreach (DataTable table in Data.Tables)
        {
            if (table.TableName == "Mitglieder" && table.TableName == tableName)
            {
                table.Columns.Add("Mitgliedsjahre", typeof(int));
                foreach (DataRow row in table.Rows)
                {
                    if (string.IsNullOrEmpty(row["angelflix ID"].ToString()))
                    {
                        continue;
                    }
                    if (DateTime.TryParse(row["Beitrittsdatum"].ToString(), out DateTime dt))
                    {
                        row["Mitgliedsjahre"] = today.Year - dt.Year;
                    }
                }
            }

            if (table.TableName == "Bestellungen" && table.TableName == tableName)
            {
                table.Columns.Add("Mitgliedsjahre", typeof(int));
                if (Data.Tables.Contains("Mitglieder"))
                {
                    DataTable mitgliederTable = Data.Tables["Mitglieder"];

                    foreach (DataRow row in table.Rows)
                    {
                        string nummer = row.Field<string>("Mitgliedsnummer");
                        var mitglied = mitgliederTable.AsEnumerable().First(x => x.Field<string>("Mitgliedsnummer *") == nummer);
                        row["Mitgliedsjahre"] = mitglied["Mitgliedsjahre"];
                    }
                }
            }
        }
    }

    public void SetPrimaryKey(string tableName, string keyColumn)
    {
        var table = Data.Tables[tableName];
        var rows = table?.AsEnumerable().Where(x => string.IsNullOrEmpty(x.Field<string>(keyColumn))).ToList();
        if (rows?.Count > 0 )
        {
            for(int indexer = 0; indexer < rows.Count; indexer++)
            {
                var row = rows[indexer];
                table.Rows.Remove(row);
            }

            table.PrimaryKey = [table.Columns[keyColumn]];
        }
    }

    public void AppendColumnsForLastYear()
    {
        DataTable table = Data.Tables["LastYear"];
        if (Data.Tables.Contains("Bestellungen"))
        {
            DataTable bestellungen = Data.Tables["Bestellungen"];
            bestellungen.Columns.Add("Karte VJ", typeof(bool));
            bestellungen.Columns.Add("GW VJ", typeof(bool));
            bestellungen.Columns.Add("Ka VJ", typeof(bool));
            foreach (DataRow row in bestellungen.Rows)
            {
                row["Karte VJ"] = false;
                row["GW VJ"] = false;
                row["Ka VJ"] = false;
                string nummer = row.Field<string>("Mitgliedsnummer");
                var list = table.AsEnumerable().Where(x => x.Field<Double>("Nr").ToString() == nummer).ToList();
                if (list.Count > 0)
                {
                    foreach (DataRow last in list)
                    {
                        string karte = last["Gewässer"].ToString();
                        if (string.IsNullOrEmpty(karte))
                        {
                            continue;
                        }
                        row["Karte VJ"] = true;
                        if (karte.Contains("Günz"))
                        {
                            row["GW VJ"] = true;
                        }
                        if (karte.Contains("Kammel"))
                        {
                            row["Ka VJ"] = true;
                        }
                    }
                }
            }
        }

    }
}

