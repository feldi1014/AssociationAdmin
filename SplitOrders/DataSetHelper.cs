using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace SplitOrders
{
    internal class DataSetHelper
    {
        private FileInfo _filePath;

        public DataSetHelper(FileInfo filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            this._filePath = filePath;
        }

        public DataSet Data { get; private set; }
        public Dictionary<string, List<MemberGroupCards>> MemberGroupCards { get; private set; }

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
                    DataRow lastRow = null;
                    foreach (DataRow row in table.Rows)
                    {
                        string bestDateText = row["Bestelldatum"].ToString();
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
                        string group = row["Mitgliedsgruppen"].ToString();
                        string karte = row["Wunschkarte"].ToString();
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
            foreach(var pair in dict)
            {
                foreach(var m in pair.Value) 
                {
                    Console.WriteLine(m.ToString());
                }
            }
        }

    }
}
