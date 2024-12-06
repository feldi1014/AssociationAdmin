using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitOrders
{
    internal class WorksheetBuilder
    {
        private readonly DataTable _bestellungen;
        private readonly List<DataRow> _bestellungenErwachsene;
        private readonly List<DataRow> _bestellungenJugend;
        private readonly XLWorkbook _workbook;
        private readonly DataTable? _mitglieder;

        public WorksheetBuilder(DataSet data, XLWorkbook workbook)
        {
            if (data.Tables.Contains("Bestellungen"))
            {
                _bestellungen = data.Tables["Bestellungen"];
                _bestellungenErwachsene = _bestellungen.AsEnumerable().Where(row =>
                  row.Field<string>("Mitgliedsgruppen").Contains("Erwachsene") || row.Field<string>("Mitgliedsgruppen").Contains("Senioren")).ToList();
                _bestellungenJugend = _bestellungen.AsEnumerable().Where(row =>
                  row.Field<string>("Mitgliedsgruppen").Contains("Jugend") || row.Field<string>("Mitgliedsgruppen").Contains("Kind")).ToList();
            }

            if (data.Tables.Contains("Mitglieder"))
            {
                _mitglieder = data.Tables["Mitglieder"];
            }
            _workbook = workbook;
        }

        public void ExportAll()
        {
            if (_bestellungen != null)
            {
                ExportOverview();

                ExportErwachseneMoos();
                ExportErwachseneGw();
                ExportErwachseneWr();
                ExportErwachseneDonau();
                ExportErwachseneKammel();

                ExportJugendMoos();
                ExportJugendGw();
                ExportJugendWr();
                ExportJugendDonau();

                ExportOverviewDonau();
            }

            if (_mitglieder != null)
            {
                ExportMitglieder(); 
            }
        }

        internal void ExportMitglieder()
        {
            Fill(_workbook.Worksheets.Add("Mitglieder"), _mitglieder.AsEnumerable().ToList(), _mitglieder);
        }

        internal void ExportErwachseneDonau()
        {
            var gefilterteBestellungen = _bestellungenErwachsene.Where(row => row.Field<string>("Wunschkarte").Contains("Donau")).ToList();
            Fill(_workbook.Worksheets.Add("Erwachsene Donau"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportErwachseneGw()
        {
            var gefilterteBestellungen = _bestellungenErwachsene.Where(row => row.Field<string>("Wunschkarte").Contains("Günzried")).ToList();
            Fill(_workbook.Worksheets.Add("Erwachsene GW"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportErwachseneKammel()
        {
            var gefilterteBestellungen = _bestellungenErwachsene.Where(row => row.Field<string>("Wunschkarte").Contains("Kammel")).ToList();
            Fill(_workbook.Worksheets.Add("Erwachsene Kammel"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportErwachseneMoos()
        {
            var gefilterteBestellungen = _bestellungenErwachsene.Where(row => row.Field<string>("Wunschkarte").Contains("Moos")).ToList();
            Fill(_workbook.Worksheets.Add("Erwachsene Moos"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportErwachseneWr()
        {
            var gefilterteBestellungen = _bestellungenErwachsene.Where(row => row.Field<string>("Wunschkarte").Contains("Weißried")).ToList();
            Fill(_workbook.Worksheets.Add("Erwachsene WR"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportJugendDonau()
        {
            var gefilterteBestellungen = _bestellungenJugend.Where(row => row.Field<string>("Wunschkarte").Contains("Donau")).ToList();
            Fill(_workbook.Worksheets.Add("Jugend Donau"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportJugendGw()
        {
            var gefilterteBestellungen = _bestellungenJugend.Where(row => row.Field<string>("Wunschkarte").Contains("Günzried")).ToList();
            Fill(_workbook.Worksheets.Add("Jugend GW"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportJugendMoos()
        {
            var gefilterteBestellungen = _bestellungenJugend.Where(row => row.Field<string>("Wunschkarte").Contains("Moos")).ToList();
            Fill(_workbook.Worksheets.Add("Jugend Moos"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportJugendWr()
        {
            var gefilterteBestellungen = _bestellungenJugend.Where(row => row.Field<string>("Wunschkarte").Contains("Weißried")).ToList();
            Fill(_workbook.Worksheets.Add("Jugend WR"), gefilterteBestellungen, _bestellungen);
        }

        internal void ExportOverview()
        {
            Fill(_workbook.Worksheets.Add("Gesamt"), _bestellungen.AsEnumerable().ToList(), _bestellungen);
        }

        internal void ExportOverviewDonau()
        {
            var gefilterteBestellungen = _bestellungen.AsEnumerable().Where(row =>
              row.Field<string>("Wunschkarte").Contains("Donau")).ToList();
            Fill(_workbook.Worksheets.Add("Donau Gesamt"), gefilterteBestellungen, _bestellungen);
        }

        internal static void Fill(IXLWorksheet worksheet, List<DataRow> rows, DataTable table)
        {
            Console.WriteLine($"{worksheet.Name} Count: {rows.Count}");
            // Kopfzeile schreiben
            for (int i = 0; i < table.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;
            }

            // Tabelle füllen
            for (int rowIndex = 0; rowIndex < rows.Count(); rowIndex++)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    worksheet.Cell(rowIndex + 2, i + 1).Value = rows.ElementAt(rowIndex)[i].ToString();
                }
            }
        }
    }
}
