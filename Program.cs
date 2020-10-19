using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

namespace DBConnection
{
    class Program
    {
        private static ExcelWorksheet Worksheet;
        private static ExcelFile LoadedFile;

        [STAThread]
        static void Main()
        {
            const string connStr = "//";

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            using (var conn = new SqlConnection(connStr))
            {
                SaveAllTables(conn);
            }

            Console.ReadLine();
        }

        private static List<string> GetAllTables(SqlConnection conn)
        {
            const string query = "SELECT name FROM sys.Tables";
            var tables = new List<string>();

            conn.Open();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string value = reader.GetValue(0).ToString();

                        tables.Add(value);
                    }
                }
            }

            conn.Close();

            return tables;
        }

        private static Dictionary<string, List<object>> GetAllElements(SqlConnection conn, string table)
        {
            string query;
            string[] columns = { "id", "date", "ora_incepere", "ora_final", "curs_alocat", "pregatire_alocat", "recuperare_alocat", "total" };
            var dic = new Dictionary<string, List<object>>();

            conn.Open();

            foreach (var elem in columns)
            {
                dic.Add(elem, new List<object>());

                query = $@"SELECT {elem} FROM ""{table}""";

                using (var command = new SqlCommand(query, conn))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            dic[elem].Add(reader.GetValue(0));
                        }
                    }
                }
            }

            conn.Close();

            return dic;
        }

        private static void ExecuteCommandDB(SqlCommand command, SqlConnection conn)
        {
            conn.Open();

            _ = command.ExecuteNonQuery();

            conn.Close();
        }

        private static void CreateTableDB(SqlConnection conn, string name)
        {
            string query = $@"CREATE TABLE ""prezenta.{name}"" (" +
                            "id                  INT     NOT NULL    IDENTITY PRIMARY KEY," +
                            "date                DATE    NOT NULL," +
                            "ora_incepere        TIME    NOT NULL," +
                            "ora_final           TIME    NOT NULL," +
                            "curs_alocat         TIME    NOT NULL," +
                            "pregatire_alocat    TIME    NOT NULL," +
                            "recuperare_alocat   TIME    NOT NULL," +
                            "total               TIME    NOT NULL," +
                            ");";

            using (var command = new SqlCommand(query, conn))
            {
                ExecuteCommandDB(command, conn);
            }
        }

        private static void SaveAllTables(SqlConnection conn)
        {
            string folder = OpenFolder();

            foreach (var table_name in GetAllTables(conn))
            {
                var local_table = GetAllElements(conn, table_name);

                SaveTable(local_table, folder, table_name);
            }
        }

        private static void SaveTable(Dictionary<string, List<object>> table, string path, string tableName)
        {
            Table tableMain, tableLittle;
            var zero = TimeSpan.FromHours(0);
            var total = (time: zero, curs: zero, pregatire: zero, recuperare: zero);
            int max = table["id"].Count(), j = 0;

            LoadedFile = new ExcelFile();
            Worksheet = LoadedFile.Worksheets.Add("Tables");

            PopulateTables(max);

            foreach (var key in table.Keys)
            {
                TimeSpan _total = TimeSpan.FromHours(0);

                for (int i = 0; i < max; i++)
                {
                    switch (key)
                    {
                        case "id":
                            Worksheet.Cells[i + 1, j].Value = ConversionWrapper<int>((int)table[key][i]);
                            break;
                        case "date":
                            Worksheet.Cells[i + 1, j].Value = ConversionWrapper<string>((DateTime)table[key][i]);
                            break;
                        case "ora_incepere":
                        case "ora_final":
                        case "curs_alocat":
                        case "pregatire_alocat":
                        case "recuperare_alocat":
                        case "total":
                            var localTime = (TimeSpan)table[key][i];
                            Worksheet.Cells[i + 1, j].Value = ConversionWrapper<string>(localTime);
                            _total += localTime;
                            break;
                        default:
                            break;
                    }
                }

                switch (j)
                {
                    case 4:
                        total.curs = _total;
                        break;
                    case 5:
                        total.pregatire = _total;
                        break;
                    case 6:
                        total.recuperare = _total;
                        break;
                    case 7:
                        total.time = _total;
                        break;
                    default:
                        break;
                }

                ++j;
            }

            var valoare = (curs: GetIndice(total.curs) * Constant.pret_curs,
                    pregatire: GetIndice(total.pregatire) * Constant.pret_pregatire,
                    recuperare: GetIndice(total.recuperare) * Constant.pret_recuperare);

            Worksheet.Cells[max + Constant.offset, 1].Value = TransformOverHour(total.time);
            Worksheet.Cells[max + Constant.offset + 1, 1].Value = TransformOverHour(total.curs);
            Worksheet.Cells[max + Constant.offset + 2, 1].Value = TransformOverHour(total.pregatire);
            Worksheet.Cells[max + Constant.offset + 3, 1].Value = TransformOverHour(total.recuperare);
            Worksheet.Cells[max + Constant.offset + 1, 2].Value = Constant.pret_curs;
            Worksheet.Cells[max + Constant.offset + 2, 2].Value = Constant.pret_pregatire;
            Worksheet.Cells[max + Constant.offset + 3, 2].Value = Constant.pret_recuperare;
            Worksheet.Cells[max + Constant.offset + 1, 3].Value = GetIndice(total.curs);
            Worksheet.Cells[max + Constant.offset + 2, 3].Value = GetIndice(total.pregatire);
            Worksheet.Cells[max + Constant.offset + 3, 3].Value = GetIndice(total.recuperare);
            Worksheet.Cells[max + Constant.offset + 1, 4].Value = valoare.curs;
            Worksheet.Cells[max + Constant.offset + 2, 4].Value = valoare.pregatire;
            Worksheet.Cells[max + Constant.offset + 3, 4].Value = valoare.recuperare;
            Worksheet.Cells[max + Constant.offset + 4, 4].Value = valoare.curs + valoare.pregatire + valoare.recuperare;

            if (max > 0)
            {
                tableMain = Worksheet.Tables.Add("Table", $"A1:H{max + 1}", true);
                tableMain.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
                tableLittle = Worksheet.Tables.Add("TableLittle", $"A{max + Constant.offset + 1}:E{max + Constant.offset + 4}", true);
                tableLittle.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
            }

            try
            {
                LoadedFile.Save($@"{path}\{tableName}.xlsx");
            }
            catch (System.IO.IOException)
            {
                Console.WriteLine("Please close the document before saving!");
            }
        }

        private static string OpenFolder()
        {
            using (var folderLocation = new FolderBrowserDialog())
            {
                var result = folderLocation.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrEmpty(folderLocation.SelectedPath))
                {
                    return folderLocation.SelectedPath;
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            return null;
        }

        private static void PopulateTables(int max)
        {
            Worksheet.Columns[0].SetWidth(135, LengthUnit.Pixel);
            Worksheet.Columns[1].SetWidth(100, LengthUnit.Pixel);
            Worksheet.Columns[2].SetWidth(100, LengthUnit.Pixel);
            Worksheet.Columns[3].SetWidth(90, LengthUnit.Pixel);
            Worksheet.Columns[4].SetWidth(90, LengthUnit.Pixel);
            Worksheet.Columns[5].SetWidth(110, LengthUnit.Pixel);
            Worksheet.Columns[6].SetWidth(120, LengthUnit.Pixel);
            Worksheet.Columns[7].SetWidth(90, LengthUnit.Pixel);

            Worksheet.Cells[0, 0].Value = "ID";
            Worksheet.Cells[0, 1].Value = "Data";
            Worksheet.Cells[0, 2].Value = "Ora incepere";
            Worksheet.Cells[0, 3].Value = "Ora sfarsit";
            Worksheet.Cells[0, 4].Value = "Curs alocat";
            Worksheet.Cells[0, 5].Value = "Pregatire alocat";
            Worksheet.Cells[0, 6].Value = "Recuperare alocat";
            Worksheet.Cells[0, 7].Value = "Ora total";

            Worksheet.Cells[max + Constant.offset, 0].Value = "TOTAL:";
            Worksheet.Cells[max + Constant.offset + 1, 0].Value = "TOTAL CURS:";
            Worksheet.Cells[max + Constant.offset + 2, 0].Value = "TOTAL PREGATIRE:";
            Worksheet.Cells[max + Constant.offset + 3, 0].Value = "TOTAL RECUPERARE:";
            Worksheet.Cells[max + Constant.offset, 2].Value = "PRET/H";
            Worksheet.Cells[max + Constant.offset, 3].Value = "INDICE";
            Worksheet.Cells[max + Constant.offset, 4].Value = "VALOARE";

            Worksheet.Cells[max + Constant.offset + 4, 3].Value = "TOTAL ORE:";
        }

        private static T ConversionWrapper<T>(object elem)
        {
            switch (elem)
            {
                case int:
                    return (T)elem;
                case DateTime:
                    return (T)(object)((DateTime)elem).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                case TimeSpan:
                    return (T)(object)((TimeSpan)elem).ToString(@"hh\:mm");
                default:
                    throw new ArgumentException();
            }
        }

        private static double GetIndice(TimeSpan time) =>
                (DateTime.Parse(time.ToString(@"hh\:mm")) - DateTime.Parse("00:00")).TotalHours;

        private static string TransformOverHour(TimeSpan span) =>
                $"{(int)span.TotalHours}:{span:mm}";
    }

    public class TableColumns
    {
        public string Date { get; }
        public TimeSpan OraIncepere { get; }
        public TimeSpan OraFinal { get; }
        public TimeSpan CursAlocat { get; }
        public TimeSpan PregatireAlocat { get; }
        public TimeSpan RecuperareAlocat { get; }
        public TimeSpan Total { get; }

        public TableColumns(string date, TimeSpan oraIncepere, TimeSpan oraFinal, TimeSpan cursAlocat, TimeSpan pregatireAlocat, TimeSpan recuperareAlocat, TimeSpan total)
        {
            Date = date;
            OraIncepere = oraIncepere;
            OraFinal = oraFinal;
            CursAlocat = cursAlocat;
            PregatireAlocat = pregatireAlocat;
            RecuperareAlocat = recuperareAlocat;
            Total = total;
        }
    }

    public class ExecuteFailure : Exception
    {
        public ExecuteFailure(string message)
        {
            Console.WriteLine(message);
        }
    }

    public class HoursOutOfBounds : Exception
    {
        public HoursOutOfBounds()
        {
            Console.WriteLine("Hours out of bounds!");
        }
    }

    public class AllParametersFalse : Exception
    {
        public AllParametersFalse()
        {
            Console.WriteLine("All parameters are false!");
        }
    }

    public static class Constant
    {
        public static int offset = 5;
        public static int pret_curs = 17, pret_pregatire = 8, pret_recuperare = 17;
    }
}
