using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

namespace FeedbackDownload
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            const string connStr = "//";

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            CheckForInternetConnection();

            using (var conn = new SqlConnection(connStr))
            {
                SaveAllTables(conn);
            }
        }

        private static int GetMaxNumber(SqlConnection conn)
        {
            int i = 0;
            string query;

            conn.Open();

            query = $@"SELECT id FROM ""feedback""";

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ++i;
                    }
                }
            }

            conn.Close();

            return i;
        }

        private static List<Entries> SortEntries(SqlConnection conn)
        {
            var entries = new List<Entries>();

            foreach (var element in GetAllElementsClass(conn))
            {
                entries.Add(element);
            }

            entries = entries.OrderBy(x => x.Name).ToList();

            return entries;
        }

        private static IEnumerable<Entries> GetAllElementsClass(SqlConnection conn)
        {
            string query;
            List<int> id;
            List<DateTime> date;
            List<string> name, entry1, entry2, entry3, entry4;
            int max = GetMaxNumber(conn);

            conn.Open();

            query = $@"SELECT id FROM ""feedback""";

            id = new List<int>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        id.Add((int)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT date FROM ""feedback""";

            date = new List<DateTime>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        date.Add((DateTime)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT name FROM ""feedback""";

            name = new List<string>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        name.Add((string)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT entry1 FROM ""feedback""";

            entry1 = new List<string>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entry1.Add((string)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT entry2 FROM ""feedback""";

            entry2 = new List<string>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entry2.Add((string)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT entry3 FROM ""feedback""";

            entry3 = new List<string>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entry3.Add((string)reader.GetValue(0));
                    }
                }
            }

            query = $@"SELECT entry4 FROM ""feedback""";

            entry4 = new List<string>();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        entry4.Add((string)reader.GetValue(0));
                    }
                }
            }

            for (int i = 0; i < max; i++)
            {
                yield return new Entries(id[i], date[i], name[i], entry1[i], entry2[i], entry3[i], entry4[i]);
            }

            conn.Close();
        }

        private static void DeleteAll(SqlConnection conn)
        {
            string query = $@"DELETE FROM ""feedback""";

            using (var command = new SqlCommand(query, conn))
            {
                ExecuteCommandDB(command, conn);
            }
        }

        private static void CreateTableDBFeedback(SqlConnection conn)
        {
            string query = $@"CREATE TABLE ""feedback"" (" +
                            "id       INT         NOT NULL    IDENTITY PRIMARY KEY," +
                            "date     DATE        NOT NULL," +
                            "name     ntext       NOT NULL," +
                            "entry1   ntext       NOT NULL," +
                            "entry2   ntext       NOT NULL," +
                            "entry3   ntext       NOT NULL," +
                            "entry4   ntext       NOT NULL," +
                            ");";

            using (var command = new SqlCommand(query, conn))
            {
                ExecuteCommandDB(command, conn);
            }
        }

        private static void SaveAllTables(SqlConnection conn, string user_name = "")
        {
            string folder = OpenFolder();
            int max = GetMaxNumber(conn);
            var local_list = SortEntries(conn);

            SaveTable(local_list, max, folder);
        }

        private static void ExecuteCommandDB(SqlCommand command, SqlConnection conn)
        {
            conn.Open();

            _ = command.ExecuteNonQuery();

            conn.Close();
        }

        private static void SaveTable(List<Entries> list, int max, string path)
        {
            ExcelFile loadedFile;
            ExcelWorksheet worksheet;
            Table tableMain;
            int i = 0, j;

            loadedFile = new ExcelFile();
            worksheet = loadedFile.Worksheets.Add("Tables");

            PopulateTables(worksheet);

            foreach (var entry in list)
            {
                j = 0;

                worksheet.Cells[i + 1, j++].Value = entry.Id;
                worksheet.Cells[i + 1, j++].Value = entry.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                worksheet.Cells[i + 1, j++].Value = entry.Name;
                worksheet.Cells[i + 1, j++].Value = entry.Entry1;
                worksheet.Cells[i + 1, j++].Value = entry.Entry2;
                worksheet.Cells[i + 1, j++].Value = entry.Entry3;
                worksheet.Cells[i + 1, j++].Value = entry.Entry4;

                ++i;
            }

            if (max > 0)
            {
                tableMain = worksheet.Tables.Add("Table", $"A1:G{max + 1}", true);
                tableMain.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
            }

            try
            {
                loadedFile.Save($@"{path}\feedback.xlsx");
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

        private static void PopulateTables(ExcelWorksheet worksheet)
        {
            worksheet.Columns[0].SetWidth(135, LengthUnit.Pixel);
            worksheet.Columns[1].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[2].SetWidth(120, LengthUnit.Pixel);
            worksheet.Columns[3].SetWidth(110, LengthUnit.Pixel);
            worksheet.Columns[4].SetWidth(110, LengthUnit.Pixel);
            worksheet.Columns[5].SetWidth(110, LengthUnit.Pixel);
            worksheet.Columns[6].SetWidth(110, LengthUnit.Pixel);

            worksheet.Cells[0, 0].Value = "ID";
            worksheet.Cells[0, 1].Value = "Data";
            worksheet.Cells[0, 2].Value = "Name";
            worksheet.Cells[0, 3].Value = "Entry1";
            worksheet.Cells[0, 4].Value = "Entry2";
            worksheet.Cells[0, 5].Value = "Entry3";
            worksheet.Cells[0, 6].Value = "Entry4";
        }

        private static void CheckForInternetConnection()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                {
                    using (client.OpenRead("http://google.com/generate_204"))
                    {
                        return;
                    }
                }
            }
            catch
            {
                Console.WriteLine("No internet connection!");
                Console.ReadLine();
                Environment.Exit(0);
            }
        }
    }

    readonly struct Entries
    {
        public int Id { get; }
        public DateTime Date { get; }
        public string Name { get; }
        public string Entry1 { get; }
        public string Entry2 { get; }
        public string Entry3 { get; }
        public string Entry4 { get; }

        public Entries(int id, DateTime date, string name, string entry1, string entry2, string entry3, string entry4) =>
                (Id, Date, Name, Entry1, Entry2, Entry3, Entry4) = (id, date, name, entry1, entry2, entry3, entry4);
    }
}
