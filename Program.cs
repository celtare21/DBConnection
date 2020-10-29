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
            const string connStr = "";

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            CheckForInternetConnection();

            using (var conn = new SqlConnection(connStr))
            {
                SaveAllTables(conn);
            }
        }

        private static Dictionary<string, List<object>> GetAllElementsLM(SqlConnection conn)
        {
            string query;
            string[] columns = { "id", "date", "name", "entry1", "entry2", "entry3", "entry4" };
            var dic = new Dictionary<string, List<object>>();

            conn.Open();

            foreach (var elem in columns)
            {
                dic.Add(elem, new List<object>());

                query = $@"SELECT {elem} FROM ""feedback""";

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
            var local_table = GetAllElementsLM(conn);

            SaveTable(local_table, folder);
        }

        private static void ExecuteCommandDB(SqlCommand command, SqlConnection conn)
        {
            conn.Open();

            _ = command.ExecuteNonQuery();

            conn.Close();
        }

        private static void SaveTable(Dictionary<string, List<object>> table, string path)
        {
            ExcelFile loadedFile;
            ExcelWorksheet worksheet;
            Table tableMain;
            int max = table["id"].Count(), j = 0;

            loadedFile = new ExcelFile();
            worksheet = loadedFile.Worksheets.Add("Tables");

            PopulateTables(worksheet);

            foreach (var key in table.Keys)
            {
                for (int i = 0; i < max; i++)
                {
                    switch (key)
                    {
                        case "id":
                            worksheet.Cells[i + 1, j].Value = ConversionWrapper<int>((int)table[key][i]);
                            break;
                        case "date":
                            worksheet.Cells[i + 1, j].Value = ConversionWrapper<string>((DateTime)table[key][i]);
                            break;
                        case "name":
                        case "entry1":
                        case "entry2":
                        case "entry3":
                        case "entry4":
                            worksheet.Cells[i + 1, j].Value = ConversionWrapper<string>((string)table[key][i]);
                            break;
                        default:
                            break;
                    }
                }

                ++j;
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

        private static T ConversionWrapper<T>(object elem)
        {
            switch (elem)
            {
                case int:
                case string:
                    return (T)elem;
                case DateTime:
                    return (T)(object)((DateTime)elem).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                default:
                    throw new ArgumentException();
            }
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
}
