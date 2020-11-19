using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using static DBConnection.Helpers.SaveHelpers;

namespace DBConnection.Helpers
{
    public static class SaveTableHelpers
    {
        public static void SaveAllTables(SqlConnection conn, string userName, bool current)
        {
            string folder = OpenFolder();

            if (folder == null)
                return;

            foreach (var tableName in DatabaseHelpers.GetAllTables(conn))
            {
                if (!string.IsNullOrEmpty(userName) && !tableName.Contains(userName.ToLowerInvariant()))
                    continue;

                var office = tableName.Contains("office");
                var localTable = DatabaseHelpers.GetAllElementsLm(conn, tableName, current, office);

                if (localTable == null)
                {
                    Console.WriteLine($"{tableName} is empty! Skipping...");
                    continue;
                }

                Console.WriteLine(tableName);

                if (office)
                    SaveTableOffice(localTable, folder, tableName);
                else
                    SaveTable(localTable, folder, tableName);
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
            }

            return null;
        }
    }
}
