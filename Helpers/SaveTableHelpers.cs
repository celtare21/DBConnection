using System;
using System.Data.SqlClient;
using System.Windows.Forms;

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

                Console.WriteLine(tableName);

                if (tableName.Contains("office"))
                {
                    var officeTable = DatabaseHelpers.GetAllElementsLm(conn, tableName, current, true);
                    SaveHelpers.SaveTableOffice(officeTable, folder, tableName);
                }
                else
                {
                    var localTable = DatabaseHelpers.GetAllElementsLm(conn, tableName, current, false);
                    SaveHelpers.SaveTable(localTable, folder, tableName);
                }
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
