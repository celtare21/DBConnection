using System;
using System.Data.SqlClient;
using DBConnection.Helpers;
using GemBox.Spreadsheet;
using static DBConnection.Helpers.DatabaseHelpers;
using static DBConnection.Helpers.SaveTableHelpers;

// ReSharper disable UnusedMember.Local
namespace DBConnection
{
    internal class Program
    {
        [STAThread]
        private static void Main()
        {
            const string connStr = "//";

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            MiscHelpers.CheckForInternetConnection();

            using (var conn = new SqlConnection(connStr))
            {
                PersonalDownloader(conn);
            }
        }

        private static void PersonalDownloader(SqlConnection conn)
        {
            string result;

            do
            {
                Console.WriteLine("Enter your pin:");

                var pin = ReadPassword();

                if (pin.Equals("9999"))
                    MasterDownloader(conn);

                result = HasUser(conn, pin);

                if (result == null)
                    Console.WriteLine("No users found!");
            } while (result == null);

            SaveAllTables(conn, result, true);
        }

        private static void MasterDownloader(SqlConnection conn)
        {
            SaveAllTables(conn, null, false);
            DeleteLastMonthE(conn);
            Environment.Exit(0);
        }

        private static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);

            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace && !string.IsNullOrEmpty(password))
                {
                    int pos = Console.CursorLeft;

                    password = password.Substring(0, password.Length - 1);

                    Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    Console.Write(" ");
                    Console.SetCursorPosition(pos - 1, Console.CursorTop);
                }
                info = Console.ReadKey(true);
            }

            Console.WriteLine();
            return password;
        }
    }
}
