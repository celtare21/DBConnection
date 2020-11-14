using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace DBConnection.Helpers
{
    public static class DatabaseHelpers
    {
        public static string HasUser(SqlConnection conn, string pin)
        {
            string query = $@"SELECT username FROM users WHERE password = '{pin}'";
            string result = null;

            conn.Open();

            using (var command = new SqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {
                    reader.Read();

                    if (reader.HasRows)
                        result = reader.GetValue(0).ToString();
                }
            }

            conn.Close();

            return result;
        }

        public static IEnumerable<string> GetAllTables(SqlConnection conn)
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

                        if (value.Contains("prezenta.") && !value.Contains("test"))
                            tables.Add(value);
                    }
                }
            }

            conn.Close();

            return tables;
        }

        public static Dictionary<string, List<object>> GetAllElementsLm(SqlConnection conn, string username, bool current, bool office)
        {
            string[] columns =
            {
                "id",
                "date",
                "oraIncepere",
                "oraFinal",
                "cursAlocat",
                "pregatireAlocat",
                "recuperareAlocat",
                "total",
                "observatii"
            };
            string[] officeColumns =
            {
                "id",
                "date",
                "oraIncepere",
                "oraFinal",
                "total"
            };

            var dic = new Dictionary<string, List<object>>();

            conn.Open();

            if (office)
            {
                foreach (var elem in officeColumns)
                {
                    GetAllElementsHelper(conn, dic, elem, username, current);
                }
            }
            else
            {
                foreach (var elem in columns)
                {
                    GetAllElementsHelper(conn, dic, elem, username, current);
                }
            }

            conn.Close();

            return dic;
        }

        private static void GetAllElementsHelper(SqlConnection conn, IDictionary<string, List<object>> dic, string elem, string table, bool current)
        {
            dic.Add(elem, new List<object>());

            var query = current
                ? $@"SELECT {elem} FROM ""{table}"""
                : $@"SELECT {elem} FROM ""{table}"" WHERE date >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) AND date < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)";

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

        public static void DeleteLastMonthE(SqlConnection conn)
        {
            foreach (var tableName in GetAllTables(conn))
            {
                string query = $@"DELETE FROM ""{tableName}"" WHERE date >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) AND date < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)";

                using (var command = new SqlCommand(query, conn))
                {
                    ExecuteCommandDb(command, conn);
                }
            }
        }

        private static void ExecuteCommandDb(IDbCommand command, IDbConnection conn)
        {
            conn.Open();

            _ = command.ExecuteNonQuery();

            conn.Close();
        }

        public static void CreateTableDb(SqlConnection conn, string name)
        {
            string query = $@"CREATE TABLE ""prezenta.{name}"" (" +
                            "id                  INT     NOT NULL    IDENTITY PRIMARY KEY," +
                            "date                DATE    NOT NULL," +
                            "oraIncepere        TIME    NOT NULL," +
                            "oraFinal           TIME    NOT NULL," +
                            "cursAlocat         TIME    NOT NULL," +
                            "pregatireAlocat    TIME    NOT NULL," +
                            "recuperareAlocat   TIME    NOT NULL," +
                            "total               TIME    NOT NULL," +
                            "observatii          ntext   NOT NULL," +
                            ");";

            using (var command = new SqlCommand(query, conn))
            {
                ExecuteCommandDb(command, conn);
            }
        }

        public static void CreateTableOfficeDb(SqlConnection conn, string name)
        {
            string query = $@"CREATE TABLE ""prezenta.office.{name}"" (" +
                           "id                  INT     NOT NULL    IDENTITY PRIMARY KEY," +
                           "date                DATE    NOT NULL," +
                           "oraIncepere        TIME    NOT NULL," +
                           "oraFinal           TIME    NOT NULL," +
                           "total               TIME    NOT NULL," +
                           ");";

            using (var command = new SqlCommand(query, conn))
            {
                ExecuteCommandDb(command, conn);
            }
        }
    }
}
