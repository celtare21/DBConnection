using System;
using System.Globalization;
// ReSharper disable UnusedMember.Global

namespace DBConnection.Helpers
{
    public static class MiscHelpers
    {
        public static string ConversionWrapper(object elem)
        {
            return elem switch
            {
                int => elem.ToString(),
                string => elem.ToString(),
                DateTime time => time.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
                TimeSpan span => span.ToString(@"hh\:mm"),
                var _ => null
            };
        }

        public static double GetIndice(TimeSpan time) =>
            (DateTime.Parse(time.ToString(@"hh\:mm")) - DateTime.Parse("00:00")).TotalHours;

        public static string TransformOverHour(TimeSpan span) =>
            $"{(int)span.TotalHours}:{span:mm}";

        public static string RemoveWhitespace(string str) =>
            string.Join("", str.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));

        public static void CheckForInternetConnection()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                {
                    using (client.OpenRead("http://google.com/generate_204"))
                    {
                        // Do Nothing.
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
