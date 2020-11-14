using System;
using System.Collections.Generic;
using System.IO;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

namespace DBConnection.Helpers
{
    public static class SaveHelpers
    {
        public static void SaveTable(Dictionary<string, List<object>> table, string path, string tableName)
        {
            var zero = TimeSpan.FromHours(0);
            var total = (time: zero, curs: zero, pregatire: zero, recuperare: zero);
            int max = table["id"].Count, j = 0;

            var loadedFile = new ExcelFile();
            ExcelWorksheet worksheet = loadedFile.Worksheets.Add("Tables");

            ExcelTableHelpers.PopulateTables(worksheet, max);

            foreach (var key in table.Keys)
            {
                TimeSpan totalTime = TimeSpan.FromHours(0);

                for (int i = 0; i < max; i++)
                {
                    switch (key)
                    {
                        case "id":
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper((int)table[key][i]);
                            break;
                        case "date":
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper((DateTime)table[key][i]);
                            break;
                        case "oraIncepere":
                        case "oraFinal":
                        case "cursAlocat":
                        case "pregatireAlocat":
                        case "recuperareAlocat":
                        case "total":
                            var localTime = (TimeSpan)table[key][i];
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper(localTime);
                            totalTime += localTime;
                            break;
                        case "observatii":
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper((string)table[key][i]);
                            break;
                    }
                }

                switch (j)
                {
                    case 4:
                        total.curs = totalTime;
                        break;
                    case 5:
                        total.pregatire = totalTime;
                        break;
                    case 6:
                        total.recuperare = totalTime;
                        break;
                    case 7:
                        total.time = totalTime;
                        break;
                }

                ++j;
            }

            var valoare = (curs: total.curs.TotalHours * Constant.PretCurs,
                pregatire: total.pregatire.TotalHours * Constant.PretPregatire,
                recuperare: total.recuperare.TotalHours * Constant.PretRecuperare);

            worksheet.Cells[max + Constant.Offset, 1].Value = MiscHelpers.TransformOverHour(total.time);
            worksheet.Cells[max + Constant.Offset + 1, 1].Value = MiscHelpers.TransformOverHour(total.curs);
            worksheet.Cells[max + Constant.Offset + 2, 1].Value = MiscHelpers.TransformOverHour(total.pregatire);
            worksheet.Cells[max + Constant.Offset + 3, 1].Value = MiscHelpers.TransformOverHour(total.recuperare);
            worksheet.Cells[max + Constant.Offset + 1, 2].Value = Constant.PretCurs;
            worksheet.Cells[max + Constant.Offset + 2, 2].Value = Constant.PretPregatire;
            worksheet.Cells[max + Constant.Offset + 3, 2].Value = Constant.PretRecuperare;
            worksheet.Cells[max + Constant.Offset + 1, 3].Value = total.curs.TotalHours;
            worksheet.Cells[max + Constant.Offset + 2, 3].Value = total.pregatire.TotalHours;
            worksheet.Cells[max + Constant.Offset + 3, 3].Value = total.recuperare.TotalHours;
            worksheet.Cells[max + Constant.Offset + 1, 4].Value = valoare.curs;
            worksheet.Cells[max + Constant.Offset + 2, 4].Value = valoare.pregatire;
            worksheet.Cells[max + Constant.Offset + 3, 4].Value = valoare.recuperare;
            worksheet.Cells[max + Constant.Offset + 4, 4].Value = valoare.curs + valoare.pregatire + valoare.recuperare;

            if (max > 0)
            {
                Table tableMain = worksheet.Tables.Add("Table", $"A1:I{max + 1}", true);
                tableMain.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
                Table tableLittle = worksheet.Tables.Add("TableLittle",
                    $"A{max + Constant.Offset + 1}:E{max + Constant.Offset + 4}", true);
                tableLittle.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
            }

            try
            {
                loadedFile.Save($@"{path}\{tableName}.xlsx");
            }
            catch (IOException)
            {
                Console.WriteLine("Please close the document before saving!");
            }
        }

        public static void SaveTableOffice(Dictionary<string, List<object>> table, string path, string tableName)
        {
            var zero = TimeSpan.FromHours(0);
            TimeSpan totalTimeSheet = zero;
            int max = table["id"].Count, j = 0;

            var loadedFile = new ExcelFile();
            ExcelWorksheet worksheet = loadedFile.Worksheets.Add("Tables");

            ExcelTableHelpers.PopulateTablesOffice(worksheet, max);

            foreach (var key in table.Keys)
            {
                for (int i = 0; i < max; i++)
                {
                    TimeSpan totalTime = TimeSpan.FromHours(0);

                    switch (key)
                    {
                        case "id":
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper((int)table[key][i]);
                            break;
                        case "date":
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper((DateTime)table[key][i]);
                            break;
                        case "oraIncepere":
                        case "oraFinal":
                        case "total":
                            var localTime = (TimeSpan)table[key][i];
                            worksheet.Cells[i + 1, j].Value = MiscHelpers.ConversionWrapper(localTime);
                            totalTime += localTime;
                            break;
                    }

                    if (j == 4)
                        totalTimeSheet += totalTime;
                }

                ++j;
            }

            var valoareTotal = totalTimeSheet.TotalHours * Constant.PretOffice;

            worksheet.Cells[max + Constant.Offset + 1, 1].Value = MiscHelpers.TransformOverHour(totalTimeSheet);
            worksheet.Cells[max + Constant.Offset + 1, 2].Value = Constant.PretOffice;
            worksheet.Cells[max + Constant.Offset + 1, 3].Value = totalTimeSheet.TotalHours;
            worksheet.Cells[max + Constant.Offset + 1, 4].Value = valoareTotal;

            if (max > 0)
            {
                Table tableMain = worksheet.Tables.Add("Table", $"A1:E{max + 1}", true);
                tableMain.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
                Table tableLittle = worksheet.Tables.Add("TableLittle",
                    $"A{max + Constant.Offset + 1}:E{max + Constant.Offset + 2}", true);
                tableLittle.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;
            }

            try
            {
                loadedFile.Save($@"{path}\{tableName}.xlsx");
            }
            catch (IOException)
            {
                Console.WriteLine("Please close the document before saving!");
            }
        }
    }
}
