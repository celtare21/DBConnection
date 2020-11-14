using GemBox.Spreadsheet;

namespace DBConnection.Helpers
{
    public static class ExcelTableHelpers
    {
        public static void PopulateTables(ExcelWorksheet worksheet, int max)
        {
            worksheet.Columns[0].SetWidth(135, LengthUnit.Pixel);
            worksheet.Columns[1].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[2].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[3].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[4].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[5].SetWidth(110, LengthUnit.Pixel);
            worksheet.Columns[6].SetWidth(120, LengthUnit.Pixel);
            worksheet.Columns[7].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[8].SetWidth(140, LengthUnit.Pixel);

            worksheet.Cells[0, 0].Value = "ID";
            worksheet.Cells[0, 1].Value = "Data";
            worksheet.Cells[0, 2].Value = "Ora incepere";
            worksheet.Cells[0, 3].Value = "Ora sfarsit";
            worksheet.Cells[0, 4].Value = "Curs alocat";
            worksheet.Cells[0, 5].Value = "Pregatire alocat";
            worksheet.Cells[0, 6].Value = "Recuperare alocat";
            worksheet.Cells[0, 7].Value = "Ora total";
            worksheet.Cells[0, 8].Value = "Observatii";

            worksheet.Cells[max + Constant.Offset, 0].Value = "TOTAL:";
            worksheet.Cells[max + Constant.Offset + 1, 0].Value = "TOTAL CURS:";
            worksheet.Cells[max + Constant.Offset + 2, 0].Value = "TOTAL PREGATIRE:";
            worksheet.Cells[max + Constant.Offset + 3, 0].Value = "TOTAL RECUPERARE:";
            worksheet.Cells[max + Constant.Offset, 1].Value = "TOTAL ORE:";
            worksheet.Cells[max + Constant.Offset, 2].Value = "PRET/H";
            worksheet.Cells[max + Constant.Offset, 3].Value = "INDICE";
            worksheet.Cells[max + Constant.Offset, 4].Value = "VALOARE";
        }

        public static void PopulateTablesOffice(ExcelWorksheet worksheet, int max)
        {
            worksheet.Columns[0].SetWidth(135, LengthUnit.Pixel);
            worksheet.Columns[1].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[2].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[3].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[4].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[5].SetWidth(110, LengthUnit.Pixel);
            worksheet.Columns[6].SetWidth(120, LengthUnit.Pixel);
            worksheet.Columns[7].SetWidth(90, LengthUnit.Pixel);
            worksheet.Columns[8].SetWidth(140, LengthUnit.Pixel);

            worksheet.Cells[0, 0].Value = "ID";
            worksheet.Cells[0, 1].Value = "Data";
            worksheet.Cells[0, 2].Value = "Ora incepere";
            worksheet.Cells[0, 3].Value = "Ora sfarsit";
            worksheet.Cells[0, 4].Value = "Ora total";

            worksheet.Cells[max + Constant.Offset, 0].Value = "PLATA";
            worksheet.Cells[max + Constant.Offset, 1].Value = "-";
            worksheet.Cells[max + Constant.Offset, 2].Value = "PRET/H";
            worksheet.Cells[max + Constant.Offset, 3].Value = "INDICE";
            worksheet.Cells[max + Constant.Offset, 4].Value = "VALOARE";
            worksheet.Cells[max + Constant.Offset + 1, 0].Value = "TOTAL:";
        }
    }
}
