using System.Linq;
using NPOI.SS.UserModel;

namespace Zeats.Legacy.NPOI.Extensions
{
    public static class CellExtensions
    {
        public static ICell BoldText(this ICell cell, string text)
        {
            var font = cell.Sheet.Workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.IsBold = true;

            cell.CellStyle = cell.Sheet.Workbook.CreateCellStyle();
            cell.CellStyle.SetFont(font);
            cell.SetCellValue(text);
            cell.CellStyle.FillBackgroundColor = IndexedColors.Aqua.Index;

            return cell;
        }

        public static void TextCenter(this ICell cell)
        {
            cell.CellStyle.Alignment = HorizontalAlignment.Center;
        }

        public static bool BooleanLikeNumericValue(this ICell cell, bool defaultValue = false)
        {
            var value = cell.NumericCellValue;

            switch (value)
            {
                case 0: return false;
                case 1: return true;
                default: return defaultValue;
            }
        }

        public static bool BooleanLikeStringValue(this ICell cell, bool defaultValue = false)
        {
            var value = cell.StringCellValue?.Trim().ToLower();

            var sim = new[] {"1", "sim", "s", "yes", "y", "true"};
            var nao = new[] {"0", "não", "nao", "n", "no", "y", "false"};

            if (sim.Contains(value))
                return true;

            if (nao.Contains(value))
                return false;

            return defaultValue;
        }
    }
}