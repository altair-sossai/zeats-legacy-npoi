using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Zeats.Legacy.NPOI.Extensions
{
    public static class RowExtensions
    {
        public static Dictionary<string, int> MapHeader(this IRow row)
        {
            var header = new Dictionary<string, int>();

            for (var column = 0; column < row.LastCellNum; column++)
            {
                var cell = row.GetCell(column);
                var value = cell.StringCellValue;

                if (string.IsNullOrEmpty(value) || header.ContainsKey(value))
                    continue;

                header.Add(value, column);
            }

            return header;
        }

        public static string StringValue(this IRow row, int column, string defaultValue = null)
        {
            if (row == null)
                return defaultValue;

            return row?.GetCell(column)?.StringCellValue ?? defaultValue;
        }

        public static bool BooleanValue(this IRow row, int column, bool defaultValue = false)
        {
            if (row == null)
                return defaultValue;

            var cell = row.GetCell(column);
            var cellType = cell.CellType;

            switch (cellType)
            {
                case CellType.Numeric:
                    return cell.BooleanLikeNumericValue(defaultValue);

                case CellType.String:
                    return cell.BooleanLikeStringValue(defaultValue);

                case CellType.Boolean:
                    return cell.BooleanCellValue;
            }

            return defaultValue;
        }
    }
}