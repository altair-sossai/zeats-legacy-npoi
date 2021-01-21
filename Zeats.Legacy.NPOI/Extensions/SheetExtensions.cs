using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Zeats.Legacy.NPOI.Extensions
{
    public static class SheetExtensions
    {
        public static Dictionary<string, int> MapHeader(this ISheet sheet, int row = 0)
        {
            return sheet.GetRow(row).MapHeader();
        }
    }
}