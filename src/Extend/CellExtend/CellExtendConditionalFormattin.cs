using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace NPOI.Extend
{
    public static partial class CellExtend
    {
        #region 1.0 添加条件格式规则

        /// <summary>
        ///     添加条件格式规则
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="cfrs">条件格式规则</param>
        public static void AddConditionalFormattingRules(this ICell cell, IConditionalFormattingRule[] cfrs)
        {
            CellRangeAddress[] regions =
            {
                new CellRangeAddress(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex)
            };
            cell.Sheet.SheetConditionalFormatting.AddConditionalFormatting(regions, cfrs);
        }

        #endregion

        #region 1.1 获取条件格式规则

        /// <summary>
        ///     获取条件格式规则
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static IConditionalFormattingRule[] GetConditionalFormattingRules(this ICell cell)
        {
            var cfrList = new List<IConditionalFormattingRule>();

            ISheetConditionalFormatting scf = cell.Sheet.SheetConditionalFormatting;
            for (int i = 0; i < scf.NumConditionalFormattings; i++)
            {
                IConditionalFormatting cf = scf.GetConditionalFormattingAt(i);
                if (cell.ExistConditionalFormatting(cf))
                {
                    for (int j = 0; j < cf.NumberOfRules; j++)
                    {
                        cfrList.Add(cf.GetRule(j));
                    }
                }
            }
            return cfrList.ToArray();
        }

        #endregion

        #region 3.0 判断单元格是否存在条件格式

        /// <summary>
        ///     单元格是否存在条件格式
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="cf">条件格式</param>
        /// <returns></returns>
        private static bool ExistConditionalFormatting(this ICell cell, IConditionalFormatting cf)
        {
            CellRangeAddress[] cfRangeAddrs = cf.GetFormattingRanges();
            foreach (CellRangeAddress cfRangeAddr in cfRangeAddrs)
            {
                if (cell.RowIndex >= cfRangeAddr.FirstRow && cell.RowIndex <= cfRangeAddr.LastRow
                    && cell.ColumnIndex >= cfRangeAddr.FirstColumn && cell.ColumnIndex <= cfRangeAddr.LastColumn)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}