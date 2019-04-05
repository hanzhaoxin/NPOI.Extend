using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Extend.CellExtend
{
    public static partial class CellExtend
    {
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="fromCell">起始单元格</param>
        /// <param name="toCell">终止单元格</param>
        /// <param name="isExpand">扩充模式</param>
        public static void Merge(this ICell fromCell, ICell toCell,bool isExpand = false)
        {
            if (!fromCell.Sheet.Equals(toCell.Sheet))
            {
                throw new Exception($"cells is not in same sheet!");
            }
            var sheet = fromCell.Sheet;
            var fromRange = fromCell.GetRangeInfo();
            var toRange = toCell.GetRangeInfo();
            var firstRowIndex = Math.Min(fromRange.FirstRow, toRange.FirstRow);
            var firstColIndex = Math.Min(fromRange.FirstCol, toRange.FirstCol);
            var lastRowIndex = Math.Max(fromRange.LastRow, toRange.LastRow);
            var lastColIndex = Math.Max(fromRange.LastCol, toRange.LastCol);
            var regionInfoList = sheet.GetMergedRegionInfos(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex, false);
            foreach (MergedRegionInfo regionInfo in regionInfoList)
            {
                if (isExpand)
                {
                    firstRowIndex = Math.Min(firstRowIndex, regionInfo.FirstRow);
                    firstColIndex = Math.Min(firstColIndex, regionInfo.FirstCol);
                    lastRowIndex = Math.Max(lastRowIndex, regionInfo.LastRow);
                    lastColIndex = Math.Max(lastColIndex, regionInfo.LastCol);
                }
                sheet.RemoveMergedRegion(regionInfo.Index);
            }
            var region = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
            fromCell.Sheet.AddMergedRegion(region);
        }

        /// <summary>
        /// 获取合并信息
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>合并信息</returns>
        private static MergedRegionInfo GetRangeInfo(this ICell cell)
        {
            var sheet = cell.Sheet;
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                if (range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return new MergedRegionInfo(i, range.FirstRow, range.LastRow, range.FirstColumn,
                        range.LastColumn);
                }
            }
            return new MergedRegionInfo(-1,cell.RowIndex,cell.RowIndex,cell.ColumnIndex,cell.ColumnIndex);
        }
    }
}
