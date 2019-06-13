/*
 类：SheetExtend
 描述：Sheet扩展方法
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：
    1: 修改人：韩兆新  日期：2015年11月21日
       修改内容：添加了移动合并区域的方法
   

*/
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace NPOI.Extend
{
    public static partial class SheetExtend
    {
        #region 1.0 添加合并区域

        /// <summary>
        ///     sheet中添加合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="regionInfo"></param>
        public static void AddMergedRegion(this ISheet sheet, MergedRegionInfo regionInfo)
        {
            var region = new CellRangeAddress(regionInfo.FirstRow, regionInfo.LastRow, regionInfo.FirstCol,
                regionInfo.LastCol);
            sheet.AddMergedRegion(region);
        }

        #endregion

        #region 1.1 获取合并区域信息

        /// <summary>
        ///     获取sheet中包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this ISheet sheet)
        {
            return sheet.GetMergedRegionInfos(null, null, null, null);
        }

        /// <summary>
        ///     获取sheet中指定区域包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this ISheet sheet, int? minRow, int? maxRow,
            int? minCol, int? maxCol, bool onlyInternal = true)
        {
            var regionInfoList = new List<MergedRegionInfo>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                if (null == range)
                {
                    continue;
                }

                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn, onlyInternal))
                {
                    regionInfoList.Add(new MergedRegionInfo(i, range.FirstRow, range.LastRow, range.FirstColumn,
                        range.LastColumn));
                }
            }
            return regionInfoList;
        }

        #endregion

        #region 1.3 删除合并区域

        /// <summary>
        ///     删除合并区域
        /// </summary>
        /// <param name="sheet"></param>
        public static void RemoveMergedRegions(this ISheet sheet)
        {
            sheet.RemoveMergedRegions(null, null, null, null);
        }

        /// <summary>
        ///     删除指定区域内的合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        public static void RemoveMergedRegions(this ISheet sheet, int? minRow, int? maxRow,
            int? minCol, int? maxCol)
        {
            List<MergedRegionInfo> regionInfoList = sheet.GetMergedRegionInfos(minRow, maxRow, minCol, maxCol, false);
            foreach (MergedRegionInfo regionInfo in regionInfoList.OrderByDescending(r => r.Index))
            {
                sheet.RemoveMergedRegion(regionInfo.Index);
            }
        }

        #endregion

        #region 1.4 移动合并区域

        /// <summary>
        ///     移动sheet中所有合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="moveRowCount"></param>
        /// <param name="moveColCount"></param>
        public static void MoveMergedRegions(this ISheet sheet, int moveRowCount, int moveColCount = 0)
        {
            sheet.MoveMergedRegions(null, null, null, null, moveRowCount, moveColCount);
        }

        /// <summary>
        ///     移动指定区域内的合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="moveRowCount"></param>
        /// <param name="moveColCount"></param>
        public static void MoveMergedRegions(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
            int moveRowCount, int moveColCount = 0)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn, true))
                {
                    range.FirstRow += moveRowCount;
                    range.LastRow += moveRowCount;
                    range.FirstColumn += moveColCount;
                    range.LastColumn += moveColCount;
                }
            }
        }

        #endregion
    }
}