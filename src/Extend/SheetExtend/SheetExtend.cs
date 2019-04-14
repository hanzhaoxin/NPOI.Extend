/*
 类：SheetExtend
 描述：Sheet扩展方法
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：
    1: 修改人：韩兆新  日期：2015年05月01日
       修改内容：添加了获取合并区域信息列表的扩展方法GetAllMergedRegionInfos();
                 添加了添加合并区域的扩展方法AddMergedRegion();
    2:修改人：韩兆新  日期：2015年11月21日
       修改内容：修改Bug：插入或删除行时，图片的移动问题。

*/

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.Extend.Formula;
using NPOI.SS.UserModel;

namespace NPOI.Extend
{
    public static partial class SheetExtend
    {
        #region 1.0 插入行

        /// <summary>
        ///     插入行
        /// </summary>
        /// <param name="sheet">要插入行的sheet</param>
        /// <param name="rowIndex">行标</param>
        /// <returns></returns>
        public static IRow InsertRow(this ISheet sheet, int rowIndex)
        {
            return sheet.InsertRows(rowIndex, 1)[0];
        }

        /// <summary>
        ///     插入行
        /// </summary>
        /// <param name="sheet">要插入行的sheet</param>
        /// <param name="rowIndex">行标</param>
        /// <param name="rowsCount">行数</param>
        /// <returns></returns>
        public static IRow[] InsertRows(this ISheet sheet, int rowIndex, int rowsCount)
        {
            if (rowIndex <= sheet.LastRowNum)
            {
                sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowsCount, true, false);
            }

            var rowList = new List<IRow>();
            for (int i = 0; i < rowsCount; i++)
            {
                IRow row = sheet.CreateRow(rowIndex + i);
            }
            sheet.MovePictures(rowIndex,null,null,null,moveRowCount:rowsCount);
            return rowList.ToArray();
        }

        #endregion

        #region 1.1 删除行

        /// <summary>
        ///     删除行
        /// </summary>
        /// <param name="sheet">要删除行的sheet</param>
        /// <param name="rowIndex">行标</param>
        public static int RemoveRow(this ISheet sheet, int rowIndex)
        {
            return sheet.RemoveRows(rowIndex, rowIndex);
        }

        /// <summary>
        ///     删除行
        /// </summary>
        /// <param name="sheet">要删除行的sheet</param>
        /// <param name="startRowIndex">开始行行标</param>
        /// <param name="endRowIndex">结束行行标</param>
        public static int RemoveRows(this ISheet sheet, int startRowIndex, int endRowIndex)
        {
            int span = endRowIndex - startRowIndex + 1;
            //删除合并区域...
            sheet.RemoveMergedRegions(startRowIndex, endRowIndex, null, null);
            //删除图片...
            sheet.RemovePictures(startRowIndex, endRowIndex, null, null);
            //删除行
            for (int i = endRowIndex; i >= startRowIndex; i--)
            {
                IRow row = sheet.GetRow(i);
                sheet.RemoveRow(row);
            }
            if ((endRowIndex + 1) <= sheet.LastRowNum)
            {
                sheet.ShiftRows((endRowIndex + 1), sheet.LastRowNum, -span, true, false);
                sheet.MovePictures(endRowIndex + 1,null,null,null,moveRowCount:-span);
            }
            return span;
        }

        #endregion

        #region 1.2 复制行

        /// <summary>
        ///     复制行
        /// </summary>
        /// <param name="sheet">要复制行的sheet</param>
        /// <param name="rowIndex">行标</param>
        public static int CopyRow(this ISheet sheet, int rowIndex)
        {
            return sheet.CopyRows(rowIndex, rowIndex);
        }

        /// <summary>
        ///     复制行
        /// </summary>
        /// <param name="sheet">要复制行的sheet</param>
        /// <param name="startRowIndex">开始行行标</param>
        /// <param name="endRowIndex">结束行行标</param>
        public static int CopyRows(this ISheet sheet, int startRowIndex, int endRowIndex)
        {
            int span = endRowIndex - startRowIndex + 1;

            int newStartRowIndex = startRowIndex + span;
            //插入空行
            sheet.InsertRows(newStartRowIndex, span);
            //复制行
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                IRow sourceRow = sheet.GetRow(i);
                IRow targetRow = sheet.GetRow(i + span);

                targetRow.Height = sourceRow.Height;
                targetRow.ZeroHeight = sourceRow.ZeroHeight;

                #region 复制单元格

                foreach (ICell sourceCell in sourceRow.Cells)
                {
                    ICell targetCell = targetRow.GetCell(sourceCell.ColumnIndex);
                    if (null == targetCell)
                    {
                        targetCell = targetRow.CreateCell(sourceCell.ColumnIndex);
                    }
                    if (null != sourceCell.CellStyle) targetCell.CellStyle = sourceCell.CellStyle;
                    if (null != sourceCell.CellComment) targetCell.CellComment = sourceCell.CellComment;
                    if (null != sourceCell.Hyperlink) targetCell.Hyperlink = sourceCell.Hyperlink;
                    IConditionalFormattingRule[] cfrs = sourceCell.GetConditionalFormattingRules(); //复制条件样式
                    if (null != cfrs && cfrs.Length > 0)
                    {
                        targetCell.AddConditionalFormattingRules(cfrs);
                    }
                    targetCell.SetCellType(sourceCell.CellType);

                    #region 复制值

                    switch (sourceCell.CellType)
                    {
                        case CellType.Numeric:
                            targetCell.SetCellValue(sourceCell.NumericCellValue);
                            break;
                        case CellType.String:
                            targetCell.SetCellValue(sourceCell.RichStringCellValue);
                            break;
                        case CellType.Formula:
                            var formula = CopyFormula(sourceCell.CellFormula,span);
                            targetCell.SetCellFormula(formula);
                            break;
                        case CellType.Blank:
                            targetCell.SetCellValue(sourceCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            targetCell.SetCellValue(sourceCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                            break;
                    }

                    #endregion
                }

                #endregion
            }
            //获取模板行内的合并区域
            List<MergedRegionInfo> regionInfoList = sheet.GetMergedRegionInfos(startRowIndex, endRowIndex, null,
                null);
            //复制合并区域
            foreach (MergedRegionInfo regionInfo in regionInfoList)
            {
                regionInfo.FirstRow += span;
                regionInfo.LastRow += span;
                sheet.AddMergedRegion(regionInfo);
            }
            //获取模板行内的图片
            List<PictureInfo> picInfoList = sheet.GetAllPictureInfos(startRowIndex, endRowIndex, null, null);
            //复制图片
            foreach (PictureInfo picInfo in picInfoList)
            {
                picInfo.MaxRow += span;
                picInfo.MinRow += span;
                sheet.AddPicture(picInfo);
            }

            return span;
        }

        #endregion

        #region 3.0 判断区域内部或交叉

        /// <summary>
        ///     判断是内部或交叉
        /// </summary>
        /// <param name="rangeMinRow"></param>
        /// <param name="rangeMaxRow"></param>
        /// <param name="rangeMinCol"></param>
        /// <param name="rangeMaxCol"></param>
        /// <param name="pictureMinRow"></param>
        /// <param name="pictureMaxRow"></param>
        /// <param name="pictureMinCol"></param>
        /// <param name="pictureMaxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >=
                     Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                    (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >=
                     Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
        }

        #endregion

        private static string CopyFormula(string formula, int span)
        {
            FormulaContext context = new FormulaContext(formula);
            context.Parse();
            return context.Formula.ToString(part =>
            {
                if (part.Type.Equals(PartType.Formula))
                {
                    Regex regex = new Regex(@"([A-Z]+)(\d+)");
                    return regex.Replace(part.ToString(), (m) => $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value) + span}");
                }
                else
                {
                    return part.ToString();
                }
            });
        }
    }
}