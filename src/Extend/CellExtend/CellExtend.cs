/*
 类：CellExtend
 描述：Cell扩展方法
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;
using System.Drawing;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace NPOI.Extend
{
    public static partial class CellExtend
    {
        #region 1.0 设置单元格值

        /// <summary>
        ///     设置单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="value">值</param>
        public static void SetValue(this ICell cell, object value)
        {
            if (null == cell)
            {
                return;
            }
            if (null == value)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            
            if (value is IRichTextString)
            {
                cell.SetCellValue((IRichTextString)value);
            }
            else if (value is Image)
            {
                int pictureIdx = cell.Sheet.Workbook.AddPicture(((Image)value).ToBuffer(), PictureType.PNG);
                IClientAnchor anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + cell.GetSpan().ColSpan;
                anchor.Row1 = cell.RowIndex;
                anchor.Row2 = cell.RowIndex + cell.GetSpan().RowSpan;
                IDrawing patriarch = cell.Sheet.CreateDrawingPatriarch();
                IPicture pic = patriarch.CreatePicture(anchor, pictureIdx);
            }
            else
            {
                value.GetType()
                .Case(typeof(String)).Do(() => cell.SetCellValue(Convert.ToString(value)))
                .Case(typeof(DateTime)).Do(() => cell.SetCellValue(Convert.ToDateTime(value)))
                .Case(typeof(Boolean)).Do(() => cell.SetCellValue(Convert.ToBoolean(value)))
                .Case(typeof(Int16), typeof(Int32), typeof(Int64), 
                    typeof(Byte), typeof(Single), typeof(Double), typeof(Decimal), 
                    typeof(UInt16), typeof(UInt32), typeof(UInt64)).Do(() => cell.SetCellValue(Convert.ToDouble(value)));
            }
        }


        /// <summary>
        ///     获取单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>值</returns>
        public static object GetValue(this ICell cell)
        {
            object value = string.Empty;
            cell.CellType
                .Case(CellType.String).Do(() => value = cell.StringCellValue)
                .Case(CellType.Numeric).Do(() => value = cell.NumericCellValue)
                .Case(CellType.Boolean).Do(() => value = cell.BooleanCellValue);

            return value;
        }

        #endregion

        #region 1.1 获取单元格CellSpan信息

        /// <summary>
        ///     获取cell的CellSpan信息
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        public static CellSpan GetSpan(this ICell cell)
        {
            var cellSpan = new CellSpan(1, 1);
            if (cell.IsMergedCell)
            {
                int regionsNum = cell.Sheet.NumMergedRegions;
                for (int i = 0; i < regionsNum; i++)
                {
                    CellRangeAddress range = cell.Sheet.GetMergedRegion(i);
                    if (range.FirstRow == cell.RowIndex && range.FirstColumn == cell.ColumnIndex)
                    {
                        cellSpan.RowSpan = range.LastRow - range.FirstRow + 1;
                        cellSpan.ColSpan = range.LastColumn - range.FirstColumn + 1;
                        break;
                    }
                }
            }
            return cellSpan;
        }

        #endregion
    }
}