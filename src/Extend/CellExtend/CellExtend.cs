/*
 类：CellExtend
 描述：Cell扩展方法
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;
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
            }
            else
            {
                if (value.GetType().FullName.Equals("System.Byte[]"))
                {
                    int pictureIdx = cell.Sheet.Workbook.AddPicture((Byte[]) value, PictureType.PNG);
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
                    TypeCode valueTypeCode = Type.GetTypeCode(value.GetType());
                    switch (valueTypeCode)
                    {
                        case TypeCode.String: //字符串类型
                            cell.SetCellValue(Convert.ToString(value));
                            break;

                        case TypeCode.DateTime: //日期类型
                            cell.SetCellValue(Convert.ToDateTime(value));
                            break;

                        case TypeCode.Boolean: //布尔型
                            cell.SetCellValue(Convert.ToBoolean(value));
                            break;

                        case TypeCode.Int16: //整型
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Byte:
                        case TypeCode.Single: //浮点型
                        case TypeCode.Double:
                        case TypeCode.Decimal:
                        case TypeCode.UInt16: //无符号整型
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                            cell.SetCellValue(Convert.ToDouble(value));
                            break;

                        default:
                            cell.SetCellValue(string.Empty);
                            break;
                    }
                }
            }
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