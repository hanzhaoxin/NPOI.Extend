/*
 类：RowExtend
 描述：Row扩展方法
 编 码 人：韩兆新 日期：2015年05月15日
 修改记录：

*/

using NPOI.SS.UserModel;

namespace NPOI.Extend
{
    public static class RowExtend
    {
        #region 1.0 清除单元格内容

        /// <summary>
        ///     清除内容
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>行</returns>
        public static IRow ClearContent(this IRow row)
        {
            foreach (ICell cell in row.Cells)
            {
                cell.SetCellValue(string.Empty);
            }
            return row;
        }

        #endregion
    }
}