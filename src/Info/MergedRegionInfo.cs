/*
 类：MergedRegionInfo
 描述：合并区域信息
 编 码 人：韩兆新 日期：2015年05月01日
 修改记录：

*/

namespace NPOI.Extend
{
    public class MergedRegionInfo
    {
        public MergedRegionInfo(int index, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            Index = index;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstCol = firstCol;
            LastCol = lastCol;
        }

        public int Index { get; set; }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
        public int FirstCol { get; set; }
        public int LastCol { get; set; }
    }
}