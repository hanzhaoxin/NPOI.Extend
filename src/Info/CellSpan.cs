/*
 结构：CellSpan
 描述：单元格跨度信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

namespace NPOI.Extend
{
    public struct CellSpan
    {
        public CellSpan(int rowSpan, int colSpan)
        {
            this._rowSpan = rowSpan;
            this._colSpan = colSpan;
        }

        private int _rowSpan;

        public int RowSpan
        {
            get { return _rowSpan; }
            set { _rowSpan = value; }
        }

        private int _colSpan;

        public int ColSpan
        {
            get { return _colSpan; }
            set { _colSpan = value; }
        }
    }
}