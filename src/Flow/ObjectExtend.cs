/*
 类：ObjectExtend
 描述：Object扩展
 编 码 人：韩兆新 日期：2019年03月02日
 修改记录：

*/

namespace NPOI.Extend
{
    public static class ObjectExtend
    {
        public static Case<T> Case<T>(this T value, params T[] asserts)
        {
            return new Case<T>(value, asserts);
        }
    }
}
