/*
 类：NPOIHelper
 描述：NPOI操作助手类
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.IO;
using NPOI.SS.UserModel;

namespace NPOI.Extend
{
    public static class NPOIHelper
    {
        #region 1.0 加载模板,获取IWorkbook对象

        /// <summary>
        ///     加载模板,获取IWorkbook对象
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static IWorkbook LoadWorkbook(string filePath)
        {
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read)) //读入excel模板
            {
                return WorkbookFactory.Create(fileStream);
            }
        }

        #endregion
    }
}