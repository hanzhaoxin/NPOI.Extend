/*
 类：Case<T>
 描述：Case
 编 码 人：韩兆新 日期：2019年03月02日
 修改记录：

*/

using System;
using System.Collections.Generic;

namespace NPOI.Extend
{
    public class Case<T>
    {
        private T _value;

        private ICollection<T> _asserts;

        public Case(T value, T[] asserts)
        {
            _value = value;
            _asserts = asserts;
        }

        public T Do(Action action)
        {
            if (_asserts.Contains(_value))
            {
                action();
            }
            return _value;
        }
    }
}
