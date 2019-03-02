using System;
using System.Collections.Generic;
using System.Text;

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
