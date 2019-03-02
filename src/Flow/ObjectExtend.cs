using System;
using System.Collections.Generic;
using System.Text;

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
