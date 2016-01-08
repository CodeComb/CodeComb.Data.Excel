using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class RowNumber 
    {
        public RowNumber(string number = null)
        {
            value = number;
        }

        private string value = null;

        public static bool operator >(RowNumber a, RowNumber b)
        {
            return string.Compare(a.value, b.value) > 0; 
        }

        public static bool operator <(RowNumber a, RowNumber b)
        {
            return string.Compare(a.value, b.value) < 0;
        }

        public static bool operator >=(RowNumber a, RowNumber b)
        {
            return string.Compare(a.value, b.value) >= 0;
        }

        public static bool operator <=(RowNumber a, RowNumber b)
        {
            return string.Compare(a.value, b.value) <= 0;
        }

        public static bool operator !=(RowNumber a, RowNumber b)
        {
            return a.value != b.value;
        }

        public static bool operator ==(RowNumber a, RowNumber b)
        {
            return a.value == b.value;
        }

        public static RowNumber operator ++(RowNumber a)
        {
            var tmp = FromNumberSystem26(a.value);
            tmp++;
            return new RowNumber(ToNumberSystem26(tmp));
        }

        public static RowNumber operator --(RowNumber a)
        {
            var tmp = FromNumberSystem26(a.value);
            tmp--;
            return new RowNumber(ToNumberSystem26(tmp));
        }

        public static implicit operator RowNumber(string s)
        {
            return new RowNumber(s);
        }

        public static implicit operator string(RowNumber rn)
        {
            return rn.value;
        }

        public static string ToNumberSystem26(ulong n)
        {
            string s = string.Empty;
            while (n > 0)
            {
                var m = n % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                n = (n - m) / 26;
            }
            return s;
        }

        /// <summary>
        /// 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]。
        /// </summary>
        /// <param name="s">26进制表示（如果无效，则返回0）。</param>
        /// <returns>自然数。</returns>
        public static ulong FromNumberSystem26(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            ulong n = 0;
            for (ulong i = (ulong)s.Length - 1, j = 1ul; i >= 0; i--, j *= 26)
            {
                char c = Char.ToUpper(s[i]);
                if (c < 'A' || c > 'Z') return 0;
                n += (c - 64ul) * j;
            }
            return n;
        }
    }
}
