using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class SharedStrings : IList<string>, IDisposable
    {
        private string xmlSource;

        private Dictionary<ulong, string> dic = new Dictionary<ulong, string>();

        public int Count
        {
            get
            {
                return dic.Count;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return true;
            }
        }

        public string this[int index]
        {
            get
            {
                return this[Convert.ToUInt64(index)];
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public SharedStrings(string XmlSource)
        {
            xmlSource = XmlSource;
            var xd = new XmlDocument();
            xd.LoadXml(xmlSource.Replace("standalone=\"true\"", "standalone=\"yes\""));
            var t = xd.GetElementsByTagName("t");
            ulong i = 0;
            foreach (XmlNode x in t)
                dic.Add(i++, x.InnerText);
            xmlSource = null;
            GC.Collect();
        }

        public string this[ulong index]
        {
            get
            {
                return dic[index];
            }
            set
            {
                dic[index] = value;
            }
        }

        public void Dispose()
        {
            dic.Clear();
            GC.Collect();
        }

        public int IndexOf(string item)
        {
            var index = dic.Where(x => x.Value == item)
                .Select(x => x.Key)
                .FirstOrDefault();
            return Convert.ToInt32(index);
        }

        public void Insert(int index, string item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            dic.Remove(Convert.ToUInt64(index));
        }

        public void Add(string item)
        {
            if (dic.Count == 0)
            {
                dic.Add(0, item);
            }
            else
            {
                var last = dic.Max(x => x.Key);
                dic.Add(last + 1, item);
            }
        }

        public void Clear()
        {
            dic.Clear();
        }

        public bool Contains(string item)
        {
            return dic.Any(x => x.Value == item);
        }

        public void CopyTo(string[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(string item)
        {
            var keys = dic.Where(x => x.Value == item)
                .Select(x => x.Key)
                .ToList();
            if (keys.Count == 0)
                return false;
            foreach (var x in keys)
                dic.Remove(x);
            return true;
        }

        public IEnumerator<string> GetEnumerator()
        {
            return dic.Select(x => x.Value)
                .ToList()
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
