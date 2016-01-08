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
            xd.LoadXml(xmlSource);
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
                throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        public void Add(string item)
        {
            throw new NotImplementedException();
        }

        public void Clear()
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
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
