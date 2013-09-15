using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MRUList<T> : IEnumerable<T>
    {
        private int size;
        private List<T> list;

        public MRUList(int size)
        {
            if (size < 0) throw new ArgumentException("Negative size not allowed", "size");
            this.size = size;
            this.list = new List<T>();
        }

        public MRUList(int size, IEnumerable<T> items)
        {
            if (size < 0) throw new ArgumentException("Negative size not allowed", "size");
            this.size = size;
            this.list = new List<T>(items);
            _Resize();
        }

        public void Touch(T item)
        {
            if (list.Contains(item))
            {
                list.Remove(item);
            }
            list.Insert(0, item);
            _Resize();
        }

        public void Clear()
        {
            list.Clear();
        }

        public IEnumerator<T> GetEnumerator()
        {
            return list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private void _Resize()
        {
            while (list.Count > size)
            {
                list.RemoveAt(list.Count - 1);
            }
        }
    }
}
