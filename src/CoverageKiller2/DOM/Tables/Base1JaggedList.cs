using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Represents a jagged 1-based collection of Base1List&lt;T&gt;, useful for table-like structures.
    /// </summary>
    /// <typeparam name="T">The element type of the inner lists.</typeparam>
    /// <remarks>
    /// Version: CK2.00.01.0037
    /// </remarks>
    public class Base1JaggedList<T> : IReadOnlyList<Base1List<T>>
    {
        private readonly List<Base1List<T>> _rows = new List<Base1List<T>>();

        public Base1JaggedList() { }

        public Base1JaggedList(List<List<T>> list)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));
            foreach (var row in list)
            {
                if (row == null) throw new ArgumentException("Row list cannot contain null elements.", nameof(list));
                _rows.Add(new Base1List<T>(row));
            }
        }

        /// <summary>
        /// Returns the largest column count across all rows.
        /// </summary>
        public int LargestRowCount => _rows.Count == 0
            ? 0
            : _rows.Max(r => r?.Count ?? 0);

        public void Add(Base1List<T> row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        public void Insert(int index, Base1List<T> row)
        {
            if (index < 1 || index > Count + 1)
                throw new ArgumentOutOfRangeException(nameof(index));
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Insert(index - 1, row);
        }

        public T Get2D(int row, int col) => this[row][col];

        public void RemoveAt(int index)
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            _rows.RemoveAt(index - 1);
        }

        public int IndexOf(Base1List<T> row)
        {
            var idx = _rows.IndexOf(row);
            return idx < 0 ? -1 : idx + 1;
        }

        public int Count => _rows.Count;

        public Base1List<T> this[int index]
        {
            get
            {
                if (index < 1) throw new ArgumentException("Index is 1 based.");
                if (index > Count) throw new ArgumentOutOfRangeException(nameof(index));
                return _rows[index - 1];
            }
        }

        public IEnumerator<Base1List<T>> GetEnumerator() => _rows.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// Represents a 1-based collection of elements of type T.
    /// </summary>
    /// <typeparam name="T">Element type.</typeparam>
    /// <remarks>
    /// Version: CK2.00.01.0038
    /// </remarks>
    public class Base1List<T> : IReadOnlyList<T>
    {
        private readonly List<T> _items_0 = new List<T>();

        public Base1List() { }
        public Base1List(IEnumerable<T> items) => _items_0.AddRange(items ?? Enumerable.Empty<T>());
        public Base1List(Base1List<T> items) => _items_0.AddRange(items ?? Enumerable.Empty<T>());
        public Base1List(IOrderedEnumerable<T> items) => _items_0.AddRange(items ?? Enumerable.Empty<T>());

        public void Add(T item) => _items_0.Add(item);

        public void Insert(int index, T item)
        {
            if (index < 1 || index > Count + 1)
                throw new ArgumentOutOfRangeException(nameof(index));
            _items_0.Insert(index - 1, item);
        }

        public void Insert(int index, IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                Insert(index++, item);
            }
        }

        public void RemoveAt(int index)
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            _items_0.RemoveAt(index - 1);
        }

        public int IndexOf(T item)
        {
            int idx_0 = _items_0.IndexOf(item);
            return idx_0 < 0 ? -1 : idx_0 + 1;
        }

        public int Count => _items_0.Count;

        public T this[int index]
        {
            get
            {
                if (index < 1) throw new ArgumentException("Index is 1 based.");
                if (index > Count) throw new ArgumentOutOfRangeException(nameof(index));
                return _items_0[index - 1];
            }
        }

        public IEnumerator<T> GetEnumerator() => _items_0.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal void Clear() => _items_0.Clear();

        public void ForEach(Action<T> action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            foreach (var item in _items_0)
            {
                action(item);
            }
        }

        /// <summary>
        /// Provides a debug string for inspecting the contents of this list.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.01.0039
        /// </remarks>
        public string DebugDump
        {
            get
            {
                var sb = new StringBuilder();
                for (int i = 1; i <= Count; i++)
                {
                    sb.AppendLine($"[{i}] {this[i]?.ToString()}");
                }
                return sb.ToString();
            }
        }
    }
}
