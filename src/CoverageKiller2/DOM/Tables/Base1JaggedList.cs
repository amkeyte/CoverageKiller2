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
        private readonly List<Base1List<T>> _rows_0 = new List<Base1List<T>>();

        public Base1JaggedList() { }

        public Base1JaggedList(List<List<T>> list)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));
            foreach (var row in list)
            {
                if (row == null) throw new ArgumentException("Row list cannot contain null elements.", nameof(list));
                _rows_0.Add(new Base1List<T>(row));
            }
        }

        public T SafeGet(int row, int col)
        {
            if (row < 1 || row > Count) return default;
            var rowList = _rows_0[row - 1];
            if (rowList == null || col < 1 || col > rowList.Count) return default;

            return rowList[col];
        }
        /// <summary>
        /// Returns the largest column count across all rows.
        /// </summary>
        public int LargestRowCount => _rows_0.Count == 0
            ? 0
            : _rows_0.Max(r => r?.Count ?? 0);

        public void Add(Base1List<T> row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows_0.Add(row);
        }

        public void Insert(int index, Base1List<T> row)
        {
            if (index < 1 || index > Count + 1)
                throw new ArgumentOutOfRangeException(nameof(index));
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows_0.Insert(index - 1, row);
        }

        /// <summary>
        /// Dumps the contents of a Base1JaggedList&lt;T&gt;, projecting each cell through a delegate.
        /// </summary>
        /// <typeparam name="T">The type of each cell in the jagged list.</typeparam>
        /// <param name="projector">A projection function to convert each cell to a string.</param>
        /// <param name="message">Optional label for the dump output.</param>
        /// <returns>A formatted string of the dumped grid.</returns>
        public string DumpGrid(Func<T, string> projector, string message = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"\n{message ?? "Grid Dump"}\n");
            sb.AppendLine("**********************");

            foreach (var row in _rows_0)
            {
                var line = row.Select(item => projector(item) ?? "[NULL]");
                sb.AppendLine(string.Join(" | ", line));
            }

            sb.AppendLine("**********************");
            return sb.ToString();
        }

        ///// <summary>
        ///// Dumps the contents of a Base1JaggedList&lt;T&gt; as a tree structure using a projection.
        ///// </summary>
        ///// <typeparam name="T">Type of the list item.</typeparam>
        ///// <param name="projector">A projection function that converts each element into a string.</param>
        ///// <param name="label">Optional label to prefix the tree.</param>
        ///// <returns>A multi-line string representing the tree structure of the grid.</returns>
        //public string DumpTree(Func<T, string> projector, string message = null)
        //{
        //    var sb = new StringBuilder();
        //    sb.AppendLine($"\n\n{message ?? "Tree Dump"}");
        //    sb.AppendLine("**********************");

        //    foreach (var row in _rows_0)
        //    {
        //        var line = row.Select(item => projector(item) ?? "[NULL]");


        //        var value = projector(row[colIndex]) ?? "[NULL]";
        //        sb.AppendLine($"│   ├─ Col {colIndex}: {value}");
        //    }


        //    return sb.ToString();


        //    //for (int rowIndex = 1; rowIndex <= _rows_0.Count; rowIndex++)
        //    //{




        //    //var row = _rows_0[rowIndex];
        //    //sb.AppendLine($"├─ Row {rowIndex}");

        //    //for (int colIndex = 1; colIndex <= row.Count; colIndex++)
        //    //{
        //    //    var value = projector(row[colIndex]) ?? "[NULL]";
        //    //    sb.AppendLine($"│   ├─ Col {colIndex}: {value}");
        //    //}

        //}


        public void RemoveAt(int index)
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            _rows_0.RemoveAt(index - 1);
        }

        public int IndexOf(Base1List<T> row)
        {
            var idx = _rows_0.IndexOf(row);
            return idx < 0 ? -1 : idx + 1;
        }

        public int Count => _rows_0.Count;

        public Base1List<T> this[int index]
        {
            get
            {
                if (index < 1) throw new ArgumentException("Index is 1 based.");
                if (index > Count) throw new ArgumentOutOfRangeException(nameof(index));
                return _rows_0[index - 1];
            }
        }

        public IEnumerator<Base1List<T>> GetEnumerator() => _rows_0.GetEnumerator();
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
                //Log.Debug($" Base1List has item count: {_items_0.Count}; trying to pull index {index} #########################################");
                return _items_0[index - 1];

            }
        }
        /// <summary>
        /// Dumps the contents of the Base1List, each entry as-is, separated by newlines.
        /// </summary>
        /// <param name="message">Optional message to prepend.</param>
        /// <returns>The formatted dump string.</returns>
        /// <remarks>Version: CK2.00.01.0038</remarks>
        public string Dump(string message = null)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(message))
                sb.AppendLine($"\n{message}\n");

            sb.AppendLine("**********************");

            foreach (var item in this)
            {
                string output = item?.ToString() ?? "[NULL]";
                sb.AppendLine(output);
            }

            sb.AppendLine("**********************");
            return sb.ToString();
        }
        /// <summary>
        /// Dumps the contents of the Base1List using a projection delegate, each entry on a new line.
        /// </summary>
        /// <param name="projector">Delegate that extracts a string from each item.</param>
        /// <param name="message">Optional message to prepend.</param>
        /// <returns>The formatted dump string.</returns>
        /// <remarks>Version: CK2.00.01.0039</remarks>
        public string Dump(Func<T, string> projector, string message = null)
        {
            if (projector == null) throw new ArgumentNullException(nameof(projector));

            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(message))
                sb.AppendLine($"\n{message}\n");

            sb.AppendLine("**********************");

            foreach (var item in this)
            {
                string output = projector(item) ?? "[NULL]";
                sb.AppendLine(output);
            }

            sb.AppendLine("**********************");
            return sb.ToString();
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
