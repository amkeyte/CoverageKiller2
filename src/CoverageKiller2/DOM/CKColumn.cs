using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    public class CKColumn : IEnumerable<CKCell>
    {
        public CKTable Parent { get; private set; }

        private IEnumerable<CKCell> _cells;
        public CKColumn(IEnumerable<CKCell> columnCells)
        {
            _cells = columnCells;
        }

        public CKColumn(CKColumns cKColumns, int index)
        {
            throw new NotImplementedException();
        }

        public bool IsDirty => _cells.Any(c => c.IsDirty);

        public int Index => Tracer.Trace(Index);

        // Example: Method to select the entire column
        public Tracer Tracer = new Tracer(typeof(CKColumn));

        //public void Delete()
        //{
        //    Tracer.Log("Deleting column",
        //        new DataPoints()
        //            .Add(nameof(Index), Index)
        //            .Add("Header", Cells[1].Text)
        //            .Add("Contents", Cells
        //                .Where(c => c.RowIndex > 1)
        //                .Select(c => c.Text)
        //                .Aggregate((current, next) => current + ", " + next))
        //        );



        //    COMObject.Delete();
        //}

        public IEnumerator<CKCell> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}
