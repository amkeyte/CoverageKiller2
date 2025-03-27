// CKCellRefLinear: Stubbed linear cell reference implementation
using System;
using System.Collections.Generic;

namespace CoverageKiller2.DOM
{
    public class CKCellRefLinear : ICellRef<IDOMObject>
    {
        public int Start => throw new NotImplementedException();
        public int End => throw new NotImplementedException();

        public IEnumerable<int> WordCells => throw new NotImplementedException();
        public int GridX1 => throw new NotImplementedException();
        public int GridY1 => throw new NotImplementedException();
        public int GridX2 => throw new NotImplementedException();
        public int GridY2 => throw new NotImplementedException();

        private CKCellRefLinear() { }

        public static CKCellRefLinear ForCells(int start, int end)
        {
            throw new NotImplementedException();
        }

        public static CKCellRefLinear ForCells(CKRange range)
        {
            throw new NotImplementedException();
        }

        public static CKCellRefLinear ForCell(int index)
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return "CKCellRefLinear (stub)";
        }
    }
}
