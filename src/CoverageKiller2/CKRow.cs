using static CoverageKiller2.CKCellReference;
namespace CoverageKiller2
{
    public class CKRow : CKCells
    {
        internal int _lastIndex;


        public CKRow(CKTable table, int index) :
            base(table, new CKCellReference(table, index, ReferenceTypes.Row))
        {
        }
        public CKCells Cells => this;

        // Property to access row's index
        public int Index => oops;


    }
}


