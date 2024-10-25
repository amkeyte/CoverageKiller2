using Serilog;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2
{
    public class CKRow
    {
        internal int _lastIndex;

        internal static CKRow Create(CKRows parent, int index)

        {
            return new CKRow(parent, index);
        }

        internal Word.Row COMObject { get; private set; }
        public CKRows Parent { get; private set; }
        // Constructor to initialize CKRow with a Word.Row
        public CKRow(CKRows parent, int index)
        {
            Parent = parent;
            COMObject = Parent.COMObject[index];
        }
        public CKCells Cells => CKCells.Create(this);
        //public bool ContainsMerged => Cells.ContainsMerged;
        // Property to access row's index
        public int Index
        {
            get
            {
                _lastIndex = COMObject.Index;
                return _lastIndex;
            }
        }




        // Set or get the row's height
        public float Height
        {
            get => COMObject.Height;
            set => COMObject.Height = value;
        }

        // Set or get the row's height rule (auto, at least, or exact)
        public Word.WdRowHeightRule HeightRule
        {
            get => COMObject.HeightRule;
            set => COMObject.HeightRule = value;
        }

        // Property to get or set row shading
        public Word.WdColor ShadingBackgroundColor
        {
            get => (Word.WdColor)COMObject.Shading.BackgroundPatternColor;
            set => COMObject.Shading.BackgroundPatternColor = (Word.WdColor)value;
        }


        // Deletes the row
        public void Delete()
        {
            Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                nameof(CKRow), nameof(Delete),
                nameof(Index), _lastIndex));


            //$"{nameof(COMObject)}({nameof(CKCell)}.{nameof(CKCell.RowIndex)}) --> ", Cells[1].RowIndex,
            //$"{nameof(COMObject)}({nameof(CKCell)}[1].{nameof(CKCell.Text)}) --> ", Cells[1].Text));

            COMObject.Delete();

            //Log.Debug(LH.TraceCaller(LH.PP.Result, "After delete",
            //    nameof(CKRow), nameof(Delete),
            //    $"{nameof(COMObject)}({nameof(CKCell)}.{nameof(CKCell.RowIndex)}) --> ", Cells[1].RowIndex,
            //    $"{nameof(COMObject)}({nameof(CKCell)}[1].{nameof(CKCell.Text)}) --> ", Cells[1].Text));

        }

        // Selects the row
        public void Select()
        {
            COMObject.Select();
        }


    }
}


