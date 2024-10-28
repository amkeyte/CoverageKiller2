using CoverageKiller2.Logging;
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
            Tracer.Enabled = false;
            Parent = parent;
            COMObject = Parent.COMObject[index];
        }
        public CKCells Cells => CKCells.Create(this);
        //public bool ContainsMerged => Cells.ContainsMerged;
        // Property to access row's index
        public int Index => Tracer.Trace(COMObject.Index);

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

        public Tracer Tracer = new Tracer(typeof(CKRow));

        // Deletes the row
        public void Delete()
        {
            Tracer.Log("Deleting Row", new DataPoints(nameof(Index)));

            COMObject.Delete();
        }

        // Selects the row
        public void Select()
        {
            COMObject.Select();
        }


    }
}


