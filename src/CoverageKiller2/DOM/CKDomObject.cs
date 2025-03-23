using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    //for multiple inheritance
    public interface IDOMObject
    {
        CKDocument Document { get; }
        Word.Application Application { get; }

        IDOMObject Parent { get; }

        bool IsDirty { get; }

        bool IsOrphan { get; }
    }
}
