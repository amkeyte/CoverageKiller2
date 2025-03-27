using CoverageKiller2.DOM;
using System;
using Word = Microsoft.Office.Interop.Word;

public class CKCell : CKRange
{
    public Word.Cell COMCell { get; }
    public CKTable Table { get; }
    public int WordRow { get; }
    public int WordColumn { get; }

    public CKCell(CKTable table, IDOMObject parent, Word.Cell wdCell, int wordRow, int wordColumn)
        : base(wdCell.Range, parent)
    {
        Table = table ?? throw new ArgumentNullException(nameof(table));
        COMCell = wdCell ?? throw new ArgumentNullException(nameof(wdCell));
        WordRow = wordRow;
        WordColumn = wordColumn;
    }

    public Word.WdColor BackgroundColor
    {
        get => COMCell.Shading.BackgroundPatternColor;
        set => COMCell.Shading.BackgroundPatternColor = value;
    }

    public Word.WdColor ForegroundColor
    {
        get => COMCell.Shading.ForegroundPatternColor;
        set => COMCell.Shading.ForegroundPatternColor = value;
    }
}
