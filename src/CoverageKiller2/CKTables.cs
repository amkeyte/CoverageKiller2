using CoverageKiller2;
using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

public class CKTables : IEnumerable<CKTable>
{
    private readonly CKDocument _document;

    public CKTables(CKDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    private Word.Tables WordTables => _document.WordDoc.Tables;

    public IEnumerator<CKTable> GetEnumerator()
    {
        for (int i = 1; i <= WordTables.Count; i++)
        {
            yield return new CKTable(WordTables[i]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
