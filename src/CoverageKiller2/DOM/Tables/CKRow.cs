using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.DOM.Tables
{

    public class CKRow : CKRange
    {
        public CKRow(Range range, IDOMObject parent = null) : base(range, parent)
        {
        }
    }

    /// <summary>
    /// Represents a collection of CKRow objects in a Word table.
    /// This collection is part of the DOM hierarchy and implements IDOMObject.
    /// </summary>
    public class CKRows : CKDOMObject, IEnumerable<CKRow>
    {
        public override IDOMObject Parent { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsDirty { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsOrphan { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }

        public IEnumerator<CKRow> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
