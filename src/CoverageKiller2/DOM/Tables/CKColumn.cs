using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.DOM.Tables
{


    public class CKColumn : IDOMObject
    {
        public CKDocument Document => throw new NotImplementedException();

        public Application Application => throw new NotImplementedException();

        public IDOMObject Parent => throw new NotImplementedException();

        public bool IsDirty => throw new NotImplementedException();

        public bool IsOrphan => throw new NotImplementedException();
    }

    public class CKColumns : IEnumerable<CKColumn>, IDOMObject
    {
        public CKDocument Document => throw new NotImplementedException();

        public Application Application => throw new NotImplementedException();

        public IDOMObject Parent => throw new NotImplementedException();

        public bool IsDirty => throw new NotImplementedException();

        public bool IsOrphan => throw new NotImplementedException();

        public IEnumerator<CKColumn> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
