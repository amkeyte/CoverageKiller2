using System;
using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.DOM.Tables
{


    public class CKColumn : CKDOMObject
    {
        public override IDOMObject Parent { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsDirty { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsOrphan { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
    }

    public class CKColumns : CKDOMObject, IEnumerable<CKColumn>
    {
        public override IDOMObject Parent { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsDirty { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }
        public override bool IsOrphan { get => throw new NotImplementedException(); protected set => throw new NotImplementedException(); }

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
