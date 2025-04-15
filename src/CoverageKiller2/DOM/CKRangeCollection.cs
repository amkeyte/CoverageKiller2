using System;

namespace CoverageKiller2.DOM
{
    public abstract class ACKRangeCollection : IDOMObject
    {
        protected ACKRangeCollection(IDOMObject parent)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        public abstract int IndexOf(object obj);



        /// <summary>
        /// Gets the number of sections in the range.
        /// </summary>
        public abstract int Count { get; }
        protected bool _isDirty { get; set; }
        public abstract bool IsDirty { get; protected set; }

        public IDOMObject Parent { get; }
        public CKDocument Document => Parent.Document;

        public CKApplication Application => Parent.Application;

        IDOMObject IDOMObject.Parent => Parent;

        public abstract bool IsOrphan { get; }
    }
}
