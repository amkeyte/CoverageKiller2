using System;

namespace CoverageKiller2
{
    public abstract class ACKRangeCollection
    {
        protected ACKRangeCollection(CKRange parent)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        /// <summary>
        /// Gets the parent <see cref="CKRange"/> associated with this instance.
        /// </summary>
        public CKRange Parent { get; protected set; }

        /// <summary>
        /// Gets the number of sections in the range.
        /// </summary>
        public abstract int Count { get; }
        protected bool _isDirty { get; set; }
        public abstract bool IsDirty { get; }
    }
}
