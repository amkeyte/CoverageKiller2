using CoverageKiller2.Logging;
using System;
using System.Diagnostics;

namespace CoverageKiller2.DOM
{
    public abstract class ACKRangeCollection : IDOMObject
    {
        protected ACKRangeCollection(IDOMObject parent)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        public abstract int IndexOf(object obj);

        public abstract void Clear();

        /// <summary>
        /// Gets the number of sections in the range.
        /// </summary>
        public abstract int Count { get; }
        protected bool _isDirty { get; set; }
        //public abstract bool IsDirty { get; protected set; }

        public IDOMObject Parent { get; protected set; }
        public CKDocument Document => Parent.Document;

        public CKApplication Application => Parent.Application;

        IDOMObject IDOMObject.Parent => Parent; //leave me here!!!

        public abstract bool IsOrphan { get; }

        private bool _isCheckingDirty = false;

        /// <inheritdoc/>

        public virtual bool IsDirty
        {
            get
            {
                this.Ping($"Parent: {Parent.GetType()}");

                if (_isDirty || _isCheckingDirty)
                {
                    this.Pong();
                    return _isDirty;
                }

                _isCheckingDirty = true;
                try
                {
                    _isDirty = _isDirty
                    || CheckDirtyFor();
                    //|| Parent.IsDirty;

                }
                catch (Exception ex)
                {
                    if (Debugger.IsAttached)
                    {
                        Debugger.Break();
                        throw ex;
                    }
                }
                finally
                {

                    _isCheckingDirty = false;
                }

                this.Pong();
                return _isDirty;
            }
            protected set => _isDirty = value;
        }

        protected virtual bool CheckDirtyFor()
        {
            return false;
        }
    }
}
