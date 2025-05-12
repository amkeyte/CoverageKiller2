using System;
using System.Runtime.Serialization;

namespace CoverageKiller2.Logging

{
    /// <summary>
    /// Represents an internal debug exception that signals an unexpected state or logic error
    /// within the CoverageKiller DOM system.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0002
    /// </remarks>
    [Serializable]
    public class CKDebugException : Exception
    {
        //TODO: add CallerMemberName and find a way to use it.
        /// <inheritdoc/>
        public CKDebugException() { }

        /// <inheritdoc/>
        public CKDebugException(string msg) : base(msg) { }

        /// <inheritdoc/>
        public CKDebugException(string msg, Exception innerException)
            : base(msg, innerException) { }

        /// <inheritdoc/>
        protected CKDebugException(SerializationInfo info, StreamingContext context)
                : base(info, context)
        {
        }
    }
}
