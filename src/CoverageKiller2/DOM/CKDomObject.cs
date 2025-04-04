namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a DOM object in the CoverageKiller2 system.
    /// Provides references to parent structures and document/application context.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0002
    /// </remarks>
    public interface IDOMObject
    {
        /// <summary>
        /// Gets the CKDocument that owns this DOM object.
        /// </summary>
        CKDocument Document { get; }

        /// <summary>
        /// Gets the CKApplication instance that owns the document.
        /// </summary>
        CKApplication Application { get; }

        /// <summary>
        /// Gets the logical parent of this DOM object.
        /// </summary>
        IDOMObject Parent { get; }

        /// <summary>
        /// Indicates whether this DOM object has unsaved changes.
        /// </summary>
        bool IsDirty { get; }

        /// <summary>
        /// Indicates whether this DOM object is orphaned (e.g., lost its backing COM reference).
        /// </summary>
        bool IsOrphan { get; }
    }
}
