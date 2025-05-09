using System;
using System.Collections.Generic;

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

    public abstract class CKDOMObject : IDOMObject
    {
        public CKDocument Document => Parent.Document;

        public CKApplication Application => Parent.Application;

        public abstract IDOMObject Parent { get; protected set; }

        public abstract bool IsDirty { get; protected set; }

        public abstract bool IsOrphan { get; protected set; }
    }


    /// <summary>
    /// Provides a registry for converting one IDOMObject into a specific derived type.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0010
    /// </remarks>
    public static class IDOMCaster
    {
        private static readonly Dictionary<Type, Func<IDOMObject, IDOMObject>> _casters = new Dictionary<Type, Func<IDOMObject, IDOMObject>>();

        /// <summary>
        /// Registers a caster function that creates a specific derived IDOMObject from a general one.
        /// </summary>
        /// <typeparam name="T">The type to cast to.</typeparam>
        /// <param name="caster">A function that accepts an IDOMObject and returns a T.</param>
        public static void Register<T>(Func<IDOMObject, T> caster) where T : IDOMObject
        {
            if (caster == null) throw new ArgumentNullException(nameof(caster));
            _casters[typeof(T)] = source => caster(source);
        }

        /// <summary>
        /// Attempts to convert the input IDOMObject to a specific derived type using a registered caster.
        /// </summary>
        /// <typeparam name="T">The derived type to return.</typeparam>
        /// <param name="input">The source object to convert.</param>
        /// <returns>The converted object of type T.</returns>
        public static T Cast<T>(IDOMObject input) where T : IDOMObject
        {
            if (input == null) throw new ArgumentNullException(nameof(input));
            if (_casters.TryGetValue(typeof(T), out var caster))
            {
                var result = (T)caster(input);

                return result;
            }

            throw new InvalidOperationException($"No caster registered for type {typeof(T).Name}.");
        }
    }
}
