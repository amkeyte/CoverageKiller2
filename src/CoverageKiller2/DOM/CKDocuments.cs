using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of CKDocument objects.
    /// This is implemented as a singleton so that there is one central collection per session.
    /// </summary>
    public class CKDocuments : IEnumerable<CKDocument>
    {
        // Singleton instance
        private static readonly CKDocuments _instance = new CKDocuments();

        /// <summary>
        /// Gets the singleton instance of CKDocuments.
        /// </summary>
        public static CKDocuments GetInstance()
        {
            return _instance;
        }

        // The internal list of CKDocument objects.
        private List<CKDocument> _documents = new List<CKDocument>();

        // Private constructor prevents external instantiation.
        private CKDocuments() { }

        /// <summary>
        /// Gets the number of documents in the collection.
        /// </summary>
        public int Count => _documents.Count;

        /// <summary>
        /// Adds a new CKDocument to the collection.
        /// Throws an exception if a document with the same FullName already exists.
        /// </summary>
        /// <param name="document">The CKDocument to add.</param>
        /// <returns>The added CKDocument.</returns>
        internal CKDocument Add(CKDocument document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // Use FullName as the unique identifier.
            if (_documents.Any(d => d.FullPath.Equals(document.FullPath, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("A document with the same FullName already exists in the collection.");

            _documents.Add(document);
            return document;
        }

        /// <summary>
        /// Removes the specified CKDocument from the collection.
        /// </summary>
        /// <param name="document">The document to remove.</param>
        internal void Remove(CKDocument document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            _documents.Remove(document);
        }

        /// <summary>
        /// Returns an enumerator that iterates through the CKDocument collection.
        /// </summary>
        public IEnumerator<CKDocument> GetEnumerator()
        {
            return _documents.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Gets the CKDocument at the specified zero-based index.
        /// </summary>
        public CKDocument this[int index] => _documents[index];

        /// <summary>
        /// Retrieves a CKDocument by its full name.
        /// Returns null if no document with that full name exists.
        /// </summary>
        /// <param name="fullName">The full name (path) of the document.</param>
        internal static CKDocument GetByName(string fullName)
        {
            if (string.IsNullOrEmpty(fullName))
                throw new ArgumentNullException(nameof(fullName));

            return _instance._documents
                .FirstOrDefault(d => d.FullPath.Equals(fullName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Retrieves a CKDocument that wraps the specified COM Document.
        /// If one does not already exist in the collection, it is created and added.
        /// </summary>
        /// <param name="document">The COM Document.</param>
        /// <returns>A CKDocument that wraps the specified COM Document.</returns>
        internal static CKDocument GetByCOMDocument(Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            CKDocument ckDoc = GetByName(document.FullName);
            if (ckDoc == null)
            {
                ckDoc = new CKDocument(document);
                GetInstance().Add(ckDoc);
            }
            return ckDoc;
        }
    }
}
