using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of <see cref="CKParagraph"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.04.0000
    /// </remarks>
    public class CKParagraphs : ACKRangeCollection, IEnumerable<CKParagraph>
    {
        private List<CKParagraph> _cachedParagraphs;
        private Word.Paragraphs _COMParagraphs;

        protected List<CKParagraph> ParagraphsList
        {
            get => Cache(ref _cachedParagraphs, () =>
            {
                Log.Debug("Pulling Paragraphs list");
                var list = new List<CKParagraph>();

                if (IsCOMDeferred)
                {
                    Log.Debug("Deferred paragraph list build started.");
                    for (int i = 1; i <= (COMParagraphs?.Count ?? 0); i++)
                    {
                        list.Add(null); // Empty slots, paragraphs built on demand
                    }
                    //does not call refresh
                }
                else
                {
                    Log.Debug("Eager paragraph list build started.");
                    for (int i = 1; i <= (COMParagraphs?.Count ?? 0); i++)
                    {
                        list.Add(new CKParagraph(COMParagraphs[i], this, i));
                    }
                }

                return list;
            });

            set => SetCache(ref _cachedParagraphs, value);
        }


        /// <summary>
        /// Gets the underlying Word.Paragraphs COM object from the parent range.
        /// May be null if deferred.
        /// </summary>
        protected Word.Paragraphs COMParagraphs
        {
            get => Cache(ref _COMParagraphs, () =>
                {
                    Log.Debug("getting _COMparagraphs");
                    if (Parent is CKRange parentRange)
                    {
                        return parentRange.COMRange.Paragraphs;
                    }
                    throw new CKDebugException("Unsupported parent type for resolving COMParagraphs.");
                });

            private set => SetCache(ref _COMParagraphs, value);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class with an eager COM collection.
        /// </summary>
        public CKParagraphs(Word.Paragraphs collection, IDOMObject parent, bool deferCOM = false)
            : base(parent, deferCOM)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (collection.Count > 1000)
                throw new ArgumentException($"Paragraphs collection is too large ({collection.Count}). Find a smaller subset.");

            Log.Debug($"CKParagraphs created with {collection.Count} entries (deferCOM={deferCOM}).");

            COMParagraphs = collection;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class with defer mode.
        /// </summary>
        public CKParagraphs(IDOMObject parent, bool deferCOM = true)
            : base(parent, deferCOM)
        {
        }

        /// <inheritdoc/>
        public override int Count => ParagraphsList?.Count ?? 0;

        /// <inheritdoc/>
        [Obsolete]
        public override bool IsOrphan => throw new NotImplementedException();

        //private void InitializeEmptyList(int count)
        //{
        //    _cachedParagraphs = new List<CKParagraph>(count);
        //    for (int i = 0; i < count; i++)
        //    {
        //        _cachedParagraphs.Add(null); // placeholder nulls
        //    }
        //}
        public CKParagraph this[int index]
        {
            get
            {
                Log.Debug("Pulling index: " + index);
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of paragraphs.");

                var list = ParagraphsList;
                var para = list[index - 1];
                if (para == null)
                {
                    para = new CKParagraph(this, index);
                    list[index - 1] = para;
                }
                return para;
            }
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            if (obj is CKRange para)
            {
                for (int i = 0; i < ParagraphsList.Count; i++)
                {
                    if (ParagraphsList[i] != null && ParagraphsList[i].Equals(para))
                        return i + 1; // one-based
                }
            }
            return -1;
        }

        /// <inheritdoc/>
        public IEnumerator<CKParagraph> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                Log.Debug("enumerating " + i);
                yield return this[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <inheritdoc/>
        public override void Clear()
        {
            ParagraphsList?.Clear();
            ParagraphsList = null;
            COMParagraphs = null;
            IsDirty = true;
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"CKParagraphs [Count: {Count}]";
        }

        /// <inheritdoc/>
        protected override void DoRefreshThings()
        {
            Log.Debug("Doing refresh things");
            //ParagraphsList = null;
            //COMParagraphs = null;
            IsDirty = false;
        }
    }
}
