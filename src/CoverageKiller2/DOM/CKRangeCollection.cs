using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Diagnostics;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Abstract base class for all collection types based on a Word range.
    /// Provides IsDirty tracking, deferCOM lazy activation, and cache helpers.
    /// </summary>
    public abstract class ACKRangeCollection : IDOMObject
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ACKRangeCollection"/> class.
        /// </summary>
        /// <param name="parent">The parent DOM object.</param>
        /// <param name="deferCOM">If true, defers COM access until first real usage.</param>
        protected ACKRangeCollection(IDOMObject parent, bool deferCOM = false)
        {
            //LH.Debug("Tracker[!sd]");
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            IsCOMDeferred = deferCOM;
            _isDirty = deferCOM;
        }

        /// <summary>
        /// Indicates whether COM access has been deferred for this collection.
        /// Flipped internally during first cache or IsDirty check if needed.
        /// </summary>
        public bool IsCOMDeferred { get; private set; }

        /// <summary>
        /// Tracks dirty state for cache invalidation.
        /// </summary>
        protected bool _isDirty { get; set; }

        private bool _isCheckingDirty = false;
        private bool _isRefreshing;

        /// <inheritdoc/>
        public IDOMObject Parent { get; protected set; }

        /// <inheritdoc/>
        public CKDocument Document => Parent.Document;

        /// <inheritdoc/>
        public CKApplication Application => Parent.Application;

        IDOMObject IDOMObject.Parent => Parent; // explicit interface version

        /// <summary>
        /// Gets the number of items in the collection.
        /// </summary>
        public abstract int Count { get; }

        /// <summary>
        /// Determines whether the collection has lost association with the underlying Word objects.
        /// </summary>
        public abstract bool IsOrphan { get; }

        /// <summary>
        /// Attempts to find the index of an object within the collection.
        /// </summary>
        public abstract int IndexOf(object obj);

        /// <summary>
        /// Clears the collection cache.
        /// </summary>
        public abstract void Clear();

        /// <inheritdoc/>
        public virtual bool IsDirty
        {
            get
            {
                this.Ping(msg: $"Parent: {Parent.GetType()}");

                if (_isDirty || _isCheckingDirty)
                {
                    this.Pong();
                    return _isDirty;
                }

                _isCheckingDirty = true;
                try
                {
                    _isDirty = _isDirty || CheckDirtyFor();
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

                this.Pong(msg: _isDirty.ToString());
                return _isDirty;
            }
            protected set => _isDirty = value;
        }


        /// <summary>
        /// Checks whether the collection has become dirty.
        /// Override in descendants to customize.
        /// </summary>
        /// <returns>True if dirty; otherwise, false.</returns>
        protected virtual bool CheckDirtyFor()
        {
            return this.PingPong(() => false, msg: false.ToString());
        }

        /// <summary>
        /// Provides lazy cache loading for a field, automatically lifting defer if necessary.
        /// </summary>
        /// <typeparam name="T">The field type.</typeparam>
        /// <param name="cachedField">The field to cache.</param>
        /// <returns>The cached value, refreshed if needed.</returns>
        protected T Cache<T>(ref T cachedField)
        {
            if (IsDirty || cachedField == null)
            {
                if (IsCOMDeferred)
                {
                    Log.Debug($"Deferred COM access standard refresh for {GetType().Name}.");
                    IsCOMDeferred = false;
                }

                Refresh();
            }
            return cachedField;
        }

        /// <summary>
        /// Provides lazy cache loading for a field with a custom refresh function.
        /// </summary>
        /// <typeparam name="T">The field type.</typeparam>
        /// <param name="cachedField">The field to cache.</param>
        /// <param name="refreshFunc">A function to refresh the cache.</param>
        /// <returns>The cached value, refreshed if needed.</returns>
        protected T Cache<T>(ref T cachedField, Func<T> refreshFunc)
        {
            if (IsDirty || cachedField == null)
            {
                Log.Debug($"renewing cache: custom call of {typeof(T)} in {GetType().Name}.");
                cachedField = refreshFunc();
                if (IsCOMDeferred)
                {
                    Log.Debug($"Deferred COM access triggered inside Cache<T> (CUSTOM refresh) for {GetType().Name}.");
                    IsCOMDeferred = false;
                }
                Refresh();
            }
            Log.Debug($"retrieved cache with return value {cachedField.ToString()}");
            return cachedField;
        }
        protected void SetCache<T>(ref T field, T value, Action<T> setter = null)
        {
            setter?.Invoke(value);
            field = value;
            IsDirty = true;
        }

        /// <summary>
        /// Forces the collection to refresh its cache from the underlying COM objects.
        /// Descendants should override this to implement their specific refresh logic.
        /// </summary>
        public void Refresh()
        {
            Log.Debug("refreshing!!");
            if (_isRefreshing) return;
            _isRefreshing = true;
            Log.Debug("even more refreshing!");
            DoRefreshThings();
            IsDirty = false;
            _isRefreshing = false;
        }
        protected virtual void DoRefreshThings()
        {
            //no op
        }
    }
}
