using System;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents an application instance. Handles Document operations, creation, opening, show/hide etc... 
    /// Other top level application stuff as needed.
    /// </summary>
    public class CKApplication : IDisposable
    {
        public CKApplication()
        {
            throw new NotImplementedException();
        }
        public CKDocument GetDocument(string fullPath, bool visible)
        {
            throw new NotImplementedException();
        }

        public bool IsAddIn { get { throw new NotImplementedException(); } }

        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~CKApplication()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
