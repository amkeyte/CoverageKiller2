using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Root class of the DOM, handles system startup, shutdown and ensuring that COM objects are
    /// handled to prevent leaks.
    /// </summary>
    public class CKOffice_Word : IDisposable
    {
        public bool IsRunning { get; private set; }
        private static int Start(Word.Application initialApplication)
        {
            throw new NotImplementedException();
        }
        private bool disposedValue;
        public static int ShutDown()
        {
            throw new NotSupportedException();
        }


        /// <summary>
        /// This is to act as a placeholder in the Applications set when an application gets shut down.
        /// The purpose is to maintain index integrity for the lifetime of CKOffice operation.
        /// It will generally throw exceptions or act dead in whatever appropriate way during access attmpts.
        /// </summary>
        private class CKAlreadyBeenClosedApp : CKApplication
        {
            public CKAlreadyBeenClosedApp()
            {
                throw new NotImplementedException();
            }
        }
        //1 based list -> make sure to add the tombstone at zero.Use an aleadyDed to get the desired outcome.
        public static IEnumerable<CKApplication> Applications { get; } = new HashSet<CKApplication>();
        public static int TryGetNewApp(out CKApplication app)
        {
            throw new NotImplementedException();
        }
        public static int TryPutAddin(ThisAddIn addin)
        {
            throw new NotImplementedException();
        }

        protected virtual void Dispose(bool disposing)
        {
            throw new NotImplementedException();
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
        // ~CKOffice_Word()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
