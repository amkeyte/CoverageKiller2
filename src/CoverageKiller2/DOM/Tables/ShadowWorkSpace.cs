using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using Serilog;
using System;

/// <summary>
/// Provides a hidden, disposable workspace for safe content processing using encapsulation.
/// Wraps a CKDocument inside a private context and uses suppressed alerts.
/// </summary>
/// <remarks>
/// Version: CK2.00.01.0011
/// </remarks>
public class ShadowWorkspace : IDOMObject, IDisposable
{
    private readonly CKApplication _app;
    private readonly CKDocument _doc;
    private readonly bool _keepOpen;

    internal ShadowWorkspace(CKDocument doc, CKApplication app, bool keepOpen)
    {
        this.Ping(msg: "$$$");
        _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        _app = app ?? throw new ArgumentNullException(nameof(app));
        _keepOpen = keepOpen;
        _doc.Activate();//see what happens.
        this.Pong();
    }

    /// <summary>
    /// Gets the internal CKDocument managed by this shadow wrapper.
    /// </summary>
    public CKDocument Document => _doc;

    /// <summary>
    /// Makes the shadow document visible to the user.
    /// </summary>
    public void ShowDebuggerWindow()
    {
        Log.Information($" Activating shadow document {_doc.FileName}");
        try
        {
            _app.Visible = true;
            _doc.Visible = true;
            _doc.Activate();
        }
        catch
        {
            throw new InvalidOperationException("Could not show shadow document.");
        }
    }

    /// <summary>
    /// Hides the shadow document window, if visible.
    /// </summary>
    public void HideDebuggerWindow()
    {
        try
        {
            _doc.Visible = false;
        }
        catch
        {
            // No-op; safe to fail silently
        }
    }



    /// <summary>
    /// Clones a CKRange-based object into the specified CKRange target location.
    /// </summary>
    /// <typeparam name="T">The type to return, such as CKParagraph.</typeparam>
    /// <param name="objToClone">The object to clone.</param>
    /// <param name="cloneToTarget">Target insertion range in this document.</param>
    /// <returns>A new object of type T wrapping the cloned content.</returns>
    /// <remarks>
    /// Version: CK2.00.01.0012
    /// </remarks>
    public T CloneFrom<T>(T objToClone, CKRange cloneToTarget) where T : IDOMObject
    {
        this.Ping(new[] { typeof(T) });

        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));
        if (cloneToTarget == null) throw new ArgumentNullException(nameof(cloneToTarget));

        if (cloneToTarget.Document != this.Document) throw new InvalidOperationException("Cannot clone to other documents");
        if (!(objToClone is CKRange sourceRange))
            throw new NotSupportedException($"CloneFrom<T> only supports CKRange-based objects. Type was {typeof(T).Name}.");

        return _app.WithSuppressedAlerts(() => //why? I forget.
        {
            var insertionPoint = cloneToTarget.CollapseToStart();

            // Perform the paste at the insertion point
            insertionPoint.FormattedText = sourceRange;

            // Determine how much was inserted
            int insertStart = insertionPoint.Start;
            int insertEnd = insertionPoint.End;

            var result = IDOMCaster.Cast<T>(_doc.Range(insertStart, insertEnd));

            this.Pong(new[] { typeof(T) });
            return result;
        });
    }

    /// <summary>
    /// Clones a CKRange-based object and inserts it at the end of this document.
    /// </summary>
    /// <typeparam name="T">The type to return, such as CKParagraph.</typeparam>
    /// <param name="objToClone">The object to clone.</param>
    /// <returns>A new object of type T wrapping the cloned content.</returns>
    /// <remarks>
    /// Version: CK2.00.01.0013
    /// </remarks>
    public T CloneFrom<T>(T objToClone) where T : IDOMObject
    {
        this.Pong(new[] { typeof(T) });

        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));

        var result = CloneFrom(objToClone, _doc.Content.CollapseToEnd());

        LH.Pong(GetType(), new Type[] { typeof(T) });
        return result;
    }

    /// <summary>
    /// Clones a CKRange-based object to a specified location in this document.
    /// </summary>
    /// <typeparam name="T">The type to return, such as CKParagraph.</typeparam>
    /// <param name="objToClone">The object to clone.</param>
    /// <param name="start">Start position of the target range.</param>
    /// <param name="end">End position of the target range.</param>
    /// <returns>A new object of type T wrapping the cloned content.</returns>
    /// <remarks>
    /// Version: CK2.00.01.0010
    /// </remarks>
    public T CloneFrom<T>(T objToClone, int start, int end) where T : IDOMObject
    {
        this.Ping(new[] { typeof(T) });

        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));
        if (start < 0 || end < start) throw new ArgumentOutOfRangeException();

        var range = _doc.Range(start, end);
        var result = CloneFrom(objToClone, range);

        this.Pong(new[] { typeof(T) });
        return result;
    }

    /// <summary>
    /// Disposes and optionally closes the document.
    /// it never gets called.
    /// </summary>
    public void Dispose()
    {
        if (!_keepOpen)
        {
            try
            {
                _app.CloseDocument(_doc, force: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ShadowWorkspace.Dispose] Error: {ex.Message}");
            }
        }
    }

    public CKApplication Application => _app;
    public IDOMObject Parent => _doc;
    public bool IsDirty => _doc.IsDirty;
    public bool IsOrphan => _doc.IsOrphan;
    public string LogId => _doc.LogId;
}
