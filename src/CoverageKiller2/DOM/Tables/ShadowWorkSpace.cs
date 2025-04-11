using CoverageKiller2.DOM;
using System;
using Word = Microsoft.Office.Interop.Word;

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
        _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        _app = app ?? throw new ArgumentNullException(nameof(app));
        _keepOpen = keepOpen;
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
    /// Clones a table from any Word document into this shadow document.
    /// </summary>
    /// <param name="source">The table to clone.</param>
    /// <returns>The pasted table in this document.</returns>
    public Word.Table CloneTable(Word.Table source)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));

        _app.WithSuppressedAlerts(() =>
        {
            _doc.Content.Delete();
            _doc.Content.COMRange.FormattedText = source.Range.FormattedText;
        });

        return _doc.Tables[1].COMTable;
    }

    /// <summary>
    /// Clones a CKRange-derived object to the end of this document.
    /// </summary>
    /// <typeparam name="T">The type of CKRange to return.</typeparam>
    /// <param name="objToClone">The CKRange object to clone.</param>
    /// <returns>The cloned CKRange instance in this document.</returns>
    public T CloneRange<T>(T objToClone) where T : CKRange
    {
        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));

        return _app.WithSuppressedAlerts(() =>
        {
            var insertAt = _doc.Range().End - 1;
            var targetRange = _doc.Range(insertAt, insertAt);
            targetRange.FormattedText = objToClone.FormattedText;
            var resultRange = _doc.Range(insertAt, insertAt + targetRange.Text.Length);
            return (T)Activator.CreateInstance(typeof(T), new object[] { resultRange.COMRange, _doc });
        });
    }

    /// <summary>
    /// Clones a CKRange-based object into the specified range of a target document,
    /// and returns a strongly typed wrapper using a registered IDOMCaster.
    /// </summary>
    /// <typeparam name="T">The type to return, such as CKParagraph.</typeparam>
    /// <param name="objToClone">The object to clone (must be CKRange-based).</param>
    /// <param name="cloneToTarget">The destination range in this document.</param>
    /// <returns>A new object of type T wrapping the cloned content.</returns>
    /// <remarks>
    /// Version: CK2.00.01.0008
    /// </remarks>
    public T CloneFrom<T>(T objToClone, CKRange cloneToTarget) where T : IDOMObject
    {
        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));
        if (cloneToTarget == null) throw new ArgumentNullException(nameof(cloneToTarget));

        return _app.WithSuppressedAlerts(() =>
        {
            if (!(objToClone is CKRange sourceRange))
                throw new NotSupportedException($"CloneFrom<T> only supports CKRange-based objects. Type was {typeof(T).Name}.");

            cloneToTarget.COMRange.FormattedText = sourceRange.FormattedText.COMRange;

            var resultRange = _doc.Range(cloneToTarget.Start, cloneToTarget.Start + sourceRange.COMRange.Text.Length);
            var wrapped = new CKRange(resultRange.COMRange, _doc);

            return IDOMCaster.Cast<T>(wrapped);
        });
    }

    /// <summary>
    /// Clones a CKRange-based object to the end of this document.
    /// </summary>
    /// <typeparam name="T">The type to return, such as CKParagraph.</typeparam>
    /// <param name="objToClone">The object to clone.</param>
    /// <returns>A new object of type T wrapping the cloned content at the end of the document.</returns>
    /// <remarks>
    /// Version: CK2.00.01.0009
    /// </remarks>
    public T CloneFrom<T>(T objToClone) where T : IDOMObject
    {
        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));

        var insertAt = _doc.Range().End - 1;
        var targetRange = _doc.Range(insertAt, insertAt);
        return CloneFrom(objToClone, new CKRange(targetRange.COMRange, _doc));
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
        if (objToClone == null) throw new ArgumentNullException(nameof(objToClone));
        if (start < 0 || end < start) throw new ArgumentOutOfRangeException();

        var range = _doc.Range(start, end);
        return CloneFrom(objToClone, new CKRange(range.COMRange, _doc));
    }

    /// <summary>
    /// Disposes and optionally closes the document.
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
