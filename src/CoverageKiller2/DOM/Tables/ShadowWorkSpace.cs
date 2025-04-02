using CoverageKiller2.DOM;
using System;
using Word = Microsoft.Office.Interop.Word;

public class ShadowWorkspace : IDisposable
{
    private readonly Word.Application _app;
    private CKDocument _doc;
    private bool _visible;

    public CKDocument Document => _doc;

    public ShadowWorkspace(Word.Application application)
    {
        _app = application ?? throw new ArgumentNullException(nameof(application));
    }

    private bool _keepOpen = false;

    /// <summary>
    /// Call to make the shadow document visible and persist after test cleanup.
    /// </summary>
    public void ShowDebuggerWindow(bool keepOpen = false)
    {
        _visible = true;
        _keepOpen = keepOpen;

        if (_doc != null)
        {
            _doc.COMDocument.Windows[1].Visible = true;
            _doc.COMDocument.Activate();
        }
        else
        {
            _app.Visible = true;
        }
    }
    public Word.Table CloneTable(Word.Table source)
    {
        EnsureShadowDocument();

        _doc.COMDocument.Content.Delete();
        source.Range.Copy();
        _doc.COMDocument.Content.Paste();

        return _doc.COMDocument.Tables[1];
    }

    private void EnsureShadowDocument()
    {
        if (_doc != null) return;

        var wordDoc = _app.Documents.Add(Visible: _visible);
        _doc = new CKDocument(wordDoc);
    }

    public void Dispose()
    {
        if (_keepOpen) return;

        _doc?.Close(false);
        _doc?.Dispose();
        _doc = null;
    }
}
