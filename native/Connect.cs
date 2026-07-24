using System;
using System.IO;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    /// <summary>
    /// COM add-in entry point. Loaded by Outlook via HKCU\...\Outlook\Addins
    /// (no admin, no login). Kept deliberately minimal at load time: it only
    /// registers a ribbon button. The task pane is created lazily on first click
    /// so that any pane/WinForms problem cannot crash Outlook during startup.
    /// </summary>
    [ComVisible(true)]
    [Guid("E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51")]
    [ProgId("EmailNotes.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private const string RibbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
            "<ribbon><tabs><tab idMso=\"TabMail\">" +
            "<group id=\"emailNotesGroup\" label=\"Email Notes\">" +
            "<button id=\"emailNotesToggle\" label=\"Notes\" showImage=\"false\" size=\"normal\"" +
            " onAction=\"OnNotesButton\" screentip=\"E-Mail-Notizen\"" +
            " supertip=\"Notizblock neben der Mail ein-/ausblenden.\" />" +
            "</group></tab></tabs></ribbon></customUI>";

        private object _app;       // Outlook.Application (late-bound)
        private object _factory;   // ICTPFactory (late-bound)
        private object _ctp;       // CustomTaskPane (late-bound)
        private NotesPane _pane;

        static Connect() { Log("static ctor"); }
        public Connect() { Log("ctor"); }

        // ---- IDTExtensibility2 ----

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection mode=" + connectMode);
            _app = application;
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            Log("OnDisconnection");
            try { if (_ctp != null) { dynamic c = _ctp; c.Delete(); } } catch { }
            _ctp = null; _pane = null; _app = null; _factory = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { }

        // ---- ICustomTaskPaneConsumer ----

        public void CTPFactoryAvailable(object ctpFactoryInst)
        {
            // Only remember the factory here — do NOT build the pane yet.
            Log("CTPFactoryAvailable");
            _factory = ctpFactoryInst;
        }

        // ---- IRibbonExtensibility ----

        public string GetCustomUI(string ribbonID)
        {
            Log("GetCustomUI " + ribbonID);
            return RibbonXml;
        }

        // Ribbon callback (via IDispatch/AutoDispatch).
        [ComVisible(true)]
        public void OnNotesButton(object control)
        {
            Log("OnNotesButton");
            try
            {
                if (!EnsurePane()) return;
                dynamic ctp = _ctp;
                ctp.Visible = !(bool)ctp.Visible;
            }
            catch (Exception ex) { Log("OnNotesButton error: " + ex); }
        }

        /// <summary>Create the task pane on demand. Returns true if it exists.</summary>
        private bool EnsurePane()
        {
            if (_ctp != null) return true;
            if (_factory == null) { Log("EnsurePane: no factory"); return false; }
            try
            {
                dynamic factory = _factory;
                _ctp = factory.CreateCTP("EmailNotes.NotesPane", "E-Mail-Notizen", Type.Missing);
                dynamic ctp = _ctp;
                ctp.DockPosition = 2; // right
                ctp.Width = 360;
                object content = ctp.ContentControl;
                Log("CTP content=" + (content == null ? "null" : content.GetType().FullName));
                _pane = content as NotesPane;
                if (_pane != null) _pane.Initialize(_app);
                return true;
            }
            catch (Exception ex) { Log("EnsurePane error: " + ex); _ctp = null; return false; }
        }

        // ---- diagnostics ----

        internal static void Log(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "EmailNotes");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "addin.log"),
                    DateTime.Now.ToString("s") + "  " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
