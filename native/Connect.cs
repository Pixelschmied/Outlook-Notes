using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    /// <summary>
    /// The COM add-in entry point. Outlook loads this class via the
    /// HKCU\...\Outlook\Addins registration (no admin, no login). It adds a
    /// ribbon button and a docked task pane hosting <see cref="NotesPane"/>.
    /// </summary>
    [ComVisible(true)]
    [Guid("E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51")]
    [ProgId("EmailNotes.Connect")]
    // AutoDispatch so the ribbon can find callbacks (e.g. OnNotesButton) by name.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private object _app;          // Outlook.Application (late-bound)
        private object _addInInst;
        private object _ctp;          // CustomTaskPane (late-bound)
        private NotesPane _pane;

        // ---- IDTExtensibility2 ----

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            _app = application;
            _addInInst = addInInst;
            Log("OnConnection (mode " + connectMode + ")");
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            Log("OnDisconnection");
            try
            {
                if (_ctp != null) { dynamic c = _ctp; c.Delete(); }
            }
            catch { }
            _ctp = null;
            _pane = null;
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { }

        // ---- ICustomTaskPaneConsumer ----

        public void CTPFactoryAvailable(object ctpFactoryInst)
        {
            Log("CTPFactoryAvailable");
            try
            {
                dynamic factory = ctpFactoryInst;
                _ctp = factory.CreateCTP("EmailNotes.NotesPane", "E-Mail-Notizen", Type.Missing);
                dynamic ctp = _ctp;
                ctp.DockPosition = 2; // msoCTPDockPositionRight
                ctp.Width = 360;
                object content = ctp.ContentControl;
                _pane = content as NotesPane;
                if (_pane != null) _pane.Initialize(_app);
                ctp.Visible = true;
                Log("Task pane created; content=" + (content == null ? "null" : content.GetType().FullName));
            }
            catch (Exception ex) { Log("CTP error: " + ex); }
        }

        // ---- IRibbonExtensibility ----

        public string GetCustomUI(string ribbonID)
        {
            Log("GetCustomUI " + ribbonID);
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("EmailNotes.Ribbon.xml"))
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
            catch (Exception ex) { Log("GetCustomUI error: " + ex); return null; }
        }

        // Ribbon callback (invoked by name via IDispatch).
        [ComVisible(true)]
        public void OnNotesButton(object control)
        {
            try
            {
                if (_ctp == null) return;
                dynamic ctp = _ctp;
                ctp.Visible = !(bool)ctp.Visible;
            }
            catch (Exception ex) { Log("OnNotesButton error: " + ex); }
        }

        // ---- diagnostics (helps debugging on a real machine) ----

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
