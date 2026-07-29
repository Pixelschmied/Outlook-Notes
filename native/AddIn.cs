using System;
using System.IO;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.OutlookApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace OutlookNotes
{
    /// <summary>
    /// Outlook-Notes add-in (NetOffice COMAddin). Adds a ribbon button and a
    /// docked task pane hosting the WinForms notepad. Fully local — no login,
    /// no internet, per-user install.
    /// </summary>
    [COMAddin("Outlook-Notes", "A private notepad docked next to your mail.", LoadBehavior.LoadAtStartup)]
    [ProgId("OutlookNotes.AddIn")]
    [Guid("E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51")]
    [Codebase]
    public class AddIn : COMAddin
    {
        private Office._CustomTaskPane _ctp;
        private NotesPane _pane;

        public AddIn()
        {
            Log("ctor");
            OnStartupComplete += AddIn_OnStartupComplete;
        }

        // Ribbon XML (returned to Outlook). Inline to avoid embedded-resource lookup.
        public override string GetCustomUI(string ribbonID)
        {
            Log("GetCustomUI " + ribbonID);
            string label;
            try { label = Localization.Get(Localization.DetectLanguage(GetRawApp())).PaneTitle; }
            catch { label = "Notes"; }
            label = System.Security.SecurityElement.Escape(label);
            return
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabMail\">" +
                "<group id=\"outlookNotesGroup\" label=\"Outlook-Notes\">" +
                "<button id=\"outlookNotesToggle\" label=\"" + label + "\" showImage=\"false\" size=\"normal\"" +
                " onAction=\"OnNotesButton\" screentip=\"" + label + "\" />" +
                "</group></tab></tabs></ribbon></customUI>";
        }

        /// <summary>The raw Outlook Application COM object (for late-bound calls), or null.</summary>
        private object GetRawApp()
        {
            try { return Application.UnderlyingObject; }
            catch { try { return Application; } catch { return null; } }
        }

        // Ribbon button callback (invoked by name).
        public void OnNotesButton(object control)
        {
            Log("OnNotesButton");
            try
            {
                if (EnsurePane()) _ctp.Visible = !_ctp.Visible;
            }
            catch (Exception ex) { Log("OnNotesButton err: " + ex); }
        }

        private void AddIn_OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
            try
            {
                if (EnsurePane()) _ctp.Visible = true;
            }
            catch (Exception ex) { Log("startup pane err: " + ex); }
        }

        /// <summary>Create the docked task pane on demand (returns true if it exists).</summary>
        private bool EnsurePane()
        {
            if (_ctp != null) return true;
            if (TaskPaneFactory == null) { Log("no TaskPaneFactory yet"); return false; }
            object rawApp = GetRawApp();
            var loc = Localization.Get(Localization.DetectLanguage(rawApp));
            _ctp = TaskPaneFactory.CreateCTP("OutlookNotes.NotesPane", loc.PaneTitle, Type.Missing);
            _ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _ctp.Width = 360;
            object content = _ctp.ContentControl;
            Log("pane content=" + (content == null ? "null" : content.GetType().FullName));
            _pane = content as NotesPane;
            if (_pane != null) _pane.Initialize(rawApp);
            return true;
        }

        internal static void Log(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookNotes");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "addin.log"),
                    DateTime.Now.ToString("s") + "  " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
