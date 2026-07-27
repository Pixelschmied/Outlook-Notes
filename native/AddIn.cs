using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.Tools;

namespace EmailNotes
{
    /// <summary>
    /// Outlook-Notes add-in, built on NetOffice so the COM add-in plumbing
    /// (IDTExtensibility2, ribbon, task pane) is handled by a proven library
    /// instead of hand-declared interfaces. This first cut only proves it loads:
    /// it logs and shows a confirmation box at startup.
    /// </summary>
    [COMAddin("Outlook-Notes", "Privater Notizblock neben der Mail.", 3)]
    [ProgId("OutlookNotes.AddIn")]
    [Guid("E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51")]
    [RegistryLocation(RegistrySaveLocation.CurrentUser)]
    public class AddIn : COMAddin
    {
        public AddIn()
        {
            Log("ctor");
            OnStartupComplete += AddIn_OnStartupComplete;
        }

        private void AddIn_OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
            try { MessageBox.Show("Outlook-Notes läuft! 🎉"); }
            catch (Exception ex) { Log("msgbox: " + ex.Message); }
        }

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
