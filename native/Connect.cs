using System;
using System.IO;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    /// <summary>
    /// ISOLATION BUILD: the add-in implements ONLY IDTExtensibility2 and does
    /// nothing but log. No ribbon, no task pane, no WinForms, no AutoDispatch.
    /// If Outlook now reaches OnConnection/OnStartupComplete (visible in the
    /// log), the foundation is sound and the earlier crash was in the ribbon /
    /// task-pane interfaces — which we then add back correctly.
    /// </summary>
    [ComVisible(true)]
    [Guid("E7A9C1F4-3B2D-4A6E-8C1B-9D0E2F3A4B51")]
    [ProgId("EmailNotes.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        static Connect() { Log("static ctor"); }
        public Connect() { Log("ctor"); }

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection mode=" + connectMode);
        }

        public void OnDisconnection(int removeMode, ref Array custom) { Log("OnDisconnection"); }
        public void OnAddInsUpdate(ref Array custom) { Log("OnAddInsUpdate"); }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { Log("OnBeginShutdown"); }

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
