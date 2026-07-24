using System;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    // Minimal, hand-declared Office/add-in COM interfaces so the project builds
    // WITHOUT any Office PIA or Office installed on the build machine. All access
    // to the Outlook object model itself is late-bound via `dynamic`.

    /// <summary>The classic add-in entry point interface (Extensibility.dll).</summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        void OnDisconnection(int removeMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    /// <summary>Provides the ribbon XML (Office.IRibbonExtensibility).</summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    /// <summary>Lets the host hand us the task-pane factory (Office.ICustomTaskPaneConsumer).</summary>
    [ComImport]
    [Guid("000C033A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICustomTaskPaneConsumer
    {
        void CTPFactoryAvailable([MarshalAs(UnmanagedType.IDispatch)] object ctpFactoryInst);
    }
}
