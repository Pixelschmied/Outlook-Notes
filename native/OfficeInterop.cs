using System;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    // Add-in COM interfaces, hand-declared so the project builds WITHOUT any
    // Office PIA. They are declared as DUAL interfaces (not ComImport, not
    // IDispatch-only): the host calls them through the v-table, so the CCW must
    // expose that v-table with the methods in the original order. Getting this
    // wrong causes an access violation / Outlook crash on load.
    //
    // Outlook object-model access itself stays late-bound via `dynamic`.

    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        [DispId(2)]
        void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);
        [DispId(4)]
        void OnStartupComplete(ref Array custom);
        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    [ComVisible(true)]
    [Guid("000C033A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        void CTPFactoryAvailable([MarshalAs(UnmanagedType.IDispatch)] object ctpFactoryInst);
    }
}
