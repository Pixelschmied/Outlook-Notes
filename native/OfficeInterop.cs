using System;
using System.Runtime.InteropServices;

namespace EmailNotes
{
    // Add-in COM interfaces, hand-declared so the project builds WITHOUT any
    // Office PIA. Declared exactly like the real primary interop assemblies:
    // IDispatch-based, ComVisible, no ComImport (we IMPLEMENT them, the CCW
    // exposes them). Outlook object-model access stays late-bound via `dynamic`.

    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
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
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }

    [ComVisible(true)]
    [Guid("000C033A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICustomTaskPaneConsumer
    {
        void CTPFactoryAvailable([MarshalAs(UnmanagedType.IDispatch)] object ctpFactoryInst);
    }
}
